import io
import json
import re
import sqlite3
import tempfile
import unittest
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate
from werkzeug.security import generate_password_hash

from app import (
    PageNumberCanvas,
    create_app,
    get_db,
    import_machine_workbook,
    init_db,
    log_user_activity,
)


def get_static_css(client):
    response = client.get("/static/css/app.css")
    try:
        return response.get_data(as_text=True)
    finally:
        response.close()


class RepairReportAppTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "test.sqlite3"
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(self.database_path),
                "UPLOAD_DIR": str(Path(self.temp_dir.name) / "uploads"),
                "SECRET_KEY": "test",
            }
        )
        self.client = self.app.test_client()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_pdf_page_numbers_include_total_pages(self):
        page_numbers = []

        class RecordingPageNumberCanvas(PageNumberCanvas):
            def drawCentredString(self, x, y, text):
                page_numbers.append(text)
                super().drawCentredString(x, y, text)

        output = io.BytesIO()
        paragraph_style = getSampleStyleSheet()["Normal"]
        SimpleDocTemplate(output, pagesize=A4).build(
            [
                Paragraph("Pierwsza strona", paragraph_style),
                PageBreak(),
                Paragraph("Druga strona", paragraph_style),
            ],
            canvasmaker=RecordingPageNumberCanvas,
        )

        self.assertEqual(page_numbers, ["1/2", "2/2"])

    def test_record_pdf_export_has_data_page_and_image_page(self):
        from PIL import Image

        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-11-04T09:01",
                "repaired_at": "2026-11-04T10:00",
                "machine_name": "Prasa / 1",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-001",
                "event_type": "Awaria",
                "description": "Błąd komunikacji",
                "labor_cost": "2.5",
                "material_cost": "1234.5",
                "spare_parts_cost": "0",
                "service_cost": "0",
            },
        )
        with self.app.app_context():
            record_id = int(get_db().execute("SELECT id FROM repair_records").fetchone()["id"])

        def count_pages(pdf_bytes):
            return len(re.findall(rb"/Type /Page(?!s)", pdf_bytes))

        without_images = self.client.get(f"/record/{record_id}/pdf")
        self.assertEqual(without_images.status_code, 200)
        self.assertEqual(without_images.mimetype, "application/pdf")
        self.assertTrue(without_images.data.startswith(b"%PDF"))
        self.assertEqual(count_pages(without_images.data), 1)
        self.assertIn(f"rekord_{record_id}_Prasa_1.pdf", without_images.headers["Content-Disposition"])

        image_bytes = io.BytesIO()
        Image.new("RGB", (40, 30), "white").save(image_bytes, format="JPEG")
        self.client.post(
            f"/api/record/{record_id}/attachments",
            data={"attachments": (io.BytesIO(image_bytes.getvalue()), "foto.jpg", "image/jpeg")},
        )
        with_images = self.client.get(f"/record/{record_id}/pdf")
        self.assertEqual(with_images.status_code, 200)
        self.assertEqual(count_pages(with_images.data), 2)

        self.assertEqual(self.client.get("/record/9999/pdf").status_code, 404)

    def test_create_record_calculates_downtime_and_cost(self):
        response = self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "repaired_at": "2026-08-30T10:30",
                "machine_name": "Prasa 1",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-001",
                "event_type": "Awaria",
                "description": "Wymiana czujnika",
                "cause": "Zużycie",
                "repair_method": "Wymiana",
                "replaced_parts": "Czujnik",
                "labor_cost": "1",
                "material_cost": "20",
                "spare_parts_cost": "100",
                "service_cost": "0",
            },
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn("Rekord został zapisany.", page)
        self.assertIn("2.5", page)
        self.assertIn("200.0", page)

    def test_copy_record_creates_new_record_with_submitted_occurrence(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "repaired_at": "2026-08-30T10:30",
                "machine_name": "Prasa 1",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-001",
                "event_type": "Awaria",
                "description": "Wymiana czujnika",
                "cause": "Zużycie",
                "repair_method": "Wymiana",
                "replaced_parts": "Czujnik",
                "participants": "Jan",
                "labor_cost": "1",
                "material_cost": "20",
                "spare_parts_cost": "100",
                "service_cost": "0",
                "notes": "Test",
            },
        )
        with self.app.app_context():
            original = get_db().execute("SELECT id FROM repair_records").fetchone()
            original_id = int(original["id"])

        response = self.client.post(
            f"/api/record/{original_id}/copy",
            json={
                "occurrence_at": "2026-11-11T14:25",
                "repaired_at": "",
                "machine_name": "Prasa 1",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-001",
                "event_type": "Awaria",
                "description": "Wymiana czujnika",
                "cause": "Zużycie",
                "repair_method": "Wymiana",
                "replaced_parts": "Czujnik",
                "participants": "Jan",
                "labor_cost": "1",
                "material_cost": "20",
                "spare_parts_cost": "100",
                "service_cost": "0",
                "notes": "Test",
            },
        )

        self.assertEqual(response.status_code, 200)
        copied_id = response.get_json()["id"]
        self.assertNotEqual(copied_id, original_id)
        with self.app.app_context():
            db = get_db()
            original = db.execute("SELECT occurrence_at FROM repair_records WHERE id = ?", (original_id,)).fetchone()
            copied = db.execute("SELECT occurrence_at, repaired_at, description FROM repair_records WHERE id = ?", (copied_id,)).fetchone()
            self.assertEqual(original["occurrence_at"], "2026-08-30 08:00")
            self.assertEqual(copied["occurrence_at"], "2026-11-11 14:25")
            self.assertIsNone(copied["repaired_at"])
            self.assertEqual(copied["description"], "Wymiana czujnika")

    def test_guest_can_view_but_cannot_export_or_create_records(self):
        app = create_app(
            {
                "TESTING": False,
                "REQUIRE_LOGIN": False,
                "DATABASE": str(self.database_path),
                "UPLOAD_DIR": str(Path(self.temp_dir.name) / "uploads"),
                "SECRET_KEY": "test",
            }
        )
        client = app.test_client()

        self.assertEqual(client.get("/").status_code, 200)
        data_page = client.get("/data")
        self.assertEqual(data_page.status_code, 200)
        data_page_html = data_page.get_data(as_text=True)
        self.assertNotIn("Eksportuj całą bazę", data_page_html)
        self.assertNotIn("Import / export bazy maszyn", data_page_html)

        filter_page = client.get("/filter?description_terms=test")
        self.assertEqual(filter_page.status_code, 200)
        self.assertNotIn("Eksportuj wyniki (.xlsx)", filter_page.get_data(as_text=True))
        self.assertNotIn("eksportuj całą bazę", filter_page.get_data(as_text=True))

        export_response = client.get("/export?description_terms=test")
        self.assertEqual(export_response.status_code, 302)
        denied_page = client.get(export_response.headers["Location"], follow_redirects=True)
        self.assertIn(
            "Konto gościa nie ma uprawnień do eksportu danych.",
            denied_page.get_data(as_text=True),
        )

        response = client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Maszyna gościa",
                "event_type": "Awaria",
                "description": "Nieautoryzowany wpis",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("Brak uprawnień do dodawania rekordów.", response.get_data(as_text=True))

        with app.app_context():
            record_count = get_db().execute("SELECT COUNT(*) FROM repair_records").fetchone()[0]
        self.assertEqual(record_count, 0)

    def test_polish_date_format_is_displayed_and_accepted(self):
        response = self.client.post(
            "/",
            data={
                "occurrence_at": "08-10-2026 14:35",
                "repaired_at": "08-10-2026 15:05",
                "machine_name": "Maszyna dat",
                "machine_type": "Testowa",
                "serial_number": "DATE-001",
                "event_type": "Awaria",
                "description": "Test formatu daty",
            },
        )
        self.assertEqual(response.status_code, 302)

        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("08-10-2026\n14:35", page)
        for path in ("/new", "/", "/filter"):
            with self.subTest(path=path):
                page = self.client.get(path).get_data(as_text=True)
                self.assertIn('type="datetime-local"', page)
                self.assertIn("syncDisplayDateTime(this)", page)
                self.assertIn("syncNativeDateTime(this)", page)
                self.assertIn('onclick="openDateTimePicker(this)"', page)
                self.assertIn("picker.showPicker()", page)

        with self.app.app_context():
            record_id = get_db().execute(
                """
                SELECT r.id
                FROM repair_records r
                JOIN machines m ON m.id = r.machine_id
                WHERE m.serial_number = ?
                """,
                ("DATE-001",),
            ).fetchone()
            self.assertIsNotNone(record_id)
            record_id = int(record_id["id"])

        record = self.client.get(f"/api/record/{record_id}").get_json()
        self.assertEqual(record["occurrence_at"], "2026-10-08 14:35")
        history = self.client.get(
            "/api/machines?machine=Maszyna%20dat&machine_type=Testowa&serial_number=DATE-001"
        ).get_json()
        self.assertEqual(history["previous_entries"][0]["occurrence_at"], "08-10-2026 14:35")

        filtered_page = self.client.get("/filter?occurrence_from=08-10-2026").get_data(as_text=True)
        self.assertIn("Test formatu daty", filtered_page)
        self.assertIn('value="08-10-2026"', filtered_page)

    def test_machine_api_and_filtered_export(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "repaired_at": "2026-08-30T09:00",
                "machine_name": "Linia A",
                "machine_type": "Pakowarka",
                "serial_number": "P-100",
                "event_type": "Przegląd",
                "description": "Kontrola napędu",
            },
        )
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-31T08:00",
                "repaired_at": "2026-08-31T09:00",
                "machine_name": "Linia B",
                "machine_type": "Pakowarka",
                "serial_number": "P-200",
                "event_type": "Awaria",
                "description": "Naprawa silnika",
                "notes": "Kontrola łożyska",
                "participants": "Jan Kowalski, Anna Nowak",
            },
        )

        api_response = self.client.get("/api/machines?machine=Linia%20A&machine_type=Pakowarka")
        payload = api_response.get_json()
        self.assertEqual(payload["types"], ["Pakowarka"])
        self.assertEqual(payload["serial_numbers"], ["P-100"])
        self.assertEqual(len(payload["previous_entries"]), 1)

        export_response = self.client.get("/export?machine=Linia%20B&description_terms=silnika")
        self.assertEqual(export_response.status_code, 200)
        workbook = load_workbook(io.BytesIO(export_response.data))
        rows = list(workbook.active.iter_rows(values_only=True))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][4], "Linia B")
        self.assertEqual(rows[1][13], "31-08-2026 08:00")
        self.assertEqual(rows[1][14], "31-08-2026 09:00")
        self.assertEqual(rows[0][12], "Osoby biorące udział w interwencji")
        self.assertEqual(rows[1][12], "Jan Kowalski, Anna Nowak")

        filter_response = self.client.get("/filter?description_terms=łożyska")
        filter_page = filter_response.get_data(as_text=True)
        self.assertIn("Linia B", filter_page)
        self.assertIn("31-08-2026\n08:00", filter_page)

        notes_export_response = self.client.get("/export?description_terms=łożyska")
        self.assertEqual(notes_export_response.status_code, 200)
        notes_workbook = load_workbook(io.BytesIO(notes_export_response.data))
        notes_rows = list(notes_workbook.active.iter_rows(values_only=True))
        self.assertEqual(len(notes_rows), 2)
        self.assertEqual(notes_rows[1][4], "Linia B")

    def test_task_status_filter_includes_completed_and_incomplete_records(self):
        for index, repaired_at in enumerate(("2026-08-30T09:00", "")):
            self.client.post(
                "/",
                data={
                    "occurrence_at": f"2026-08-30T08:0{index}",
                    "repaired_at": repaired_at,
                    "machine_name": f"Maszyna {index}",
                    "machine_type": "Testowa",
                    "serial_number": f"T-{index}",
                    "event_type": "Awaria",
                    "description": f"Zadanie {index}",
                },
            )

        completed_page = self.client.get("/filter?task_status=completed").get_data(as_text=True)
        self.assertIn("Zadanie 0", completed_page)
        self.assertNotIn("Zadanie 1", completed_page)
        self.assertIn('<option value="completed" selected>', completed_page)

        incomplete_page = self.client.get("/filter?task_status=incomplete").get_data(as_text=True)
        self.assertIn("Zadanie 1", incomplete_page)
        self.assertNotIn("Zadanie 0", incomplete_page)
        self.assertIn('<option value="incomplete" selected>', incomplete_page)

        export_response = self.client.get("/export?task_status=incomplete")
        self.assertEqual(export_response.status_code, 200)
        workbook = load_workbook(io.BytesIO(export_response.data))
        rows = list(workbook.active.iter_rows(values_only=True))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][8], "Zadanie 1")

    def test_scheduled_event_form_uses_existing_machine_lookup(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Prasa A",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-101",
                "event_type": "Awaria",
                "description": "Testowa naprawa",
            },
        )

        page = self.client.get("/scheduled/new")
        self.assertEqual(page.status_code, 200)
        html = page.get_data(as_text=True)
        css = get_static_css(self.client)
        self.assertIn('id="scheduled-machine-list"', html)
        self.assertIn('id="scheduled-machine-type-list"', html)
        self.assertIn('id="scheduled-serial-number-list"', html)
        self.assertIn('value="Prasa A"', html)
        self.assertIn("fetch(`/api/machines?${params}`)", html)
        self.assertIn('id="planned_at" name="planned_at" type="text"', html)
        self.assertIn('placeholder="dd-mm-rrrr gg:mm"', html)
        self.assertIn('type="datetime-local" class="datetime-picker-input"', html)
        prefilled_page = self.client.get("/scheduled/new?planned_at=2026-05-14T10:00")
        self.assertIn('value="14-05-2026 10:00"', prefilled_page.get_data(as_text=True))
        self.assertIn("grid-template-columns: repeat(4, minmax(0, 1fr))", css)
        self.assertIn('class="field scheduled-third-row-start"', html)
        self.assertIn('class="field scheduled-notes-field"', html)
        self.assertIn('id="repeat-interval-field" hidden', html)
        self.assertIn(".scheduled-form-grid > .field[hidden] { display: none; }", css)
        self.assertIn("repeatUnitInput.addEventListener('change', updateRepeatIntervalVisibility)", html)

        lookup = self.client.get(
            "/api/machines?machine=Prasa%20A&machine_type=Hydrauliczna"
        ).get_json()
        self.assertEqual(lookup["types"], ["Hydrauliczna"])
        self.assertEqual(lookup["serial_numbers"], ["H-101"])

    def test_scheduled_event_calendar_month_and_year_views(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        events = (
            ("Przegląd", "2026-01-05T07:00", "Prasa Blue", "Blue calendar event"),
            ("Interwencja", "2026-05-14T10:30", "Prasa Yellow", "Yellow calendar event"),
            ("Modernizacja", "2026-12-01T12:15", "Prasa Purple", "Purple calendar event"),
        )
        for event_type, planned_at, machine_name, description in events:
            response = self.client.post(
                "/scheduled/new",
                data={
                    "machine_name": machine_name,
                    "machine_type": "Hydrauliczna",
                    "serial_number": f"{machine_name}-1",
                    "event_type": event_type,
                    "planned_at": planned_at,
                    "repeat_interval": "0",
                    "repeat_unit": "none",
                    "description": description,
                },
            )
            self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            completed_event_id = get_db().execute(
                "SELECT id FROM scheduled_events WHERE description = ?",
                ("Blue calendar event",),
            ).fetchone()["id"]
        response = self.client.post(f"/scheduled/{completed_event_id}/done")
        self.assertEqual(response.status_code, 302)

        year_page = self.client.get("/scheduled?calendar_view=year&year=2026")
        css = get_static_css(self.client)
        self.assertEqual(year_page.status_code, 200)
        year_html = year_page.get_data(as_text=True)
        self.assertIn('data-calendar-view="year"', year_html)
        self.assertEqual(year_html.count('class="scheduled-events-table"'), 2)
        self.assertEqual(year_html.count('<col class="u-width-10">'), 2)
        self.assertIn("table-layout: fixed;", css)
        self.assertEqual(year_html.count('class="calendar-month-heading"'), 12)
        self.assertIn('class="calendar-week-label">Tydz.</div>', year_html)
        self.assertRegex(year_html, r'class="calendar-week-number">\s*1\s*</div>')
        self.assertIn('class="calendar-dot badge-przeglad calendar-scheduled-event calendar-scheduled-done"', year_html)
        self.assertIn('.calendar-year-view .calendar-dot.calendar-scheduled-done::before { content: "✓"; }', css)
        self.assertIn('class="calendar-dot badge-interwencja calendar-scheduled-event"', year_html)
        self.assertIn('class="calendar-dot badge-modernizacja calendar-scheduled-event"', year_html)
        self.assertIn(".calendar-year-view .calendar-dot {", css)
        self.assertIn("width: .85rem;", css)
        self.assertIn(".calendar-year-view .calendar-dot.badge-interwencja { background: var(--yellow); }", css)
        self.assertIn(".calendar-year-view .calendar-dot.badge-modernizacja { background: var(--purple); }", css)

        month_page = self.client.get("/scheduled?calendar_view=month&year=2026&month=5")
        self.assertEqual(month_page.status_code, 200)
        month_html = month_page.get_data(as_text=True)
        self.assertIn('data-calendar-view="month"', month_html)
        self.assertRegex(month_html, r'class="calendar-week-number">\s*18\s*</div>')
        self.assertIn("title=\"14-05-2026 10:30 — Prasa Yellow — Interwencja — Yellow calendar event\"", month_html)
        self.assertIn('class="calendar-event badge badge-interwencja calendar-scheduled-event"', month_html)
        self.assertIn('class="badge calendar-status-done">✓ Wykonano</span>', month_html)
        self.assertIn('data-machine="Prasa Yellow"', month_html)
        self.assertIn('data-description="Yellow calendar event"', month_html)
        self.assertIn('role="button" tabindex="0" aria-label="Otwórz szczegóły:', month_html)
        self.assertIn("event.stopPropagation();\n        openScheduledModal(calendarEvent);", month_html)
        self.assertNotIn("title=\"05-01-2026 07:00 — Prasa Blue", month_html)
        self.assertNotIn("title=\"01-12-2026 12:15 — Prasa Purple", month_html)

        completed_month_html = self.client.get(
            "/scheduled?calendar_view=month&year=2026&month=1"
        ).get_data(as_text=True)
        self.assertIn(
            'class="calendar-event badge badge-przeglad calendar-scheduled-event calendar-scheduled-done"',
            completed_month_html,
        )
        self.assertIn(">✓ Przegląd · Prasa Blue</span>", completed_month_html)
        self.assertIn('data-status="done"', completed_month_html)

        today = datetime.now().date()
        current_year_html = self.client.get(
            f"/scheduled?calendar_view=year&year={today.year}"
        ).get_data(as_text=True)
        current_month_html = self.client.get(
            f"/scheduled?calendar_view=month&year={today.year}&month={today.month}"
        ).get_data(as_text=True)
        self.assertIn('aria-current="date"', current_year_html)
        self.assertIn("calendar-day-today", current_year_html)
        self.assertRegex(
            current_year_html,
            rf'href="/(?:scheduled/new|planowanie/nowe)\?planned_at={today.isoformat()}T08:00"',
        )
        self.assertIn('aria-current="date"', current_month_html)
        self.assertIn("calendar-day-today", current_month_html)
        self.assertRegex(
            month_html,
            r'href="/(?:scheduled/new|planowanie/nowe)\?planned_at=2026-05-14T08:00"',
        )

        week_page = self.client.get("/scheduled?calendar_view=week")
        self.assertEqual(week_page.status_code, 200)
        week_html = week_page.get_data(as_text=True)
        self.assertIn('data-calendar-view="week"', week_html)
        self.assertEqual(week_html.count('class="calendar-week-grid calendar-hour-row"'), 24)
        self.assertIn("23:00", week_html)
        self.assertIn("aria-current=\"date\"", week_html)
        self.assertIn("calendar-day-today", week_html)

        target_week = datetime.fromisocalendar(2026, 20, 1)
        selected_week = self.client.get("/scheduled?calendar_view=week&iso_week=2026-W20")
        self.assertEqual(selected_week.status_code, 200)
        selected_week_html = selected_week.get_data(as_text=True)
        css = get_static_css(self.client)
        self.assertIn(target_week.strftime("%d.%m"), selected_week_html)
        self.assertIn("10:30 · Interwencja · Prasa Yellow", selected_week_html)
        self.assertIn('class="calendar-event badge badge-interwencja calendar-scheduled-event"', selected_week_html)
        self.assertIn('id="calendar-iso-week" type="week" name="iso_week" value="2026-W20"', selected_week_html)
        self.assertIn('.calendar-controls input[type="week"] { width: 9.5rem; }', css)
        self.assertRegex(
            selected_week_html,
            r'href="/(?:scheduled/new|planowanie/nowe)\?planned_at=2026-05-14T10:00"',
        )
        self.assertRegex(
            selected_week_html,
            r'href="/(?:scheduled|planowanie)\?calendar_view=week&amp;iso_week=2026-W19"',
        )
        completed_week = self.client.get("/scheduled?calendar_view=week&iso_week=2026-W02")
        self.assertEqual(completed_week.status_code, 200)
        completed_week_html = completed_week.get_data(as_text=True)
        self.assertIn("✓ 07:00 · Przegląd · Prasa Blue", completed_week_html)
        self.assertIn("WYKONANO — 05-01-2026 07:00 — Prasa Blue", completed_week_html)

    def test_records_calendar_shows_events_and_only_day_view_creates_records(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        for occurrence_at, description in (
            ("2026-03-02T09:30", "Calendar record one"),
            ("2026-04-15T12:15", "Calendar record two"),
        ):
            response = self.client.post(
                "/",
                data={
                    "occurrence_at": occurrence_at,
                    "machine_name": "Prasa calendar",
                    "machine_type": "Hydrauliczna",
                    "serial_number": "CAL-001",
                    "event_type": "Awaria",
                    "description": description,
                },
            )
            self.assertEqual(response.status_code, 302)

        home_page = self.client.get("/")
        self.assertIn(
            'href="/records-calendar?calendar_view=day"',
            home_page.get_data(as_text=True),
        )
        default_calendar = self.client.get("/records-calendar").get_data(as_text=True)
        self.assertIn('data-calendar-view="day"', default_calendar)
        with self.app.app_context():
            record_id = get_db().execute(
                "SELECT id FROM repair_records WHERE description = ?",
                ("Calendar record one",),
            ).fetchone()["id"]

        year_html = self.client.get(
            "/records-calendar?calendar_view=year&year=2026"
        ).get_data(as_text=True)
        self.assertIn("Calendar record one", year_html)
        self.assertIn("Calendar record two", year_html)
        self.assertIn('data-calendar-view="year"', year_html)

        month_html = self.client.get(
            "/records-calendar?calendar_view=month&year=2026&month=3"
        ).get_data(as_text=True)
        self.assertIn("02-03-2026 09:30 — Prasa calendar — Awaria — Calendar record one", month_html)
        self.assertNotIn("Calendar record two", month_html)
        self.assertNotIn("/new?occurrence_at=", month_html)
        self.assertIn(f'data-record-id="{record_id}"', month_html)
        record_details = self.client.get(f"/api/record/{record_id}").get_json()
        self.assertEqual(record_details["description"], "Calendar record one")
        image_upload = self.client.post(
            f"/api/record/{record_id}/attachments",
            data={"attachments": (io.BytesIO(b"calendar image"), "calendar-photo.jpg", "image/jpeg")},
            content_type="multipart/form-data",
        )
        self.assertEqual(image_upload.status_code, 200)
        record_details = self.client.get(f"/api/record/{record_id}").get_json()
        self.assertEqual(record_details["attachments"][0]["content_type"], "image/jpeg")

        week_html = self.client.get(
            "/records-calendar?calendar_view=week&iso_week=2026-W10"
        ).get_data(as_text=True)
        self.assertIn("09:30 · Awaria · Prasa calendar", week_html)
        self.assertNotIn("/new?occurrence_at=", week_html)

        day_html = self.client.get(
            "/records-calendar?calendar_view=day&calendar_date=2026-03-02"
        ).get_data(as_text=True)
        self.assertIn('data-calendar-view="day"', day_html)
        self.assertIn('id="calendar-record-delete"', day_html)
        self.assertIn('class="btn btn-outline btn-danger btn-sm u-display-none" type="button" id="calendar-record-delete"', day_html)
        self.assertIn('id="calendar-record-edit"', day_html)
        self.assertIn('id="calendar-record-pdf"', day_html)
        self.assertIn("elements.edit.classList.toggle('u-display-none'", day_html)
        self.assertIn("elements.delete.classList.toggle('u-display-none'", day_html)
        self.assertIn("elements.pdf.classList.toggle('u-display-none'", day_html)
        self.assertIn('id="calendar-edit-file-input"', day_html)
        self.assertIn("attachment.content_type?.startsWith('image/')", day_html)
        self.assertIn("calendar-edit-attachment-zone", day_html)
        self.assertIn(
            'data-create-url="/new?occurrence_at=2026-03-02T09:30&amp;return_calendar_date=2026-03-02"',
            day_html,
        )
        self.assertIn("09:30 · Awaria · Prasa calendar", day_html)
        self.assertIn(
            'class="calendar-event-description">Calendar record one</span>',
            day_html,
        )
        self.assertIn("Prasa calendar · Hydrauliczna", day_html)
        self.assertNotIn("CAL-001", day_html)
        self.assertEqual(day_html.count('class="calendar-day-slot calendar-create-slot"'), 48)
        self.assertIn(
            'href="/records-calendar?calendar_view=day&amp;calendar_date=2026-03-01"',
            day_html,
        )
        self.assertIn(
            'href="/records-calendar?calendar_view=day&amp;calendar_date=2026-03-03"',
            day_html,
        )

        new_record_html = self.client.get(
            "/new?occurrence_at=2026-03-02T09:30"
        ).get_data(as_text=True)
        self.assertIn('name="occurrence_at" value="02-03-2026 09:30"', new_record_html)

    def test_new_record_returns_to_calendar_day_when_created_from_day_view(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        form_page = self.client.get(
            "/new?occurrence_at=2026-03-02T09:30&return_calendar_date=2026-03-02"
        )
        self.assertIn(
            'name="return_calendar_date" value="2026-03-02"',
            form_page.get_data(as_text=True),
        )
        response = self.client.post(
            "/new",
            data={
                "return_calendar_date": "2026-03-02",
                "occurrence_at": "02-03-2026 09:30",
                "machine_name": "Prasa return",
                "machine_type": "Hydrauliczna",
                "serial_number": "RET-001",
                "event_type": "Awaria",
                "description": "Powrót do kalendarza dziennego",
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            response.location,
            "/records-calendar?calendar_view=day&calendar_date=2026-03-02",
        )

    def test_scheduled_event_form_fields_and_dashboard_counts(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Prasa B",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-202",
                "event_type": "Przegląd",
                "planned_at": "2020-12-15T13:45",
                "repeat_interval": "0",
                "repeat_unit": "none",
                "description": "Coroczny przegląd",
                "cause": "Termin okresowej kontroli",
                "repair_method": "Sprawdzenie układu hydraulicznego",
                "replaced_parts": "Filtr oleju",
                "notes": "Przygotować części",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertIn(response.location, {"/scheduled", "/planowanie"})

        with self.app.app_context():
            event = get_db().execute(
                """
                SELECT description, cause, repair_method, replaced_parts
                FROM scheduled_events
                WHERE description = ?
                """,
                ("Coroczny przegląd",),
            ).fetchone()
        self.assertEqual(
            tuple(event),
            (
                "Coroczny przegląd",
                "Termin okresowej kontroli",
                "Sprawdzenie układu hydraulicznego",
                "Filtr oleju",
            ),
        )
        with self.app.app_context():
            planned_at = get_db().execute(
                "SELECT planned_at FROM scheduled_events WHERE description = ?",
                ("Coroczny przegląd",),
            ).fetchone()["planned_at"]
        self.assertEqual(planned_at, "2020-12-15 13:45")

        page = self.client.get("/scheduled")
        html = page.get_data(as_text=True)
        css = get_static_css(self.client)
        self.assertIn("Zaplanuj nowe zdarzenie", html)
        self.assertIn("15-12-2020 13:45", html)
        form_page = self.client.get("/scheduled/new").get_data(as_text=True)
        self.assertIn('id="scheduled-machine-name"', form_page)
        self.assertIn('placeholder="dd-mm-rrrr gg:mm"', form_page)
        self.assertNotIn('<option value="Awaria">', form_page)
        self.assertIn('<option value="Przegląd">', form_page)
        home_page = self.client.get("/").get_data(as_text=True)
        self.assertIn("na 15-12-2020 13:45.", home_page)
        self.assertIn('id="kpi-weekly-failures"', home_page)
        self.assertIn('id="kpi-downtime"', home_page)
        self.assertIn('id="kpi-cost"', home_page)
        self.assertIn('id="kpi-unrepaired"', home_page)
        self.assertIn('class="scheduled-summary" aria-label="Podsumowanie zdarzeń"', html)
        self.assertIn('class="u-color-var-accent">1</strong>', html)
        self.assertIn('class="u-color-var-green">0</strong>', html)
        self.assertIn(".scheduled-summary-item {", css)
        self.assertIn('class="record-row scheduled-event-row" tabindex="0"', html)
        self.assertIn('data-status="planned"', html)
        self.assertIn('class="badge badge-przeglad">Przegląd</span>', html)
        self.assertRegex(html, r'action="/(?:scheduled|planowanie)/\d+/delete"')
        self.assertIn(">Usuń</button>", html)
        self.assertNotIn(">Szczegóły</button>", html)
        self.assertIn('id="scheduled-modal-delete-form"', html)
        self.assertIn('id="scheduled-modal-complete-form" method="post" class="u-display-none"', html)
        self.assertIn('id="scheduled-modal-undo-form" method="post" class="u-display-none"', html)
        self.assertIn('class="btn btn-success btn-sm">Wykonano</button>', html)
        self.assertIn('class="btn btn-outline btn-sm">Cofnij Wykonanie</button>', html)
        self.assertIn("scheduledModalCompleteForm.classList.toggle('u-display-none', status !== 'planned');", html)
        self.assertIn("scheduledModalUndoForm.classList.toggle('u-display-none', status !== 'done');", html)
        self.assertEqual(
            html.count("onsubmit=\"return confirm('Czy na pewno usunąć to zaplanowane zdarzenie?"),
            2,
        )
        self.assertIn('scheduledModalDeleteForm.action = `/scheduled/${id}/delete`', html)
        self.assertIn("row.addEventListener('click', (event) =>", html)
        self.assertIn("event.target.closest('button, form, a')", html)
        self.assertIn('id="scheduled-modal-edit-button" type="button"', html)
        self.assertIn('id="scheduled-modal-edit-form" method="post" class="u-display-none"', html)
        self.assertIn('id="edit-repeat-interval-field" hidden', html)
        self.assertIn(".scheduled-edit-grid > .field[hidden] { display: none; }", css)
        self.assertIn("editRepeatUnitInput.addEventListener('change', updateEditRepeatIntervalVisibility)", html)
        self.assertIn('name="cause"', html)
        self.assertRegex(css, r"\.scheduled-edit-description,\s+\.scheduled-edit-notes \{ grid-column: 1 / -1; \}")
        self.assertIn(".scheduled-edit-third { grid-column: span 4; }", css)
        self.assertIn('class="field scheduled-edit-description"', html)
        self.assertIn('class="field scheduled-edit-notes"', html)

        with self.app.app_context():
            event_id = get_db().execute(
                "SELECT id FROM scheduled_events WHERE description = ?",
                ("Coroczny przegląd",),
            ).fetchone()["id"]
        edit_response = self.client.post(
            f"/scheduled/{event_id}/edit",
            data={
                "machine_name": "Prasa B",
                "machine_type": "Hydrauliczna",
                "serial_number": "H-202",
                "event_type": "Przegląd",
                "planned_at": "2020-12-16T14:30",
                "repeat_interval": "0",
                "repeat_unit": "none",
                "description": "Przegląd zaktualizowany",
                "cause": "Nowa przyczyna",
                "repair_method": "Nowy sposób",
                "replaced_parts": "Nowy filtr",
                "notes": "Nowe uwagi",
            },
        )
        self.assertEqual(edit_response.status_code, 302)
        with self.app.app_context():
            edited_event = get_db().execute(
                "SELECT description, cause, planned_at FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()
        self.assertEqual(tuple(edited_event), ("Przegląd zaktualizowany", "Nowa przyczyna", "2020-12-16 14:30"))
        completed_response = self.client.post(f"/scheduled/{event_id}/done")
        self.assertEqual(completed_response.status_code, 302)
        main_page = self.client.get("/").get_data(as_text=True)
        self.assertIn("Przegląd zaktualizowany", main_page)
        with self.app.app_context():
            linked_record = get_db().execute(
                """
                SELECT r.event_type, r.description, r.cause, r.repair_method,
                       r.replaced_parts, r.occurrence_at, r.repaired_at,
                       r.machine_id, se.machine_id AS scheduled_machine_id,
                       se.repair_record_id, r.id AS record_id
                FROM scheduled_events se
                JOIN repair_records r ON r.id = se.repair_record_id
                WHERE se.id = ?
                """,
                (event_id,),
            ).fetchone()
            record_count = get_db().execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Przegląd zaktualizowany",),
            ).fetchone()[0]
        self.assertEqual(linked_record["event_type"], "Przegląd")
        self.assertEqual(linked_record["description"], "Przegląd zaktualizowany")
        self.assertEqual(linked_record["cause"], "Nowa przyczyna")
        self.assertEqual(linked_record["repair_method"], "Nowy sposób")
        self.assertEqual(linked_record["replaced_parts"], "Nowy filtr")
        self.assertEqual(linked_record["occurrence_at"], "2020-12-16 14:30")
        self.assertIsNotNone(linked_record["repaired_at"])
        self.assertEqual(linked_record["machine_id"], linked_record["scheduled_machine_id"])
        self.assertEqual(linked_record["repair_record_id"], linked_record["record_id"])
        self.assertEqual(record_count, 1)

        self.client.post(f"/scheduled/{event_id}/done")
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                    ("Przegląd zaktualizowany",),
                ).fetchone()[0],
                1,
            )
        completed_page = self.client.get("/scheduled").get_data(as_text=True)
        self.assertIn('class="u-color-var-accent">0</strong>', completed_page)
        self.assertIn('class="u-color-var-green">1</strong>', completed_page)
        self.assertIn('data-default-sort-column="3" data-default-sort-direction="desc"', completed_page)
        self.assertIn('class="scheduled-sort-button">Powtarzanie', completed_page)
        self.assertIn('class="scheduled-sort-button">Typ / numer', completed_page)
        self.assertIn("sortScheduledTable(table, defaultColumn, defaultDirection);", completed_page)
        self.assertIn("<td class=\"text-muted\">Hydrauliczna<br>H-202</td>", completed_page)
        self.assertIn('class="badge badge-przeglad">Przegląd</span>', completed_page)
        self.assertIn("<td class=\"text-muted\">Brak</td>", completed_page)
        self.assertIn('class="record-row scheduled-event-row" tabindex="0"', completed_page)
        self.assertIn('data-status="done"', completed_page)
        self.assertIn(">Cofnij</button>", completed_page)
        self.assertRegex(completed_page, r'action="/(?:scheduled|planowanie)/\d+/undo"')
        self.assertRegex(completed_page, r'action="/(?:scheduled|planowanie)/\d+/delete"')
        self.assertNotIn(">Szczegóły</button>", completed_page)
        self.assertEqual(
            completed_page.count("onsubmit=\"return confirm('Czy na pewno usunąć to zaplanowane zdarzenie?"),
            2,
        )

    def test_personal_scheduled_event_is_private_and_never_added_to_register(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "user"
            session["user_id"] = 101

        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Prasa prywatna",
                "machine_type": "Hydrauliczna",
                "serial_number": "OWN-101",
                "event_type": "Własne",
                "planned_at": "2020-11-05T09:00",
                "description": "Prywatny przegląd",
                "repeat_interval": "0",
                "repeat_unit": "none",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            event = get_db().execute(
                "SELECT id, created_by_user_id FROM scheduled_events WHERE description = ?",
                ("Prywatny przegląd",),
            ).fetchone()
        event_id = int(event["id"])
        self.assertEqual(event["created_by_user_id"], 101)

        form_html = self.client.get("/scheduled/new").get_data(as_text=True)
        self.assertIn('<option value="Własne">Własne</option>', form_html)
        owner_page = self.client.get("/scheduled").get_data(as_text=True)
        self.assertIn("Prywatny przegląd", owner_page)
        self.assertIn('class="badge badge-wlasne">Własne</span>', owner_page)
        css = get_static_css(self.client)
        self.assertIn(".badge-wlasne       { background: var(--pink-dim);", css)
        self.assertEqual(
            self.client.get("/api/scheduled-events").get_json()[0]["description"],
            "Prywatny przegląd",
        )

        with self.client.session_transaction() as session:
            session["user_id"] = 202
        other_user_page = self.client.get("/scheduled").get_data(as_text=True)
        self.assertNotIn("Prywatny przegląd", other_user_page)
        self.assertNotIn("Prywatny przegląd", self.client.get("/").get_data(as_text=True))
        self.assertEqual(self.client.get("/api/scheduled-events").get_json(), [])
        self.assertEqual(self.client.get(f"/api/scheduled-events/{event_id}").status_code, 404)
        self.assertEqual(
            self.client.patch(
                f"/api/scheduled-events/{event_id}",
                json={"description": "Nieuprawniona zmiana"},
            ).status_code,
            404,
        )
        self.assertEqual(
            self.client.delete(f"/api/scheduled-events/{event_id}").status_code,
            404,
        )
        self.client.post(f"/scheduled/{event_id}/done")
        self.client.post(
            f"/scheduled/{event_id}/edit",
            data={"description": "Nieuprawniona zmiana"},
        )
        self.client.post(f"/scheduled/{event_id}/delete")
        with self.app.app_context():
            unchanged = get_db().execute(
                "SELECT status, description FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()
        self.assertEqual(tuple(unchanged), ("planned", "Prywatny przegląd"))

        with self.client.session_transaction() as session:
            session["user_id"] = 101
        completed = self.client.post(f"/scheduled/{event_id}/done")
        self.assertEqual(completed.status_code, 302)
        with self.app.app_context():
            db = get_db()
            completed_event = db.execute(
                "SELECT status, repair_record_id FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()
            record_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Prywatny przegląd",),
            ).fetchone()[0]
            init_db()
            materialized_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Prywatny przegląd",),
            ).fetchone()[0]
        self.assertEqual(completed_event["status"], "done")
        self.assertIsNone(completed_event["repair_record_id"])
        self.assertEqual(record_count, 0)
        self.assertEqual(materialized_count, 0)

    def test_recurring_personal_scheduled_event_keeps_owner_and_skips_register(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "user"
            session["user_id"] = 303
        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Prasa cykliczna prywatna",
                "machine_type": "Hydrauliczna",
                "serial_number": "OWN-303",
                "event_type": "Własne",
                "planned_at": "2020-11-05T09:30",
                "repeat_interval": "2",
                "repeat_unit": "weeks",
                "description": "Prywatny przegląd cykliczny",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            event = get_db().execute(
                "SELECT id, planned_at FROM scheduled_events WHERE description = ?",
                ("Prywatny przegląd cykliczny",),
            ).fetchone()

        response = self.client.post(
            f"/scheduled/{event['id']}/done",
            data={"expected_planned_at": event["planned_at"]},
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            rows = get_db().execute(
                """
                SELECT status, planned_at, repair_record_id, created_by_user_id
                FROM scheduled_events
                WHERE description = ?
                ORDER BY status
                """,
                ("Prywatny przegląd cykliczny",),
            ).fetchall()
            record_count = get_db().execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Prywatny przegląd cykliczny",),
            ).fetchone()[0]
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["status"], "done")
        self.assertIsNone(rows[0]["repair_record_id"])
        self.assertEqual(rows[0]["created_by_user_id"], 303)
        self.assertEqual(rows[1]["status"], "planned")
        self.assertEqual(rows[1]["planned_at"], "2020-11-19 09:30")
        self.assertEqual(rows[1]["created_by_user_id"], 303)
        self.assertEqual(record_count, 0)

    def test_undoing_scheduled_completion_removes_generated_record(self):
        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Tokarka cofnięcie",
                "machine_type": "CNC",
                "serial_number": "UNDO-1",
                "event_type": "Przegląd",
                "planned_at": "2026-03-10T09:00",
                "repeat_interval": "0",
                "repeat_unit": "none",
                "description": "Przegląd do cofnięcia",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            event = get_db().execute(
                "SELECT id, planned_at FROM scheduled_events WHERE description = ?",
                ("Przegląd do cofnięcia",),
            ).fetchone()
        event_id = int(event["id"])
        self.client.post(f"/scheduled/{event_id}/done")
        with self.app.app_context():
            record_id = get_db().execute(
                "SELECT repair_record_id FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()["repair_record_id"]
        self.assertIsNotNone(record_id)
        response = self.client.post(
            f"/api/record/{record_id}/attachments",
            data={"attachments": (io.BytesIO(b"undo attachment"), "undo.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        record_upload_dir = Path(self.app.config["UPLOAD_DIR"]) / str(record_id)
        self.assertTrue(record_upload_dir.exists())

        response = self.client.post(
            f"/scheduled/{event_id}/undo",
            data={"expected_planned_at": event["planned_at"]},
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            undone_event = db.execute(
                "SELECT status, planned_at, repair_record_id FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()
            record_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE id = ?",
                (record_id,),
            ).fetchone()[0]
        self.assertEqual(undone_event["status"], "planned")
        self.assertEqual(undone_event["planned_at"], event["planned_at"])
        self.assertIsNone(undone_event["repair_record_id"])
        self.assertEqual(record_count, 0)
        self.assertFalse(record_upload_dir.exists())

    def test_undoing_recurring_completions_rolls_back_in_reverse_order(self):
        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Tokarka cykliczna cofnięcie",
                "machine_type": "CNC",
                "serial_number": "UNDO-2",
                "event_type": "Przegląd",
                "planned_at": "2026-03-10T09:00",
                "repeat_interval": "2",
                "repeat_unit": "weeks",
                "description": "Przegląd cykliczny do cofnięcia",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            root = db.execute(
                "SELECT id, planned_at FROM scheduled_events WHERE description = ?",
                ("Przegląd cykliczny do cofnięcia",),
            ).fetchone()
        root_id = int(root["id"])
        first_date = str(root["planned_at"])

        self.client.post(
            f"/scheduled/{root_id}/done",
            data={"expected_planned_at": first_date},
        )
        with self.app.app_context():
            db = get_db()
            first = db.execute(
                "SELECT id, planned_at FROM scheduled_events WHERE status = 'done' AND description = ?",
                ("Przegląd cykliczny do cofnięcia",),
            ).fetchone()
            second_date = db.execute(
                "SELECT planned_at FROM scheduled_events WHERE id = ?",
                (root_id,),
            ).fetchone()["planned_at"]

        self.client.post(
            f"/scheduled/{root_id}/done",
            data={"expected_planned_at": second_date},
        )
        with self.app.app_context():
            db = get_db()
            second = db.execute(
                """
                SELECT id, planned_at FROM scheduled_events
                WHERE status = 'done' AND id != ? AND description = ?
                """,
                (first["id"], "Przegląd cykliczny do cofnięcia"),
            ).fetchone()
            self.assertEqual(
                db.execute(
                    "SELECT recurrence_next_event_id FROM scheduled_events WHERE id = ?",
                    (first["id"],),
                ).fetchone()["recurrence_next_event_id"],
                second["id"],
            )

        response = self.client.post(
            f"/scheduled/{second['id']}/undo",
            data={"expected_planned_at": second["planned_at"]},
        )
        self.assertEqual(response.status_code, 302)
        response = self.client.post(
            f"/scheduled/{first['id']}/undo",
            data={"expected_planned_at": first["planned_at"]},
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            rows = db.execute(
                "SELECT status, planned_at, repair_record_id FROM scheduled_events WHERE description = ?",
                ("Przegląd cykliczny do cofnięcia",),
            ).fetchall()
            record_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Przegląd cykliczny do cofnięcia",),
            ).fetchone()[0]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["status"], "planned")
        self.assertEqual(rows[0]["planned_at"], first_date)
        self.assertIsNone(rows[0]["repair_record_id"])
        self.assertEqual(record_count, 0)

    def test_undo_rejects_modified_next_recurring_event(self):
        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Tokarka cykliczna zmieniona",
                "machine_type": "CNC",
                "serial_number": "UNDO-3",
                "event_type": "Przegląd",
                "planned_at": "2026-03-10T09:00",
                "repeat_interval": "2",
                "repeat_unit": "weeks",
                "description": "Przegląd z edytowanym terminem",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            event = db.execute(
                "SELECT id, planned_at FROM scheduled_events WHERE description = ?",
                ("Przegląd z edytowanym terminem",),
            ).fetchone()
        event_id = int(event["id"])
        self.client.post(
            f"/scheduled/{event_id}/done",
            data={"expected_planned_at": event["planned_at"]},
        )
        with self.app.app_context():
            db = get_db()
            completed = db.execute(
                """
                SELECT id, repair_record_id FROM scheduled_events
                WHERE status = 'done' AND description = ?
                """,
                ("Przegląd z edytowanym terminem",),
            ).fetchone()
            db.execute(
                "UPDATE scheduled_events SET description = 'Zmieniony następny termin' WHERE id = ?",
                (event_id,),
            )
            db.commit()

        response = self.client.post(
            f"/scheduled/{completed['id']}/undo",
            data={"expected_planned_at": event["planned_at"]},
            follow_redirects=True,
        )
        self.assertIn("Następny termin tego zdarzenia został zmieniony", response.get_data(as_text=True))
        with self.app.app_context():
            db = get_db()
            completed_status = db.execute(
                "SELECT status FROM scheduled_events WHERE id = ?",
                (completed["id"],),
            ).fetchone()["status"]
            record_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE id = ?",
                (completed["repair_record_id"],),
            ).fetchone()[0]
            next_description = db.execute(
                "SELECT description FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()["description"]
        self.assertEqual(completed_status, "done")
        self.assertEqual(record_count, 1)
        self.assertEqual(next_description, "Zmieniony następny termin")

    def test_existing_completed_scheduled_event_is_added_to_register_once(self):
        response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Tokarka",
                "machine_type": "CNC",
                "serial_number": "T-303",
                "event_type": "Przegląd",
                "planned_at": "2020-11-05T09:00",
                "description": "Stary wykonany przegląd",
                "cause": "Przegląd okresowy",
            },
        )
        self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            db = get_db()
            db.execute(
                "UPDATE scheduled_events SET status = 'done' WHERE description = ?",
                ("Stary wykonany przegląd",),
            )
            db.commit()
            init_db()
            init_db()
            linked = db.execute(
                """
                SELECT se.repair_record_id, r.description, r.cause
                FROM scheduled_events se
                JOIN repair_records r ON r.id = se.repair_record_id
                WHERE se.description = ?
                """,
                ("Stary wykonany przegląd",),
            ).fetchone()
            count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Stary wykonany przegląd",),
            ).fetchone()[0]

        self.assertIsNotNone(linked["repair_record_id"])
        self.assertEqual(linked["description"], "Stary wykonany przegląd")
        self.assertEqual(linked["cause"], "Przegląd okresowy")
        self.assertEqual(count, 1)
        self.assertIn("Stary wykonany przegląd", self.client.get("/").get_data(as_text=True))

    def test_completing_recurring_event_keeps_next_occurrence_planned(self):
        create_response = self.client.post(
            "/scheduled/new",
            data={
                "machine_name": "Prasa cykliczna",
                "machine_type": "Hydrauliczna",
                "serial_number": "CYK-1",
                "event_type": "Przegląd",
                "planned_at": "2020-11-05T09:30",
                "repeat_interval": "2",
                "repeat_unit": "weeks",
                "description": "Dwutygodniowy przegląd",
                "cause": "Kontrola cykliczna",
                "repair_method": "Sprawdzenie układu",
                "replaced_parts": "Filtr",
                "notes": "Uwagi cykliczne",
            },
        )
        self.assertEqual(create_response.status_code, 302)
        with self.app.app_context():
            event = get_db().execute(
                "SELECT id, planned_at FROM scheduled_events WHERE description = ?",
                ("Dwutygodniowy przegląd",),
            ).fetchone()
        event_id = int(event["id"])
        original_planned_at = str(event["planned_at"])

        response = self.client.post(
            f"/scheduled/{event_id}/done",
            data={"expected_planned_at": original_planned_at},
        )
        self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            db = get_db()
            next_event = db.execute(
                "SELECT status, planned_at, repair_record_id FROM scheduled_events WHERE id = ?",
                (event_id,),
            ).fetchone()
            completed_occurrences = db.execute(
                """
                SELECT se.status, se.planned_at, se.repair_record_id, r.description,
                       r.cause, r.repair_method, r.replaced_parts, r.notes
                FROM scheduled_events se
                JOIN repair_records r ON r.id = se.repair_record_id
                WHERE se.status = 'done' AND se.description = ?
                """,
                ("Dwutygodniowy przegląd",),
            ).fetchall()
            record_count = db.execute(
                "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                ("Dwutygodniowy przegląd",),
            ).fetchone()[0]

        self.assertEqual(next_event["status"], "planned")
        self.assertEqual(next_event["planned_at"], "2020-11-19 09:30")
        self.assertIsNone(next_event["repair_record_id"])
        self.assertEqual(len(completed_occurrences), 1)
        completed = completed_occurrences[0]
        self.assertEqual(completed["status"], "done")
        self.assertEqual(completed["planned_at"], original_planned_at)
        self.assertEqual(completed["description"], "Dwutygodniowy przegląd")
        self.assertEqual(completed["cause"], "Kontrola cykliczna")
        self.assertEqual(completed["repair_method"], "Sprawdzenie układu")
        self.assertEqual(completed["replaced_parts"], "Filtr")
        self.assertEqual(completed["notes"], "Uwagi cykliczne")
        self.assertEqual(record_count, 1)
        completed_page = self.client.get("/scheduled").get_data(as_text=True)
        self.assertIn("<td class=\"text-muted\">2 tygodnie</td>", completed_page)
        self.assertIn('data-repeat-display="2 tygodnie"', completed_page)

        stale_response = self.client.post(
            f"/scheduled/{event_id}/done",
            data={"expected_planned_at": original_planned_at},
        )
        self.assertEqual(stale_response.status_code, 302)
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM repair_records WHERE description = ?",
                    ("Dwutygodniowy przegląd",),
                ).fetchone()[0],
                1,
            )

    def test_admin_can_edit_machine_details_for_all_linked_records(self):
        for day in ("30", "31"):
            self.client.post(
                "/",
                data={
                    "occurrence_at": f"2026-08-{day}T08:00",
                    "machine_name": "Linia A",
                    "machine_type": "Pakowarka",
                    "serial_number": "P-100",
                    "event_type": "Awaria",
                    "description": f"Wpis {day}",
                },
            )
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-09-01T08:00",
                "machine_name": "Linia B",
                "machine_type": "Pakowarka",
                "serial_number": "P-200",
                "event_type": "Awaria",
                "description": "Druga maszyna",
            },
        )
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        with self.app.app_context():
            machine_id = get_db().execute(
                "SELECT id FROM machines WHERE serial_number = ?", ("P-100",)
            ).fetchone()["id"]

        response = self.client.post(
            f"/admin/machines/{machine_id}/edit",
            data={
                "name": "Linia A po zmianie",
                "machine_type": "Nowa pakowarka",
                "serial_number": "P-101",
            },
        )
        self.assertEqual(response.status_code, 302)
        page = self.client.get("/admin/machines").get_data(as_text=True)
        self.assertIn("Linia A po zmianie", page)
        self.assertIn("P-101", page)
        with self.app.app_context():
            names = get_db().execute(
                """
                SELECT DISTINCT m.name, m.machine_type, m.serial_number
                FROM repair_records r JOIN machines m ON m.id = r.machine_id
                WHERE r.description IN (?, ?)
                """,
                ("Wpis 30", "Wpis 31"),
            ).fetchall()
        self.assertEqual(len(names), 1)
        self.assertEqual(tuple(names[0]), ("Linia A po zmianie", "Nowa pakowarka", "P-101"))

        response = self.client.post(
            f"/admin/machines/{machine_id}/edit",
            data={
                "name": "Konflikt",
                "machine_type": "Pakowarka",
                "serial_number": "P-200",
            },
            follow_redirects=True,
        )
        self.assertIn("Maszyna o takich danych lub numerze seryjnym już istnieje.", response.get_data(as_text=True))
        with self.app.app_context():
            serial_number = get_db().execute(
                "SELECT serial_number FROM machines WHERE id = ?", (machine_id,)
            ).fetchone()["serial_number"]
        self.assertEqual(serial_number, "P-101")

    def test_admin_can_delete_machine_with_records_and_attachments(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Linia do usunięcia",
                "machine_type": "Pakowarka",
                "serial_number": "P-DEL",
                "event_type": "Awaria",
                "description": "Wpis do usunięcia",
            },
        )
        with self.app.app_context():
            record_id = int(
                get_db().execute(
                    "SELECT id FROM repair_records WHERE description = ?",
                    ("Wpis do usunięcia",),
                ).fetchone()["id"]
            )
        attachment_response = self.client.post(
            f"/api/record/{record_id}/attachments",
            data={"attachments": (io.BytesIO(b"test file"), "test.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(attachment_response.status_code, 200)
        attachment_dir = Path(self.app.config["UPLOAD_DIR"]) / str(record_id)
        self.assertTrue(attachment_dir.exists())

        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"
        with self.app.app_context():
            machine_id = get_db().execute(
                "SELECT id FROM machines WHERE serial_number = ?", ("P-DEL",)
            ).fetchone()["id"]

        response = self.client.post(f"/admin/machines/{machine_id}/delete")
        self.assertEqual(response.status_code, 302)
        self.assertFalse(attachment_dir.exists())
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM repair_records WHERE machine_id = ?", (machine_id,)
                ).fetchone()[0],
                0,
            )
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM attachments WHERE record_id = ?", (record_id,)
                ).fetchone()[0],
                0,
            )
            self.assertIsNone(
                get_db().execute("SELECT id FROM machines WHERE id = ?", (machine_id,)).fetchone()
            )

    def test_record_attachments_can_be_downloaded_as_one_zip(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Maszyna archiwum",
                "machine_type": "Testowa",
                "serial_number": "ZIP-001",
                "event_type": "Awaria",
                "description": "Rekord z wieloma załącznikami",
            },
        )
        with self.app.app_context():
            record_id = get_db().execute(
                "SELECT id FROM repair_records WHERE description = ?",
                ("Rekord z wieloma załącznikami",),
            ).fetchone()["id"]

        upload_response = self.client.post(
            f"/api/record/{record_id}/attachments",
            data={
                "attachments": [
                    (io.BytesIO(b"first image"), "zdjecie.jpg"),
                    (io.BytesIO(b"second image"), "zdjecie.jpg"),
                ]
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(upload_response.status_code, 200)

        response = self.client.get(f"/record/{record_id}/attachments.zip")
        self.assertEqual(response.status_code, 200)
        self.assertIn("zalaczniki_rekordu_", response.headers["Content-Disposition"])
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertEqual(archive.namelist(), ["zdjecie.jpg", "zdjecie_1.jpg"])
            self.assertEqual(archive.read("zdjecie.jpg"), b"first image")
            self.assertEqual(archive.read("zdjecie_1.jpg"), b"second image")
        response.close()

    def test_machine_management_requires_admin_role(self):
        response = self.client.get("/admin/machines")
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.location, "/")
        edit_response = self.client.post(
            "/admin/machines/1/edit",
            data={"name": "Nieuprawniona zmiana", "machine_type": "Typ", "serial_number": "123"},
        )
        delete_response = self.client.post("/admin/machines/1/delete")
        self.assertEqual(edit_response.status_code, 302)
        self.assertEqual(delete_response.status_code, 302)
        self.assertEqual(edit_response.location, "/")
        self.assertEqual(delete_response.location, "/")

    def test_machine_import_export_controls_are_on_data_page(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        data_page = self.client.get("/data").get_data(as_text=True)
        machines_page = self.client.get("/admin/machines").get_data(as_text=True)

        self.assertIn("Import / export bazy maszyn", data_page)
        self.assertIn('/admin/machines/export', data_page)
        self.assertIn('/admin/machines/import', data_page)
        self.assertNotIn("Eksportuj do Excel", machines_page)
        self.assertNotIn("Importuj z Excel", machines_page)
        self.assertIn("Dodaj maszynę", machines_page)

    def test_admin_machine_table_shows_related_record_count(self):
        for day in ("30", "31"):
            self.client.post(
                "/",
                data={
                    "occurrence_at": f"2026-08-{day}T08:00",
                    "machine_name": "Linia A",
                    "machine_type": "Pakowarka",
                    "serial_number": "P-100",
                    "event_type": "Awaria",
                    "description": f"Wpis {day}",
                },
            )
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        page = self.client.get("/admin/machines").get_data(as_text=True)

        self.assertIn(
            '<th aria-sort="none"><button type="button" class="machine-sort-button">POWIĄZANE WPISY',
            page,
        )
        self.assertIn(
            '<td>Linia A</td>\n                    <td>Pakowarka</td>\n'
            '                    <td>P-100</td>\n                    <td>—</td>\n'
            '                    <td>2</td>',
            page,
        )

    def test_serial_number_reuses_existing_machine_even_with_different_label(self):
        first = {
            "occurrence_at": "2026-08-30T08:00",
            "machine_name": "BOSCHERT",
            "machine_type": "Twin 1000",
            "serial_number": "282",
            "event_type": "Awaria",
            "description": "Pierwszy wpis",
        }
        second = {
            **first,
            "machine_name": "Inna nazwa",
            "machine_type": "Inny typ",
            "description": "Drugi wpis",
        }

        self.client.post("/", data=first)
        self.client.post("/", data=second)

        with self.app.app_context():
            machines = get_db().execute(
                "SELECT id, name, machine_type FROM machines WHERE serial_number = ?",
                ("282",),
            ).fetchall()
            records = get_db().execute(
                "SELECT DISTINCT machine_id FROM repair_records WHERE description IN (?, ?)",
                ("Pierwszy wpis", "Drugi wpis"),
            ).fetchall()

        self.assertEqual(len(machines), 1)
        self.assertEqual(machines[0]["name"], "BOSCHERT")
        self.assertEqual(machines[0]["machine_type"], "Twin 1000")
        self.assertEqual(len(records), 1)

    def test_recent_records_limit_all_displays_all_rows(self):
        valid_dates = [
            ("2026-08-30T08:00", "2026-08-30T09:00"),
            ("2026-08-31T08:00", "2026-08-31T09:00"),
            ("2026-09-01T08:00", "2026-09-01T09:00"),
        ]
        for idx, (occurrence_at, repaired_at) in enumerate(valid_dates):
            self.client.post(
                "/",
                data={
                    "occurrence_at": occurrence_at,
                    "repaired_at": repaired_at,
                    "machine_name": f"Maszyna {idx}",
                    "machine_type": "Typ",
                    "serial_number": f"SER-{idx}",
                    "event_type": "Awaria",
                    "description": f"Opis {idx}",
                },
            )

        response = self.client.get("/?limit=all")
        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn("Ostatnie 3 rekordów", page)
        self.assertIn("Opis 0", page)
        self.assertIn("Opis 1", page)
        self.assertIn("Opis 2", page)

    def test_dashboard_kpis_use_all_records_from_current_month_not_visible_limit(self):
        with self.client.session_transaction() as session:
            session["role"] = "admin"
        month_start = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        records = [
            (month_start, None, "Miesięczny rekord starszy"),
            (month_start.replace(hour=2), month_start.replace(hour=3), "Miesięczny rekord widoczny"),
            (month_start - timedelta(minutes=1), None, "Rekord z poprzedniego miesiąca"),
        ]
        for occurrence_at, repaired_at, description in records:
            form = {
                "occurrence_at": occurrence_at.strftime("%Y-%m-%dT%H:%M"),
                "machine_name": description,
                "machine_type": "Typ",
                "serial_number": description,
                "event_type": "Awaria",
                "description": description,
                "material_cost": "10",
            }
            if repaired_at:
                form["repaired_at"] = repaired_at.strftime("%Y-%m-%dT%H:%M")
            self.client.post("/", data=form)

        response = self.client.get("/?limit=1")
        page = response.get_data(as_text=True)
        match = re.search(r"const DASHBOARD_KPIS = (\{.*?\});", page)
        self.assertIsNotNone(match)
        kpis = json.loads(match.group(1))
        self.assertEqual(kpis["failures"], 2)
        self.assertEqual(kpis["downtime"], 1)
        self.assertEqual(kpis["cost"], 20)
        self.assertEqual(kpis["unrepaired"], 1)
        self.assertNotIn("Miesięczny rekord starszy", page)
        self.assertNotIn("Rekord z poprzedniego miesiąca", page)
        self.assertIn("Awarie w tym miesiącu", page)

    def test_dashboard_kpis_are_hidden_from_non_admin_users(self):
        response = self.client.get("/")
        page = response.get_data(as_text=True)

        self.assertNotIn("dashboard-kpis", page)
        self.assertNotIn("DASHBOARD_KPIS", page)
        self.assertNotIn("Awarie w tym miesiącu", page)

    def test_record_summaries_are_limited_in_recent_and_filtered_tables(self):
        description = "Opis z limitem\n" + ("szczegółowy tekst " * 60)
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "repaired_at": "2026-08-30T09:00",
                "machine_name": "Maszyna testowa",
                "machine_type": "Typ",
                "serial_number": "SER-400",
                "event_type": "Awaria",
                "description": description,
            },
        )
        expected_summary = description[:397] + "..."

        for path in ("/?limit=all", "/filter?description_terms=Opis"):
            with self.subTest(path=path):
                response = self.client.get(path)
                page = response.get_data(as_text=True)
                self.assertEqual(response.status_code, 200)
                self.assertIn(expected_summary, page)
                self.assertNotIn(description, page)
                self.assertIn("class=\"record-summary\"", page)

    def test_records_without_repair_date_are_highlighted_in_recent_and_filtered_tables(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Maszyna bez daty naprawienia",
                "machine_type": "Typ",
                "serial_number": "SER-401",
                "event_type": "Awaria",
                "description": "Oczekuje na naprawę",
            },
        )

        for path in ("/?limit=all", "/filter?description_terms=Oczekuje"):
            with self.subTest(path=path):
                response = self.client.get(path)
                page = response.get_data(as_text=True)
                self.assertEqual(response.status_code, 200)
                self.assertIn("date-no-repair", page)
                self.assertIn('title="Brak daty naprawienia"', page)
        css = get_static_css(self.client)
        self.assertIn(".date-no-repair { color: var(--orange); }", css)

    def test_import_excel_file(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "Maszyna",
                "Typ",
                "Numer",
                "Rodzaj zdarzenia",
                "Opis",
                "Osoby biorące udział w interwencji",
                "Data wystąpienia",
                "Data naprawienia/ interwencji",
            ]
        )
        sheet.append(
            [
                "Gilotyna",
                "Tnąca",
                "G-12",
                "Awaria",
                "Wymiana noża",
                "Marek Nowak, Piotr Kowalski",
                "2026-08-20 07:00",
                "2026-08-20 08:30",
            ]
        )
        sheet.append(
            [
                "",
                "",
                "",
                "Awaria",
                "Niekompletny wiersz",
                "",
                "2026-08-21 07:00",
                "2026-08-21 08:00",
            ]
        )
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)

        response = self.client.post(
            "/import",
            data={"file": (buffer, "raport.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn("Zaimportowano 1 rekordów, pominięto 1.", page)
        self.assertIn("Gilotyna", page)
        with self.app.app_context():
            participants = get_db().execute(
                "SELECT participants FROM repair_records WHERE description = ?",
                ("Wymiana noża",),
            ).fetchone()["participants"]
        self.assertEqual(participants, "Marek Nowak, Piotr Kowalski")

    def test_importing_exported_records_again_skips_duplicates(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "Maszyna", "Typ", "Numer", "Rodzaj zdarzenia", "Opis", "Przyczyna",
                "Sposób naprawy", "Data wystąpienia", "Data naprawienia/ interwencji",
                "MTG maszyny", "Praco chłonność [80PLN/h]", "Materiał własny PLN",
                "Części zamienne PLN", "Serwis PLN", "Uwagi",
            ]
        )
        sheet.append(
            [
                "Tokarka", "CNC", "T-7", "Awaria", "Wymiana łożyska", "Zużycie",
                "Wymiana", "2026-08-20 07:00", "2026-08-20 08:30", "1250", 1.5,
                10, 20, 5, "Test importu",
            ]
        )
        source = io.BytesIO()
        workbook.save(source)

        first_import = self.client.post(
            "/import",
            data={"file": (io.BytesIO(source.getvalue()), "zdarzenia.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertIn("Zaimportowano 1 rekordów, pominięto 0.", first_import.get_data(as_text=True))

        exported = self.client.get("/export")
        second_import = self.client.post(
            "/import",
            data={"file": (io.BytesIO(exported.data), "eksport.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertIn("Zaimportowano 0 rekordów, pominięto 1.", second_import.get_data(as_text=True))
        with self.app.app_context():
            count = get_db().execute("SELECT COUNT(*) FROM repair_records").fetchone()[0]
        self.assertEqual(count, 1)

    def test_records_bundle_restores_attachments_for_imported_and_duplicate_records(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Maszyna paczki",
                "machine_type": "Testowa",
                "serial_number": "BUNDLE-1",
                "event_type": "Awaria",
                "description": "Zdarzenie z załącznikiem",
            },
        )
        with self.app.app_context():
            source_record_id = get_db().execute(
                "SELECT id FROM repair_records WHERE description = ?", ("Zdarzenie z załącznikiem",)
            ).fetchone()["id"]
        self.client.post(
            f"/api/record/{source_record_id}/attachments",
            data={"attachments": (io.BytesIO(b"plik do odzyskania"), "raport.txt")},
            content_type="multipart/form-data",
        )
        bundle_response = self.client.get("/export/bundle")
        self.assertEqual(bundle_response.status_code, 200)
        bundle_bytes = bundle_response.data

        target_app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(Path(self.temp_dir.name) / "target.sqlite3"),
                "UPLOAD_DIR": str(Path(self.temp_dir.name) / "target_uploads"),
                "SECRET_KEY": "test",
            }
        )
        target_client = target_app.test_client()
        target_client.post(
            "/",
            data={
                "occurrence_at": "2026-08-29T08:00",
                "machine_name": "Inna maszyna",
                "machine_type": "Testowa",
                "serial_number": "OTHER-1",
                "event_type": "Awaria",
                "description": "Rekord przesuwający numer ID",
            },
        )

        first_import = target_client.post(
            "/data",
            data={"file": (io.BytesIO(bundle_bytes), "zdarzenia.zip")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertIn("Załączniki: dodano 1, pominięto 0.", first_import.get_data(as_text=True))
        with target_app.app_context():
            target_record_id = get_db().execute(
                "SELECT id FROM repair_records WHERE description = ?", ("Zdarzenie z załącznikiem",)
            ).fetchone()["id"]
            self.assertNotEqual(target_record_id, source_record_id)
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM attachments WHERE record_id = ?", (target_record_id,)
                ).fetchone()[0],
                1,
            )
        second_import = target_client.post(
            "/data",
            data={"file": (io.BytesIO(bundle_bytes), "zdarzenia.zip")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertIn("Załączniki: dodano 0, pominięto 1.", second_import.get_data(as_text=True))
        with target_app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM attachments WHERE record_id = ?", (target_record_id,)
                ).fetchone()[0],
                1,
            )

    def test_admin_can_download_full_instance_backup_with_database_and_uploads(self):
        self.client.post(
            "/",
            data={
                "occurrence_at": "2026-08-30T08:00",
                "machine_name": "Maszyna backup",
                "machine_type": "Testowa",
                "serial_number": "BACKUP-1",
                "event_type": "Awaria",
                "description": "Rekord kopii pełnej",
            },
        )
        with self.app.app_context():
            record_id = get_db().execute(
                "SELECT id FROM repair_records WHERE description = ?", ("Rekord kopii pełnej",)
            ).fetchone()["id"]
        self.client.post(
            f"/api/record/{record_id}/attachments",
            data={"attachments": (io.BytesIO(b"zawartosc zalacznika"), "plik.txt")},
            content_type="multipart/form-data",
        )
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        response = self.client.get("/admin/export/full-instance-backup")
        self.assertEqual(response.status_code, 200)
        self.assertIn("instance_backup_", response.headers["Content-Disposition"])
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertIn("instance/repair_reports.sqlite3", archive.namelist())
            self.assertIn(f"instance/uploads/{record_id}/plik.txt", archive.namelist())
            self.assertEqual(archive.read(f"instance/uploads/{record_id}/plik.txt"), b"zawartosc zalacznika")
            snapshot = sqlite3.connect(":memory:")
            snapshot.deserialize(archive.read("instance/repair_reports.sqlite3"))
            self.assertEqual(
                snapshot.execute("SELECT COUNT(*) FROM repair_records WHERE description = ?", ("Rekord kopii pełnej",)).fetchone()[0],
                1,
            )
            snapshot.close()

    def test_user_login_and_create_permission(self):
        app = create_app(
            {
                "TESTING": False,
                "REQUIRE_LOGIN": True,
                "DATABASE": str(self.database_path),
                "SECRET_KEY": "test",
            }
        )
        client = app.test_client()

        with app.app_context():
            get_db().execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    is_admin INTEGER NOT NULL DEFAULT 0,
                    can_view_records INTEGER NOT NULL DEFAULT 1,
                    can_create_records INTEGER NOT NULL DEFAULT 0,
                    can_edit_records INTEGER NOT NULL DEFAULT 0,
                    can_delete_records INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            get_db().execute(
                "INSERT INTO users(username, password_hash, is_admin, can_view_records, can_create_records, can_edit_records, can_delete_records, created_at, updated_at) VALUES (?, ?, 0, 1, 1, 1, 0, ?, ?)",
                ("operator", generate_password_hash("test"), "2026-01-01T00:00:00", "2026-01-01T00:00:00"),
            )
            get_db().commit()

        response = client.post(
            "/login",
            data={"login": "operator", "password": "test"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("Strona główna", response.get_data(as_text=True))
        self.assertIn(">Rejestr zdarzeń</a>", response.get_data(as_text=True))
        page = response.get_data(as_text=True)
        self.assertIn('href="/change-password"', page)
        self.assertIn(">operator</a>", page)
        self.assertNotIn(">Zmień hasło</a>", page)

        account_page = client.get("/change-password").get_data(as_text=True)
        self.assertIn("Zmiana hasła", account_page)
        self.assertIn('href="/logout"', account_page)

        response = client.get("/new")
        self.assertEqual(response.status_code, 200)

        response = client.post(
            "/new",
            data={
                "occurrence_at": "2026-09-21T08:00",
                "machine_name": "Maszyna operatora",
                "machine_type": "Testowa",
                "serial_number": "OP-001",
                "event_type": "Awaria",
                "description": "Rekord utworzony przez operatora",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)

        with app.app_context():
            activity = get_db().execute(
                """
                SELECT action, username, target_type
                FROM user_activity_log
                WHERE action = 'create_record'
                ORDER BY id DESC
                LIMIT 1
                """
            ).fetchone()
        self.assertIsNotNone(activity)
        self.assertEqual(activity["action"], "create_record")
        self.assertEqual(activity["username"], "operator")
        self.assertEqual(activity["target_type"], "record")

        with app.app_context():
            get_db().execute("DELETE FROM users WHERE username = ?", ("operator",))
            get_db().commit()

    def test_admin_can_edit_existing_user_permissions(self):
        with self.app.app_context():
            cursor = get_db().execute(
                """
                INSERT INTO users(
                    username, password_hash, is_admin, can_view_records,
                    can_create_records, can_edit_records, can_delete_records,
                    created_at, updated_at
                ) VALUES (?, ?, 0, 1, 1, 0, 0, ?, ?)
                """,
                (
                    "permissions_user",
                    generate_password_hash("unchanged"),
                    "2026-01-01T00:00:00",
                    "2026-01-01T00:00:00",
                ),
            )
            user_id = int(cursor.lastrowid)
            get_db().commit()

        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"
            session["user_id"] = 1
            session["username"] = "admin"

        response = self.client.get(f"/admin/users?edit={user_id}")
        self.assertEqual(response.status_code, 200)
        self.assertIn('name="user_id"', response.get_data(as_text=True))
        self.assertIn("permissions_user", response.get_data(as_text=True))

        response = self.client.post(
            "/admin/users",
            data={
                "user_id": str(user_id),
                "username": "permissions_user",
                "is_admin": "0",
                "can_view_records": "on",
                "can_create_records": "on",
                "can_edit_records": "on",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)

        with self.app.app_context():
            user = get_db().execute(
                "SELECT can_view_records, can_create_records, can_edit_records, can_delete_records FROM users WHERE id = ?",
                (user_id,),
            ).fetchone()
        self.assertEqual(tuple(user), (1, 1, 1, 0))

    def test_admin_can_export_full_activity_log(self):
        with self.app.app_context():
            get_db().execute(
                """
                INSERT INTO user_activity_log(
                    user_id, username, action, target_type, target_id, details, ip_address, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    1,
                    "admin",
                    "create_record",
                    "record",
                    123,
                    "Testowe zdarzenie",
                    "203.0.113.10",
                    "2026-09-21T11:50:00+00:00",
                ),
            )
            get_db().commit()

        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"
            session["user_id"] = 1
            session["username"] = "admin"

        page_response = self.client.get("/admin/users")
        self.assertEqual(page_response.status_code, 200)
        self.assertIn("203.0.113.10", page_response.get_data(as_text=True))

        response = self.client.get("/admin/users/activity/export")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        workbook = load_workbook(io.BytesIO(response.data))
        rows = list(workbook.active.iter_rows(values_only=True))
        self.assertEqual(
            rows[0],
            ("Data", "Użytkownik", "Akcja", "Typ obiektu", "ID obiektu", "Szczegóły", "Adres IP"),
        )
        self.assertIn(
            ("21-09-2026 13:50", "admin", "create_record", "record", 123, "Testowe zdarzenie", "203.0.113.10"),
            rows,
        )

    def test_admin_user_activity_sort_is_preserved_across_pages(self):
        with self.app.app_context():
            for index in range(21):
                get_db().execute(
                    """
                    INSERT INTO user_activity_log(
                        user_id, username, action, target_type, target_id, details, ip_address, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        1,
                        f"user-{index:02d}",
                        "login",
                        None,
                        None,
                        f"Zdarzenie {index:02d}",
                        "203.0.113.10",
                        "2026-09-21T11:50:00+00:00",
                    ),
                )
            get_db().commit()

        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"
            session["user_id"] = 1
            session["username"] = "admin"

        response = self.client.get("/admin/users?sort=username&direction=asc")
        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertEqual(page.count('class="activity-sort-button"'), 5)
        self.assertIn('aria-sort="ascending"', page)
        self.assertIn("page=2&amp;sort=username&amp;direction=asc", page)
        self.assertLess(page.index("user-00"), page.index("user-19"))
        self.assertNotIn("user-20", page)

        second_page_response = self.client.get("/admin/users?page=2&sort=username&direction=asc")
        self.assertEqual(second_page_response.status_code, 200)
        second_page = second_page_response.get_data(as_text=True)
        self.assertIn("user-20", second_page)
        self.assertIn("page=1&amp;sort=username&amp;direction=asc", second_page)

    def test_activity_log_records_request_ip(self):
        with self.app.test_request_context(
            "/admin/users",
            environ_overrides={"REMOTE_ADDR": "203.0.113.25"},
        ):
            log_user_activity("test", "Testowe zdarzenie")

        with self.app.app_context():
            activity = get_db().execute(
                "SELECT ip_address FROM user_activity_log WHERE action = 'test'"
            ).fetchone()
        self.assertEqual(activity["ip_address"], "203.0.113.25")

    def test_existing_activity_log_database_is_migrated_for_ip_address(self):
        legacy_database = Path(self.temp_dir.name) / "legacy.sqlite3"
        connection = sqlite3.connect(legacy_database)
        try:
            connection.execute(
                """
                CREATE TABLE user_activity_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    username TEXT,
                    action TEXT NOT NULL,
                    target_type TEXT,
                    target_id INTEGER,
                    details TEXT,
                    created_at TEXT NOT NULL
                )
                """
            )
            connection.commit()
        finally:
            connection.close()

        app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(legacy_database),
                "UPLOAD_DIR": str(Path(self.temp_dir.name) / "legacy_uploads"),
                "SECRET_KEY": "test",
            }
        )
        with app.app_context():
            columns = {
                row["name"]
                for row in get_db().execute("PRAGMA table_info(user_activity_log)").fetchall()
            }
        self.assertIn("ip_address", columns)

    def test_editor_can_open_edit_controls_in_filter_and_machine_history(self):
        app = create_app(
            {
                "TESTING": False,
                "REQUIRE_LOGIN": True,
                "DATABASE": str(self.database_path),
                "SECRET_KEY": "test",
            }
        )
        client = app.test_client()
        with app.app_context():
            get_db().execute(
                """
                INSERT INTO users(
                    username, password_hash, is_admin, can_view_records,
                    can_create_records, can_edit_records, can_delete_records,
                    created_at, updated_at
                ) VALUES (?, ?, 0, 1, 1, 1, 1, ?, ?)
                """,
                (
                    "editor",
                    generate_password_hash("editor-pass"),
                    "2026-01-01T00:00:00",
                    "2026-01-01T00:00:00",
                ),
            )
            get_db().commit()

        response = client.post(
            "/login",
            data={"login": "editor", "password": "editor-pass"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)

        filter_page = client.get("/filter")
        new_record_page = client.get("/new")
        for page in (filter_page, new_record_page):
            html = page.get_data(as_text=True)
            self.assertEqual(page.status_code, 200)
            self.assertIn("const CAN_EDIT = 1;", html)
            self.assertIn("(USER_ROLE === 'admin' || CAN_EDIT)", html)

    def test_admin_can_create_and_edit_machines_with_warranty_calculation(self):
        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        page = self.client.get("/admin/machines?new=1")
        self.assertEqual(page.status_code, 200)
        page_html = page.get_data(as_text=True)
        self.assertIn("Dodaj maszynę", page_html)
        self.assertIn('name="machine_category"', page_html)
        self.assertEqual(page_html.count('class="machine-sort-button"'), 5)
        for heading in (
            "NAZWA / FIRMA",
            "TYP / MODEL",
            "NUMER SERYJNY",
            "RODZAJ",
            "POWIĄZANE WPISY",
        ):
            self.assertIn(heading, page_html)
        self.assertIn('id="machine-modal-view"', page_html)
        self.assertIn('id="machine-modal-edit"', page_html)
        self.assertIn('id="machine-modal-save"', page_html)
        self.assertIn("machineDeleteForm.addEventListener('submit'", page_html)
        self.assertNotIn('id="machine-modal-pdf"', page_html)

        response = self.client.post(
            "/admin/machines/new",
            data={
                "name": "Producent",
                "machine_type": "Model 1",
                "serial_number": "M-001",
                "machine_category": "Prasa",
                "commissioning_date": "2024-02-29",
                "warranty_period_months": "12",
                "power": "15.5",
                "power_unit": "kW",
                "protection_a": "32",
                "seller_data": "Serwis Example",
                "notes": "Nowa maszyna",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("Maszyna została dodana.", response.get_data(as_text=True))

        with self.app.app_context():
            machine = get_db().execute(
                "SELECT * FROM machines WHERE serial_number = ?", ("M-001",)
            ).fetchone()
            machine_id = machine["id"]
            self.assertEqual(machine["warranty_end_date"], "2025-02-28")
            self.assertEqual(machine["machine_category"], "Prasa")
            self.assertEqual(machine["power"], 15.5)

        response = self.client.post(
            f"/admin/machines/{machine_id}/edit",
            data={
                "name": "Producent",
                "machine_type": "Model 1",
                "serial_number": "M-001",
                "machine_category": "Prasa hydrauliczna",
                "commissioning_date": "2024-01-31",
                "warranty_period_months": "1",
                "power": "15.5",
                "power_unit": "kW",
                "protection_a": "32",
                "seller_data": "Serwis Example",
                "notes": "Zmieniono",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with self.app.app_context():
            machine = get_db().execute(
                "SELECT machine_category, warranty_end_date, notes FROM machines WHERE id = ?",
                (machine_id,),
            ).fetchone()
        self.assertEqual(tuple(machine), ("Prasa hydrauliczna", "2024-02-29", "Zmieniono"))

    def test_admin_can_import_and_export_machine_workbook(self):
        with self.app.app_context():
            get_db().execute(
                """
                INSERT INTO machines(name, machine_type, serial_number, notes)
                VALUES (?, ?, ?, ?)
                """,
                ("Producent", "Model istniejący", "SER-001", "Zachowaj tę uwagę"),
            )
            get_db().commit()

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "MASZYNY"
        sheet.append(
            [
                "Nazwa / firma",
                "Typ / model",
                "Numer seryjny",
                "Rodzaj maszyny",
                "Data uruchomienia",
                "Okres gwarancji (miesiące)",
                "Data końca gwarancji",
                "Moc",
                "Jednostka mocy",
                "Zabezpieczenie [A]",
                "Dane sprzedawcy",
                "Uwagi",
            ]
        )
        sheet.append(
            [
                "Inna nazwa z pliku",
                "Inny model z pliku",
                "SER001",
                "Tokarka",
                datetime(2023, 1, 31),
                12,
                datetime(2024, 1, 31),
                7.5,
                "kW",
                16,
                "Dostawca A",
                None,
            ]
        )
        sheet.append(
            [
                "Nowy producent",
                "Nowy model",
                "SER-002",
                "Frezarka",
                datetime(2024, 3, 1),
                24,
                datetime(2026, 3, 1),
                11,
                "kW",
                20,
                "Dostawca B",
                "Maszyna z importu",
            ]
        )
        workbook_bytes = io.BytesIO()
        workbook.save(workbook_bytes)
        workbook_bytes.seek(0)

        with self.client.session_transaction() as session:
            session["logged_in"] = True
            session["role"] = "admin"

        response = self.client.post(
            "/admin/machines/import",
            data={"file": (workbook_bytes, "maszyny.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.request.path, "/data")
        self.assertIn("Dodano: 1, zaktualizowano: 1, pominięto: 0", response.get_data(as_text=True))

        with self.app.app_context():
            machines = get_db().execute(
                "SELECT * FROM machines WHERE serial_number IN (?, ?) ORDER BY serial_number",
                ("SER-001", "SER-002"),
            ).fetchall()
        self.assertEqual(len(machines), 2)
        existing, new_machine = machines
        self.assertEqual(existing["name"], "Producent")
        self.assertEqual(existing["machine_type"], "Model istniejący")
        self.assertEqual(existing["machine_category"], "Tokarka")
        self.assertEqual(existing["commissioning_date"], "2023-01-31")
        self.assertEqual(existing["warranty_end_date"], "2024-01-31")
        self.assertEqual(existing["notes"], "Zachowaj tę uwagę")
        self.assertEqual(new_machine["warranty_end_date"], "2026-03-01")
        self.assertEqual(new_machine["notes"], "Maszyna z importu")

        exported = self.client.get("/admin/machines/export")
        self.assertEqual(exported.status_code, 200)
        exported_workbook = load_workbook(io.BytesIO(exported.data), data_only=True)
        exported_rows = list(exported_workbook.active.iter_rows(values_only=True))
        self.assertEqual(
            exported_rows[0],
            (
                "Nazwa / firma",
                "Typ / model",
                "Numer seryjny",
                "Rodzaj maszyny",
                "Data uruchomienia",
                "Okres gwarancji (miesiące)",
                "Data końca gwarancji",
                "Moc",
                "Jednostka mocy",
                "Zabezpieczenie [A]",
                "Dane sprzedawcy",
                "Uwagi",
            ),
        )
        self.assertEqual(exported_workbook.active.title, "MASZYNY")

    def test_existing_machine_database_is_migrated_without_losing_rows(self):
        legacy_database = Path(self.temp_dir.name) / "legacy_machines.sqlite3"
        connection = sqlite3.connect(legacy_database)
        try:
            connection.execute(
                """
                CREATE TABLE machines (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    machine_type TEXT NOT NULL,
                    serial_number TEXT NOT NULL,
                    UNIQUE(name, machine_type, serial_number)
                )
                """
            )
            connection.execute(
                "INSERT INTO machines(name, machine_type, serial_number) VALUES (?, ?, ?)",
                ("Producent", "Model", "LEGACY-001"),
            )
            connection.commit()
        finally:
            connection.close()

        app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(legacy_database),
                "UPLOAD_DIR": str(Path(self.temp_dir.name) / "legacy_uploads"),
                "SECRET_KEY": "test",
            }
        )
        with app.app_context():
            machine = get_db().execute(
                "SELECT * FROM machines WHERE serial_number = ?", ("LEGACY-001",)
            ).fetchone()
            columns = {
                row["name"] for row in get_db().execute("PRAGMA table_info(machines)")
            }
        self.assertEqual(machine["name"], "Producent")
        self.assertTrue(
            {
                "machine_category",
                "commissioning_date",
                "warranty_period_months",
                "warranty_end_date",
                "power",
                "power_unit",
                "protection_a",
                "seller_data",
                "notes",
            }.issubset(columns)
        )

    def test_machine_import_can_skip_creation_of_unmatched_machines(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Firma", "Nazwa", "Nr fab.", "Rodzaj"])
        sheet.append(["Producent", "Model", "NEW-001", "Prasa"])
        content = io.BytesIO()
        workbook.save(content)

        with self.app.app_context():
            result = import_machine_workbook(content.getvalue(), create_missing=False)
            count = get_db().execute("SELECT COUNT(*) FROM machines").fetchone()[0]
        self.assertEqual(result[:3], (0, 0, 1))
        self.assertEqual(count, 0)
        self.assertIn("nie znaleziono maszyny w bazie", result[3][0])


if __name__ == "__main__":
    unittest.main()
