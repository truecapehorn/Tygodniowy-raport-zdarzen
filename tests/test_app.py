import io
import sqlite3
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate
from werkzeug.security import generate_password_hash

from app import PageNumberCanvas, create_app, get_db, init_db, log_user_activity


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
        self.assertNotIn("Eksportuj całą bazę", data_page.get_data(as_text=True))

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
        self.assertIn("grid-template-columns: repeat(4, minmax(0, 1fr))", html)
        self.assertIn('class="field scheduled-third-row-start"', html)
        self.assertIn('class="field scheduled-notes-field"', html)
        self.assertIn('id="repeat-interval-field" hidden', html)
        self.assertIn(".scheduled-form-grid > .field[hidden] { display: none; }", html)
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
        self.assertEqual(year_page.status_code, 200)
        year_html = year_page.get_data(as_text=True)
        self.assertIn('data-calendar-view="year"', year_html)
        self.assertEqual(year_html.count('class="calendar-month-heading"'), 12)
        self.assertIn('class="calendar-week-label">Tydz.</div>', year_html)
        self.assertRegex(year_html, r'class="calendar-week-number">\s*1\s*</div>')
        self.assertIn('class="calendar-dot badge-przeglad"', year_html)
        self.assertIn('class="calendar-dot badge-interwencja"', year_html)
        self.assertIn('class="calendar-dot badge-modernizacja"', year_html)
        self.assertIn(".calendar-year-view .calendar-dot {", year_html)
        self.assertIn("width: .85rem;", year_html)
        self.assertIn(".calendar-year-view .calendar-dot.badge-interwencja { background: var(--yellow); }", year_html)
        self.assertIn(".calendar-year-view .calendar-dot.badge-modernizacja { background: var(--purple); }", year_html)

        month_page = self.client.get("/scheduled?calendar_view=month&year=2026&month=5")
        self.assertEqual(month_page.status_code, 200)
        month_html = month_page.get_data(as_text=True)
        self.assertIn('data-calendar-view="month"', month_html)
        self.assertRegex(month_html, r'class="calendar-week-number">\s*18\s*</div>')
        self.assertIn("title=\"14-05-2026 10:30 — Prasa Yellow — Interwencja — Yellow calendar event\"", month_html)
        self.assertNotIn("title=\"05-01-2026 07:00 — Prasa Blue", month_html)
        self.assertNotIn("title=\"01-12-2026 12:15 — Prasa Purple", month_html)

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
        self.assertIn(target_week.strftime("%d.%m"), selected_week_html)
        self.assertIn("10:30 · Interwencja · Prasa Yellow", selected_week_html)
        self.assertIn('id="calendar-iso-week" type="week" name="iso_week" value="2026-W20"', selected_week_html)
        self.assertIn('.calendar-controls input[type="week"] { width: 9.5rem; }', selected_week_html)
        self.assertRegex(
            selected_week_html,
            r'href="/(?:scheduled/new|planowanie/nowe)\?planned_at=2026-05-14T10:00"',
        )
        self.assertRegex(
            selected_week_html,
            r'href="/(?:scheduled|planowanie)\?calendar_view=week&amp;iso_week=2026-W19"',
        )

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
        self.assertIn('href="/records-calendar"', home_page.get_data(as_text=True))
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

        week_html = self.client.get(
            "/records-calendar?calendar_view=week&iso_week=2026-W10"
        ).get_data(as_text=True)
        self.assertIn("09:30 · Awaria · Prasa calendar", week_html)
        self.assertNotIn("/new?occurrence_at=", week_html)

        day_html = self.client.get(
            "/records-calendar?calendar_view=day&calendar_date=2026-03-02"
        ).get_data(as_text=True)
        self.assertIn('data-calendar-view="day"', day_html)
        self.assertIn('data-create-url="/new?occurrence_at=2026-03-02T09:30"', day_html)
        self.assertIn("09:30 · Awaria · Prasa calendar", day_html)
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
        self.assertIn("Zaplanuj nowe zdarzenie", html)
        self.assertIn("15-12-2020 13:45", html)
        form_page = self.client.get("/scheduled/new").get_data(as_text=True)
        self.assertIn('id="scheduled-machine-name"', form_page)
        self.assertIn('placeholder="dd-mm-rrrr gg:mm"', form_page)
        self.assertNotIn('<option value="Awaria">', form_page)
        self.assertIn('<option value="Przegląd">', form_page)
        home_page = self.client.get("/").get_data(as_text=True)
        self.assertIn("na 15-12-2020 13:45.", home_page)
        self.assertIn("Zaplanowane zdarzenia</div>\n        <strong style=\"display:block;font-size:2rem;color:var(--accent);\">1", html)
        self.assertIn("Zrealizowane zdarzenia</div>\n        <strong style=\"display:block;font-size:2rem;color:var(--green);\">0", html)
        self.assertIn('class="record-row scheduled-event-row" tabindex="0"', html)
        self.assertRegex(html, r'action="/(?:scheduled|planowanie)/\d+/delete"')
        self.assertIn(">Usuń</button>", html)
        self.assertNotIn(">Szczegóły</button>", html)
        self.assertIn('id="scheduled-modal-delete-form"', html)
        self.assertEqual(
            html.count("onsubmit=\"return confirm('Czy na pewno usunąć to zaplanowane zdarzenie?"),
            2,
        )
        self.assertIn('scheduledModalDeleteForm.action = `/scheduled/${id}/delete`', html)
        self.assertIn("row.addEventListener('click', (event) =>", html)
        self.assertIn("event.target.closest('button, form, a')", html)
        self.assertIn('id="scheduled-modal-edit-button" type="button"', html)
        self.assertIn('id="scheduled-modal-edit-form" method="post" style="display:none;"', html)
        self.assertIn('id="edit-repeat-interval-field" hidden', html)
        self.assertIn(".scheduled-edit-grid > .field[hidden] { display: none; }", html)
        self.assertIn("editRepeatUnitInput.addEventListener('change', updateEditRepeatIntervalVisibility)", html)
        self.assertIn('name="cause"', html)
        self.assertIn("scheduled-edit-description,\n    .scheduled-edit-notes { grid-column: 1 / -1; }", html)
        self.assertIn(".scheduled-edit-third { grid-column: span 4; }", html)
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
        self.assertIn("Zaplanowane zdarzenia</div>\n        <strong style=\"display:block;font-size:2rem;color:var(--accent);\">0", completed_page)
        self.assertIn("Zrealizowane zdarzenia</div>\n        <strong style=\"display:block;font-size:2rem;color:var(--green);\">1", completed_page)
        self.assertIn('class="record-row scheduled-event-row" tabindex="0"', completed_page)
        self.assertRegex(completed_page, r'action="/(?:scheduled|planowanie)/\d+/delete"')
        self.assertNotIn(">Szczegóły</button>", completed_page)
        self.assertEqual(
            completed_page.count("onsubmit=\"return confirm('Czy na pewno usunąć to zaplanowane zdarzenie?"),
            2,
        )

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
                self.assertIn('class="date-no-repair"', page)
                self.assertIn('title="Brak daty naprawienia"', page)

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
        self.assertIn("Zaimportowano 1 rekordów.", page)
        self.assertIn("Gilotyna", page)
        with self.app.app_context():
            participants = get_db().execute(
                "SELECT participants FROM repair_records WHERE description = ?",
                ("Wymiana noża",),
            ).fetchone()["participants"]
        self.assertEqual(participants, "Marek Nowak, Piotr Kowalski")

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
        page = response.get_data(as_text=True)
        self.assertIn('class="logout-link"', page)
        self.assertIn("<strong>operator</strong>", page)

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


if __name__ == "__main__":
    unittest.main()
