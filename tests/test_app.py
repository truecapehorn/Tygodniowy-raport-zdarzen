import io
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from app import create_app


class RepairReportAppTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "test.sqlite3"
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(self.database_path),
                "SECRET_KEY": "test",
            }
        )
        self.client = self.app.test_client()

    def tearDown(self):
        self.temp_dir.cleanup()

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


if __name__ == "__main__":
    unittest.main()
