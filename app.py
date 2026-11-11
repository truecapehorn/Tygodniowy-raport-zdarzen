

from __future__ import annotations

import calendar
import io
import json
import math
import queue
import re
import secrets
import sqlite3
import tempfile
import threading
import unicodedata
import zipfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from flask import (
    Flask,
    Response,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image as PdfImage
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from werkzeug.security import check_password_hash, generate_password_hash

UTC = timezone.utc


class PageNumberCanvas(Canvas):
    def __init__(self, *args: Any, **kwargs: Any):
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict[str, Any]] = []

    def showPage(self) -> None:
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        page_count = len(self._saved_page_states)
        for page_state in self._saved_page_states:
            self.__dict__.update(page_state)
            self.setFont("Helvetica", 8)
            self.drawCentredString(A4[0] / 2, 8 * mm, f"{self._pageNumber}/{page_count}")
            super().showPage()
        super().save()


EVENT_TYPES = [
    "Awaria",
    "Przegląd",
    "Modernizacja",
    "Interwencja",
    "Inne",
]
SCHEDULED_EVENT_TYPES = [*EVENT_TYPES, "Własne"]

LABOR_RATE_PLN_PER_HOUR = 80


def get_labor_rate() -> float:
    """Get labor rate from app config or fallback to default."""
    try:
        from flask import current_app
        return float(current_app.config.get("LABOR_RATE_PLN_PER_HOUR", LABOR_RATE_PLN_PER_HOUR))
    except RuntimeError:
        # No app context
        return float(LABOR_RATE_PLN_PER_HOUR)

EXPORT_COLUMNS = [
    "Nr",
    "ROK",
    "Tydzień roku",
    "Lp. w tygodniu",
    "Maszyna",
    "Typ",
    "Numer",
    "Rodzaj zdarzenia",
    "Opis",
    "Przyczyna",
    "Sposób naprawy",
    "Wymienione części",
    "Osoby biorące udział w interwencji",
    "Data wystąpienia",
    "Data naprawienia/ interwencji",
    "MTG maszyny",
    "Wykluczenie z produkcji [rob.godzina]",
    "Praco chłonność [80PLN/h]",
    "Materiał własny PLN",
    "Części zamienne PLN",
    "Serwis PLN",
    "Koszt brutto naprawy PLN",
    "Uwagi",
]

HEADER_ALIASES = {
    "nr": "Nr",
    "numer porzadkowy": "Nr",
    "rok": "ROK",
    "tydzien roku": "Tydzień roku",
    "lp w tygodniu": "Lp. w tygodniu",
    "maszyna": "Maszyna",
    "typ": "Typ",
    "numer": "Numer",
    "nr fabryczny": "Numer",
    "rodzaj zdarzenia": "Rodzaj zdarzenia",
    "opis": "Opis",
    "przyczyna": "Przyczyna",
    "sposob naprawy": "Sposób naprawy",
    "wymienione czesci": "Wymienione części",
    "osoby biorace udzial w interwencji": "Osoby biorące udział w interwencji",
    "uczestnicy interwencji": "Osoby biorące udział w interwencji",
    "data wystapienia": "Data wystąpienia",
    "data naprawienia interwencji": "Data naprawienia/ interwencji",
    "data naprawienia": "Data naprawienia/ interwencji",
    "mtg maszyny": "MTG maszyny",
    "wykluczenie z produkcji robgodzina": "Wykluczenie z produkcji [rob.godzina]",
    "pracochlonnosc 80plnh": "Praco chłonność [80PLN/h]",
    "pracochlonnosc": "Praco chłonność [80PLN/h]",
    "praco chlonnosc 80plnh": "Praco chłonność [80PLN/h]",
    "praco chlonnosc": "Praco chłonność [80PLN/h]",
    "material wlasny pln": "Materiał własny PLN",
    "czesci zamienne pln": "Części zamienne PLN",
    "serwis pln": "Serwis PLN",
    "koszt brutto naprawy pln": "Koszt brutto naprawy PLN",
    "uwagi": "Uwagi",
}

MACHINE_EXPORT_COLUMNS = [
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

MACHINE_IMPORT_HEADER_ALIASES = {
    "nazwa firma": "name",
    "firma": "name",
    "maszyna": "name",
    "typ model": "machine_type",
    "nazwa": "machine_type",
    "model": "machine_type",
    "typ": "machine_type",
    "nr fab": "serial_number",
    "numer seryjny": "serial_number",
    "numer": "serial_number",
    "rodzaj": "machine_category",
    "rodzaj maszyny": "machine_category",
    "data uruchomienia": "commissioning_date",
    "okres gwarancji mies": "warranty_period_months",
    "okres gwarancji miesiace": "warranty_period_months",
    "okres gwarancji": "warranty_period_months",
    "data konca gwarancji": "warranty_end_date",
    "moc": "power",
    "jednostka": "power_unit",
    "jednostka mocy": "power_unit",
    "zabezpieczenie a": "protection_a",
    "sprzedawca": "seller_data",
    "dane sprzedawcy": "seller_data",
    "uwagi": "notes",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS machines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    machine_type TEXT NOT NULL,
    serial_number TEXT NOT NULL,
    machine_category TEXT,
    commissioning_date TEXT,
    warranty_period_months INTEGER,
    warranty_end_date TEXT,
    power REAL,
    power_unit TEXT,
    protection_a TEXT,
    seller_data TEXT,
    notes TEXT,
    UNIQUE(name, machine_type, serial_number)
);

CREATE TABLE IF NOT EXISTS repair_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    record_year INTEGER NOT NULL,
    record_week INTEGER NOT NULL,
    week_sequence INTEGER NOT NULL,
    machine_id INTEGER NOT NULL,
    event_type TEXT NOT NULL,
    description TEXT NOT NULL,
    cause TEXT,
    repair_method TEXT,
    replaced_parts TEXT,
    participants TEXT,
    occurrence_at TEXT NOT NULL,
    repaired_at TEXT,
    machine_mtg TEXT,
    downtime_hours REAL NOT NULL DEFAULT 0,
    labor_cost REAL NOT NULL DEFAULT 0,
    material_cost REAL NOT NULL DEFAULT 0,
    spare_parts_cost REAL NOT NULL DEFAULT 0,
    service_cost REAL NOT NULL DEFAULT 0,
    total_cost REAL NOT NULL DEFAULT 0,
    notes TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(machine_id) REFERENCES machines(id)
);

CREATE TABLE IF NOT EXISTS attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    record_id INTEGER NOT NULL,
    filename TEXT NOT NULL,
    original_name TEXT NOT NULL,
    content_type TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(record_id) REFERENCES repair_records(id)
);

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
);

CREATE TABLE IF NOT EXISTS user_activity_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    username TEXT,
    action TEXT NOT NULL,
    target_type TEXT,
    target_id INTEGER,
    details TEXT,
    ip_address TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS scheduled_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    machine_id INTEGER NOT NULL,
    created_by_user_id INTEGER,
    repair_record_id INTEGER,
    event_type TEXT NOT NULL,
    description TEXT NOT NULL,
    cause TEXT,
    repair_method TEXT,
    replaced_parts TEXT,
    notes TEXT,
    planned_at TEXT NOT NULL,
    repeat_interval INTEGER NOT NULL DEFAULT 0,
    repeat_unit TEXT NOT NULL DEFAULT 'none',
    status TEXT NOT NULL DEFAULT 'planned',
    recurrence_next_event_id INTEGER,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(machine_id) REFERENCES machines(id),
    FOREIGN KEY(created_by_user_id) REFERENCES users(id)
);
"""


def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    for source, target in {
        "\n": " ",
        "\r": " ",
        "/": " ",
        ".": " ",
        ",": " ",
        "[": "",
        "]": "",
        "(": "",
        ")": "",
    }.items():
        text = text.replace(source, target)
    replacements = str.maketrans(
        "ąćęłńóśżź",
        "acelnoszz",
    )
    text = text.translate(replacements)
    return " ".join(text.split())


def parse_datetime(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    for fmt in (
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_datetime(value: str | None) -> str:
    if not value:
        return ""
    parsed = parse_datetime(value)
    if not parsed:
        return value
    return parsed.strftime("%d-%m-%Y %H:%M")


def format_date(value: str | None) -> str:
    if not value:
        return ""
    parsed = parse_datetime(value)
    if not parsed:
        return value
    return parsed.strftime("%d-%m-%Y")


def format_activity_datetime(value: str | None) -> str:
    """Display UTC audit timestamps in the application's local timezone."""
    if not value:
        return ""
    try:
        parsed = datetime.fromisoformat(str(value))
    except ValueError:
        return str(value)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed.astimezone().strftime("%d-%m-%Y %H:%M")


def as_float(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "").replace(",", ".")
    return float(text) if text else 0.0


def row_value(values: tuple[Any, ...], header_map: dict[str, int], header: str, default: Any = "") -> Any:
    index = header_map.get(header)
    if index is None or index < 0 or index >= len(values):
        return default
    return values[index]


def calculate_warranty_end_date(
    commissioning_date: str | None,
    warranty_period_months: int | None,
) -> str | None:
    if not commissioning_date or warranty_period_months is None:
        return None
    start = date.fromisoformat(commissioning_date)
    target_month = start.year * 12 + start.month - 1 + warranty_period_months
    year, month_index = divmod(target_month, 12)
    if not 1 <= year <= 9999:
        raise ValueError("Wyliczona data końca gwarancji wykracza poza obsługiwany zakres.")
    month = month_index + 1
    day = min(start.day, calendar.monthrange(year, month)[1])
    return date(year, month, day).isoformat()


def _machine_optional_text(value: Any) -> str | None:
    if value in (None, ""):
        return None
    text = str(value).strip()
    return text or None


def _machine_optional_date(value: Any, label: str) -> str | None:
    if value in (None, ""):
        return None
    parsed = parse_datetime(value)
    if parsed is None:
        raise ValueError(f"Nieprawidłowa wartość w polu „{label}”: {value}.")
    return parsed.date().isoformat()


def _machine_optional_number(value: Any, label: str) -> float | None:
    if value in (None, ""):
        return None
    try:
        number = float(str(value).strip().replace(" ", "").replace(",", "."))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Nieprawidłowa wartość liczbowa w polu „{label}”: {value}.") from exc
    if not math.isfinite(number):
        raise ValueError(f"Nieprawidłowa wartość liczbowa w polu „{label}”: {value}.")
    return number


def _machine_optional_months(value: Any, label: str) -> int | None:
    text = _machine_optional_text(value)
    if text is None or text.casefold() in {"-", "nd", "n/a", "brak"}:
        return None
    months = _machine_optional_number(text, label)
    if months is None:
        return None
    if months < 0 or not months.is_integer():
        raise ValueError(f"{label} musi być nieujemną liczbą całkowitą miesięcy.")
    return int(months)


def _machine_form_data(form: Any) -> dict[str, Any]:
    name = str(form.get("name", "")).strip()
    machine_type = str(form.get("machine_type", "")).strip()
    serial_number = str(form.get("serial_number", "")).strip()
    if not all((name, machine_type, serial_number)):
        raise ValueError("Nazwa, typ i numer seryjny są wymagane.")

    machine_data = {
        "name": name,
        "machine_type": machine_type,
        "serial_number": serial_number,
        "machine_category": _machine_optional_text(form.get("machine_category")),
        "commissioning_date": _machine_optional_date(
            form.get("commissioning_date"), "Data uruchomienia"
        ),
        "warranty_period_months": _machine_optional_months(
            form.get("warranty_period_months"), "Okres gwarancji"
        ),
        "power": _machine_optional_number(form.get("power"), "Moc"),
        "power_unit": _machine_optional_text(form.get("power_unit")),
        "protection_a": _machine_optional_text(form.get("protection_a")),
        "seller_data": _machine_optional_text(form.get("seller_data")),
        "notes": _machine_optional_text(form.get("notes")),
    }
    machine_data["warranty_end_date"] = calculate_warranty_end_date(
        machine_data["commissioning_date"], machine_data["warranty_period_months"]
    )
    return machine_data


def _machine_match_key(value: Any) -> str:
    normalized = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(character for character in normalized if character.isalnum())


def _numeric_machine_serial_key(value: Any) -> str | None:
    key = _machine_match_key(value)
    if not key.isdigit():
        return None
    return key.lstrip("0") or "0"


def import_machine_workbook(
    binary_content: bytes,
    *,
    create_missing: bool = True,
) -> tuple[int, int, int, list[str]]:
    workbook = load_workbook(io.BytesIO(binary_content), data_only=True, read_only=True)
    try:
        rows = list(workbook.active.iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        raise ValueError("Plik Excel jest pusty.")

    header_row_index = None
    header_map: dict[str, int] = {}
    for row_index, row in enumerate(rows[:10]):
        candidate: dict[str, int] = {}
        for column_index, header in enumerate(row):
            if header is None:
                continue
            canonical = MACHINE_IMPORT_HEADER_ALIASES.get(normalize_header(header))
            if canonical:
                candidate.setdefault(canonical, column_index)
        if {"name", "machine_type", "serial_number"}.issubset(candidate):
            header_row_index = row_index
            header_map = candidate
            break
    if header_row_index is None:
        raise ValueError(
            "Nie rozpoznano nagłówków tabeli maszyn. Wymagane są kolumny Firma, Nazwa i Nr fab."
        )

    field_labels = {
        "machine_category": "Rodzaj",
        "commissioning_date": "Data uruchomienia",
        "warranty_period_months": "Okres gwarancji",
        "power": "Moc",
        "power_unit": "Jednostka mocy",
        "protection_a": "Zabezpieczenie [A]",
        "seller_data": "Dane sprzedawcy",
        "notes": "Uwagi",
    }
    prepared_rows: list[tuple[int, dict[str, Any]]] = []
    skipped = 0
    warnings: list[str] = []
    for row_index, values in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        if not any(value not in (None, "") for value in values):
            continue
        row_data: dict[str, Any] = {
            "name": _machine_optional_text(row_value(values, header_map, "name")),
            "machine_type": _machine_optional_text(
                row_value(values, header_map, "machine_type")
            ),
            "serial_number": _machine_optional_text(
                row_value(values, header_map, "serial_number")
            ),
        }
        for field, label in field_labels.items():
            column_index = header_map.get(field)
            value = values[column_index] if column_index is not None and column_index < len(values) else None
            try:
                if field == "commissioning_date":
                    row_data[field] = _machine_optional_date(value, label)
                elif field == "warranty_period_months":
                    row_data[field] = _machine_optional_months(value, label)
                elif field == "power":
                    row_data[field] = _machine_optional_number(value, label)
                else:
                    row_data[field] = _machine_optional_text(value)
            except ValueError as exc:
                raise ValueError(f"Wiersz {row_index}: {exc}") from exc
        if not row_data["name"] or not row_data["machine_type"]:
            skipped += 1
            warnings.append(f"Wiersz {row_index}: brak firmy lub modelu.")
            continue
        row_data["serial_number"] = row_data["serial_number"] or "-"
        prepared_rows.append((row_index, row_data))

    db = get_db()
    machines = db.execute(
        "SELECT id, name, machine_type, serial_number, commissioning_date, "
        "warranty_period_months FROM machines"
    ).fetchall()
    serial_index: dict[str, list[sqlite3.Row]] = {}
    numeric_serial_index: dict[str, list[sqlite3.Row]] = {}
    identity_index: dict[tuple[str, str], list[sqlite3.Row]] = {}

    def index_machine(machine: sqlite3.Row) -> None:
        serial_key = _machine_match_key(machine["serial_number"])
        if not is_placeholder_serial(str(machine["serial_number"] or "")):
            serial_index.setdefault(serial_key, []).append(machine)
            numeric_serial_key = _numeric_machine_serial_key(machine["serial_number"])
            if numeric_serial_key is not None:
                numeric_serial_index.setdefault(numeric_serial_key, []).append(machine)
        identity_key = (
            _machine_match_key(machine["name"]),
            _machine_match_key(machine["machine_type"]),
        )
        identity_index.setdefault(identity_key, []).append(machine)

    for machine in machines:
        index_machine(machine)
    fields = (
        "machine_category",
        "commissioning_date",
        "warranty_period_months",
        "warranty_end_date",
        "power",
        "power_unit",
        "protection_a",
        "seller_data",
        "notes",
    )
    created = updated = 0
    with db:
        for row_number, row_data in prepared_rows:
            serial_key = _machine_match_key(row_data["serial_number"])
            identity_key = (
                _machine_match_key(row_data["name"]),
                _machine_match_key(row_data["machine_type"]),
            )
            match = None
            ambiguity = None
            if not is_placeholder_serial(row_data["serial_number"]):
                serial_matches = serial_index.get(serial_key, [])
                if not serial_matches:
                    numeric_serial_key = _numeric_machine_serial_key(row_data["serial_number"])
                    if numeric_serial_key is not None:
                        serial_matches = numeric_serial_index.get(numeric_serial_key, [])
                if len(serial_matches) > 1:
                    ambiguity = f"numer seryjny „{row_data['serial_number']}” pasuje do kilku maszyn"
                elif serial_matches:
                    match = serial_matches[0]
            if match is None:
                identity_matches = identity_index.get(identity_key, [])
                if ambiguity is None and len(identity_matches) > 1:
                    ambiguity = "firma i model pasują do kilku maszyn"
                elif ambiguity is None and identity_matches:
                    match = identity_matches[0]
            if ambiguity:
                skipped += 1
                warnings.append(f"Wiersz {row_number}: {ambiguity}.")
                continue

            if match is None:
                if not create_missing:
                    skipped += 1
                    warnings.append(f"Wiersz {row_number}: nie znaleziono maszyny w bazie.")
                    continue
                row_data["warranty_end_date"] = calculate_warranty_end_date(
                    row_data["commissioning_date"], row_data["warranty_period_months"]
                )
                cursor = db.execute(
                    """
                    INSERT INTO machines(
                        name, machine_type, serial_number, machine_category, commissioning_date,
                        warranty_period_months, warranty_end_date, power, power_unit,
                        protection_a, seller_data, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        row_data["name"],
                        row_data["machine_type"],
                        row_data["serial_number"],
                        *(row_data[field] for field in fields),
                    ),
                )
                new_machine = db.execute(
                    "SELECT id, name, machine_type, serial_number, commissioning_date, "
                    "warranty_period_months FROM machines WHERE id = ?",
                    (cursor.lastrowid,),
                ).fetchone()
                index_machine(new_machine)
                created += 1
                continue

            merged = {
                field: row_data[field] if row_data[field] is not None else match[field]
                for field in ("commissioning_date", "warranty_period_months")
            }
            warranty_end_date = calculate_warranty_end_date(
                merged["commissioning_date"], merged["warranty_period_months"]
            )
            db.execute(
                """
                UPDATE machines SET
                    machine_category = COALESCE(?, machine_category),
                    commissioning_date = ?,
                    warranty_period_months = ?,
                    warranty_end_date = ?,
                    power = COALESCE(?, power),
                    power_unit = COALESCE(?, power_unit),
                    protection_a = COALESCE(?, protection_a),
                    seller_data = COALESCE(?, seller_data),
                    notes = COALESCE(?, notes)
                WHERE id = ?
                """,
                (
                    row_data["machine_category"],
                    merged["commissioning_date"],
                    merged["warranty_period_months"],
                    warranty_end_date,
                    row_data["power"],
                    row_data["power_unit"],
                    row_data["protection_a"],
                    row_data["seller_data"],
                    row_data["notes"],
                    match["id"],
                ),
            )
            updated += 1
    return created, updated, skipped, warnings


def calculate_downtime_hours(occurrence_at: datetime, repaired_at: datetime | None) -> float:
    if not repaired_at or repaired_at <= occurrence_at:
        return 0.0
    return round((repaired_at - occurrence_at).total_seconds() / 3600, 2)


class _QueueZipWriter:
    """Non-seekable ZIP target that yields bounded chunks to an HTTP response."""

    def __init__(self, chunks: queue.Queue, cancelled: threading.Event):
        self.chunks = chunks
        self.cancelled = cancelled
        self.position = 0

    def write(self, data: bytes) -> int:
        written = len(data)
        for offset in range(0, written, 64 * 1024):
            chunk = data[offset:offset + 64 * 1024]
            while not self.cancelled.is_set():
                try:
                    self.chunks.put(chunk, timeout=0.1)
                    self.position += len(chunk)
                    break
                except queue.Full:
                    continue
            if self.cancelled.is_set():
                raise BrokenPipeError("Klient przerwał pobieranie kopii.")
        return written

    def tell(self) -> int:
        return self.position

    def seekable(self) -> bool:
        return False

    def flush(self) -> None:
        return None


def machine_key_from_form(form: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(form.get("machine_name", "")).strip(),
        str(form.get("machine_type", "")).strip(),
        str(form.get("serial_number", "")).strip(),
    )


AUTH_FILE = Path(__file__).parent / "auth.json"
CONFIG_FILE = Path(__file__).parent / "config.json"


def _load_config() -> dict[str, Any]:
    """Load configuration from config.json with fallback defaults."""
    defaults = {
        "DATABASE": "instance/repair_reports.sqlite3",
        "UPLOAD_DIR": "instance/uploads",
        "DEBUG": False,
        "HOST": "0.0.0.0",
        "PORT": 5001,
        "LABOR_RATE_PLN_PER_HOUR": 80,
        "MAX_UPLOAD_SIZE_MB": 500,
        "REQUIRE_LOGIN": True,
        "AUTH_FILE": "auth.json",
    }
    
    if CONFIG_FILE.exists():
        try:
            user_config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            # Filter out comment fields
            user_config = {k: v for k, v in user_config.items() if not k.startswith("_")}
            defaults.update(user_config)
        except (json.JSONDecodeError, OSError):
            pass
    
    return defaults


def _resolve_path(path_str: str) -> Path:
    """
    Resolve a path that could be:
    - Relative (instance/file.db) - relative to app root
    - Absolute Windows UNC (\\\\server\\share\\file.db)
    - Absolute Windows (C:\\path\\to\\file.db)
    - Absolute Unix (/mnt/nas/file.db)
    """
    if not path_str:
        return Path("instance/repair_reports.sqlite3")
    
    p = Path(path_str)
    
    # Handle forward-slash UNC paths (Windows network shares)
    if path_str.startswith("//"):
        # Convert // to \\ for Windows, keep // for Linux (network paths)
        import sys
        if sys.platform == "win32":
            unc_path = path_str.replace("/", "\\")
            return Path(unc_path)
        else:
            # Linux: treat as absolute path
            return Path(path_str)
    
    # Already absolute
    if p.is_absolute():
        return p
    
    # Relative path: resolve relative to app root
    return Path(__file__).parent / path_str


def _ensure_auth_file() -> dict:
    """Return auth data, creating auth.json with defaults on first run."""
    if AUTH_FILE.exists():
        try:
            return json.loads(AUTH_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    data = {
        "login": "admin",
        "password_hash": generate_password_hash("admin"),
        "secret_key": secrets.token_hex(32),
    }
    AUTH_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return data


def get_current_user() -> sqlite3.Row | None:
    user_id = session.get("user_id")
    if not user_id:
        return None
    db = get_db()
    return db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()


def current_user_id() -> int | None:
    user_id = session.get("user_id")
    return int(user_id) if user_id else None


def scheduled_event_is_visible_to_current_user(event: sqlite3.Row) -> bool:
    return (
        event["event_type"] != "Własne"
        or event["created_by_user_id"] == current_user_id()
    )


def user_has_permission(permission_name: str) -> bool:
    if session.get("role") == "admin":
        return True
    user = get_current_user()
    if not user:
        return False
    if bool(user["is_admin"]):
        return True
    permission_map = {
        "can_view_records": bool(user["can_view_records"]),
        "can_create_records": bool(user["can_create_records"]),
        "can_edit_records": bool(user["can_edit_records"]),
        "can_delete_records": bool(user["can_delete_records"]),
    }
    return permission_map.get(permission_name, False)


def log_user_activity(
    action: str,
    details: str = "",
    *,
    target_type: str | None = None,
    target_id: int | None = None,
    user_id: int | None = None,
    username: str | None = None,
) -> None:
    db = get_db()
    if user_id is None:
        user = get_current_user()
        if user:
            user_id = int(user["id"])
    if username is None:
        if user_id is not None:
            user_row = db.execute("SELECT username FROM users WHERE id = ?", (user_id,)).fetchone()
            if user_row is not None:
                username = str(user_row["username"])
        if username is None:
            username = str(session.get("username") or ("admin" if session.get("role") == "admin" else "guest"))
    db.execute(
        """
        INSERT INTO user_activity_log(
            user_id, username, action, target_type, target_id, details, ip_address, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            username,
            action,
            target_type,
            target_id,
            details,
            request.remote_addr,
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()


def ensure_default_admin_user() -> None:
    db = get_db()
    auth = _ensure_auth_file()
    username = str(auth.get("login", "admin")).strip() or "admin"
    existing = db.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
    if existing:
        return
    password_hash = auth.get("password_hash") or generate_password_hash("admin")
    timestamp = datetime.now(UTC).isoformat(timespec="seconds")
    db.execute(
        """
        INSERT INTO users(username, password_hash, is_admin, can_view_records, can_create_records,
                         can_edit_records, can_delete_records, created_at, updated_at)
        VALUES (?, ?, 1, 1, 1, 1, 1, ?, ?)
        """,
        (username, password_hash, timestamp, timestamp),
    )
    db.commit()


def get_user_history(
    limit: int = 20,
    offset: int = 0,
    sort_by: str = "created_at",
    sort_direction: str = "desc",
) -> list[sqlite3.Row]:
    db = get_db()
    sort_columns = {
        "created_at": "created_at",
        "username": "username",
        "ip_address": "ip_address",
        "action": "action",
        "details": "details",
    }
    column = sort_columns.get(sort_by, "created_at")
    direction = "ASC" if sort_direction == "asc" else "DESC"
    tie_breaker_direction = "ASC" if sort_direction == "asc" else "DESC"
    return list(
        db.execute(
            f"SELECT * FROM user_activity_log ORDER BY {column} {direction}, created_at DESC, id {tie_breaker_direction} LIMIT ? OFFSET ?",
            (limit, offset),
        ).fetchall()
    )


def create_or_update_user(user_data: dict[str, Any], user_id: int | None = None) -> int:
    db = get_db()
    username = str(user_data.get("username", "")).strip()
    if not username:
        raise ValueError("Nazwa użytkownika jest wymagana.")
    if user_id is None:
        existing = db.execute("SELECT id FROM users WHERE lower(username) = lower(?)", (username,)).fetchone()
        if existing:
            raise ValueError("Użytkownik o tej nazwie już istnieje.")
    else:
        existing = db.execute("SELECT id FROM users WHERE id = ?", (user_id,)).fetchone()
        if not existing:
            raise ValueError("Użytkownik nie istnieje.")
        conflict = db.execute(
            "SELECT id FROM users WHERE lower(username) = lower(?) AND id != ?",
            (username, user_id),
        ).fetchone()
        if conflict:
            raise ValueError("Użytkownik o tej nazwie już istnieje.")

    now = datetime.now(UTC).isoformat(timespec="seconds")
    is_admin = 1 if str(user_data.get("is_admin", "0")).lower() in {"1", "true", "yes", "on"} else 0
    permissions = {
        "can_view_records": 1 if str(user_data.get("can_view_records", "0")).lower() in {"1", "true", "yes", "on"} else 0,
        "can_create_records": 1 if str(user_data.get("can_create_records", "0")).lower() in {"1", "true", "yes", "on"} else 0,
        "can_edit_records": 1 if str(user_data.get("can_edit_records", "0")).lower() in {"1", "true", "yes", "on"} else 0,
        "can_delete_records": 1 if str(user_data.get("can_delete_records", "0")).lower() in {"1", "true", "yes", "on"} else 0,
    }
    if is_admin:
        permissions = {key: 1 for key in permissions}

    if user_id is None:
        password_hash = generate_password_hash(str(user_data.get("password", "admin123")).strip() or "admin123")
        cursor = db.execute(
            """
            INSERT INTO users(username, password_hash, is_admin, can_view_records, can_create_records,
                             can_edit_records, can_delete_records, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (username, password_hash, is_admin, permissions["can_view_records"], permissions["can_create_records"],
             permissions["can_edit_records"], permissions["can_delete_records"], now, now),
        )
        user_id = int(cursor.lastrowid)
        log_user_activity("create_user", f"Utworzono użytkownika {username}", target_type="user", target_id=user_id, user_id=session.get("user_id"))
    else:
        password = str(user_data.get("password", "")).strip()
        sql = """
            UPDATE users
            SET username = ?, is_admin = ?, can_view_records = ?, can_create_records = ?,
                can_edit_records = ?, can_delete_records = ?, updated_at = ?
        """
        params = [username, is_admin, permissions["can_view_records"], permissions["can_create_records"],
                  permissions["can_edit_records"], permissions["can_delete_records"], now]
        if password:
            sql += ", password_hash = ?"
            params.append(generate_password_hash(password))
        sql += " WHERE id = ?"
        params.append(user_id)
        db.execute(sql, params)
        log_user_activity("update_user", f"Zaktualizowano użytkownika {username}", target_type="user", target_id=user_id, user_id=session.get("user_id"))
    db.commit()
    return int(user_id)


def delete_user(user_id: int) -> None:
    db = get_db()
    if int(user_id) == int(session.get("user_id") or -1):
        raise ValueError("Nie można usunąć własnego konta.")
    db.execute("DELETE FROM user_activity_log WHERE user_id = ?", (user_id,))
    db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    db.commit()
    log_user_activity("delete_user", "Usunięto użytkownika", target_type="user", target_id=user_id, user_id=session.get("user_id"))


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    app = Flask(__name__)
    app.jinja_env.filters["activity_datetime"] = format_activity_datetime
    app.jinja_env.filters["datetime"] = format_datetime
    config = _load_config()
    
    # Resolve database and upload paths
    db_path = _resolve_path(config["DATABASE"])
    upload_dir = _resolve_path(config["UPLOAD_DIR"])
    
    auth_data = _ensure_auth_file()
    app.config.from_mapping(
        SECRET_KEY=auth_data.get("secret_key", "dev"),
        DATABASE=str(db_path),
        UPLOAD_DIR=str(upload_dir),
        MAX_CONTENT_LENGTH=config["MAX_UPLOAD_SIZE_MB"] * 1024 * 1024,
    )
    
    # Store config values for access in routes
    app.config["LABOR_RATE_PLN_PER_HOUR"] = config["LABOR_RATE_PLN_PER_HOUR"]
    app.config["REQUIRE_LOGIN"] = config["REQUIRE_LOGIN"]
    
    if test_config:
        app.config.update(test_config)

    # Ensure directories exist
    db_path.parent.mkdir(parents=True, exist_ok=True)
    upload_dir.mkdir(parents=True, exist_ok=True)

    @app.before_request
    def check_auth():
        if app.config.get("TESTING"):
            return
        if not app.config.get("REQUIRE_LOGIN"):
            if not session.get("logged_in"):
                session["logged_in"] = True
                session["role"] = "guest"
            return
        exempt = {"login", "logout", "static"}
        if request.endpoint in exempt or request.endpoint is None:
            return
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        if session.get("role") in {"admin", "user"} and session.get("user_id"):
            user = get_current_user()
            if user is None:
                session.clear()
                return redirect(url_for("login"))
            if not bool(user["can_view_records"]) and request.endpoint not in {"logout", "login"}:
                flash("Brak uprawnień do przeglądania rekordów.", "error")
                return redirect(url_for("logout"))

    @app.before_request
    def open_db_connection() -> None:
        get_db()

    @app.teardown_appcontext
    def close_db(_: BaseException | None) -> None:
        connection = g.pop("db", None)
        if connection is not None:
            connection.close()

    @app.context_processor
    def inject_role():
        current_user = get_current_user()
        username = session.get("username")
        if not username and current_user:
            username = current_user["username"]
        return {
            "user_role": session.get("role", "guest"),
            "current_user": current_user,
            "current_username": username or "Gościu",
        }

    with app.app_context():
        init_db()
        ensure_default_admin_user()

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        if request.method == "POST":
            login_val = request.form.get("login", "").strip()
            pw_val = request.form.get("password", "")
            if not login_val and not pw_val:
                session["logged_in"] = True
                session["role"] = "guest"
                return redirect(url_for("index"))

            db = get_db()
            user = db.execute("SELECT * FROM users WHERE lower(username) = lower(?)", (login_val,)).fetchone()
            if user is not None and check_password_hash(user["password_hash"], pw_val):
                session["logged_in"] = True
                session["role"] = "admin" if user["is_admin"] else "user"
                session["user_id"] = int(user["id"])
                session["username"] = user["username"]
                log_user_activity("login", f"Użytkownik {user['username']} zalogował się.", user_id=int(user["id"]))
                return redirect(url_for("index"))

            auth = json.loads(AUTH_FILE.read_text(encoding="utf-8"))
            ok_login = login_val == auth.get("login", "admin")
            ok_pw = check_password_hash(auth["password_hash"], pw_val)
            if ok_login and ok_pw:
                session["logged_in"] = True
                session["role"] = "admin"
                session["username"] = auth.get("login", "admin")
                session["legacy_admin"] = True
                log_user_activity("login", "Zalogowano przez legacy admin.", username=auth.get("login", "admin"))
                return redirect(url_for("index"))
            error = "Nieprawid\u0142owy login lub has\u0142o."
        return render_template("login.html", error=error)

    @app.route("/logout")
    def logout():
        if session.get("user_id"):
            log_user_activity("logout", "Wylogowanie użytkownika.", user_id=int(session["user_id"]))
        elif session.get("role") == "admin":
            log_user_activity("logout", "Wylogowanie użytkownika.", username=str(session.get("username") or "admin"))
        session.clear()
        return redirect(url_for("login"))

    @app.route("/change-password", methods=["GET", "POST"])
    def change_password():
        if request.method == "POST":
            current = request.form.get("current_password", "")
            new_pw = request.form.get("new_password", "")
            confirm = request.form.get("confirm_password", "")
            user = get_current_user()
            if user is not None:
                if not check_password_hash(user["password_hash"], current):
                    flash("Nieprawid\u0142owe obecne has\u0142o.", "error")
                elif len(new_pw) < 4:
                    flash("Nowe has\u0142o musi mie\u0107 co najmniej 4 znaki.", "error")
                elif new_pw != confirm:
                    flash("Has\u0142a nie s\u0105 identyczne.", "error")
                else:
                    db = get_db()
                    db.execute(
                        "UPDATE users SET password_hash = ?, updated_at = ? WHERE id = ?",
                        (generate_password_hash(new_pw), datetime.now(UTC).isoformat(timespec="seconds"), user["id"]),
                    )
                    db.commit()
                    auth = json.loads(AUTH_FILE.read_text(encoding="utf-8"))
                    if auth.get("login") == user["username"]:
                        auth["password_hash"] = generate_password_hash(new_pw)
                        AUTH_FILE.write_text(json.dumps(auth, indent=2), encoding="utf-8")
                    log_user_activity("change_password", f"Zmiana hasła użytkownika {user['username']}.", user_id=int(user["id"]))
                    flash("Has\u0142o zosta\u0142o zmienione.", "success")
                    return redirect(url_for("index"))
            else:
                auth = json.loads(AUTH_FILE.read_text(encoding="utf-8"))
                if not check_password_hash(auth["password_hash"], current):
                    flash("Nieprawid\u0142owe obecne has\u0142o.", "error")
                elif len(new_pw) < 4:
                    flash("Nowe has\u0142o musi mie\u0107 co najmniej 4 znaki.", "error")
                elif new_pw != confirm:
                    flash("Has\u0142a nie s\u0105 identyczne.", "error")
                else:
                    auth["password_hash"] = generate_password_hash(new_pw)
                    AUTH_FILE.write_text(json.dumps(auth, indent=2), encoding="utf-8")
                    flash("Has\u0142o zosta\u0142o zmienione.", "success")
                    return redirect(url_for("index"))
        return render_template("change_password.html")

    @app.route("/admin/users", methods=["GET", "POST"])
    def admin_users():
        if not session.get("role") == "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        if request.method == "POST":
            user_id = request.form.get("user_id")
            payload = dict(request.form)
            try:
                if user_id:
                    create_or_update_user(payload, int(user_id))
                    flash("Użytkownik został zaktualizowany.", "success")
                else:
                    create_or_update_user(payload)
                    flash("Użytkownik został utworzony.", "success")
            except ValueError as exc:
                flash(str(exc), "error")
            return redirect(url_for("admin_users"))

        db = get_db()
        users = db.execute("SELECT * FROM users ORDER BY id").fetchall()
        edit_user = None
        edit_user_id = request.args.get("edit", type=int)
        if edit_user_id:
            edit_user = db.execute("SELECT * FROM users WHERE id = ?", (edit_user_id,)).fetchone()
        activity_page = max(1, request.args.get("page", 1, type=int) or 1)
        activity_sort = request.args.get("sort", "created_at")
        if activity_sort not in {"created_at", "username", "ip_address", "action", "details"}:
            activity_sort = "created_at"
        activity_direction = request.args.get("direction", "desc")
        if activity_direction not in {"asc", "desc"}:
            activity_direction = "desc"
        activity_per_page = 20
        activity_total = db.execute("SELECT COUNT(*) AS total FROM user_activity_log").fetchone()["total"]
        activity_pages = max(1, (int(activity_total) + activity_per_page - 1) // activity_per_page)
        activity_page = min(activity_page, activity_pages)
        return render_template(
            "admin_users.html",
            users=users,
            edit_user=edit_user,
            recent_activity=get_user_history(
                activity_per_page,
                (activity_page - 1) * activity_per_page,
                activity_sort,
                activity_direction,
            ),
            activity_page=activity_page,
            activity_pages=activity_pages,
            activity_sort=activity_sort,
            activity_direction=activity_direction,
        )

    @app.route("/admin/machines")
    def admin_machines():
        can_manage_machines = session.get("role") == "admin"
        can_view_machines = can_manage_machines or any(
            user_has_permission(permission)
            for permission in ("can_create_records", "can_edit_records", "can_delete_records")
        )
        if not can_view_machines:
            flash("Brak uprawnień do przeglądania listy maszyn.", "error")
            return redirect(url_for("index"))

        db = get_db()
        machines = db.execute(
            """
            SELECT m.*, COUNT(r.id) AS record_count
            FROM machines m
            LEFT JOIN repair_records r ON r.machine_id = m.id
            GROUP BY m.id
            ORDER BY m.name, m.machine_type, m.serial_number
            """
        ).fetchall()
        edit_machine_id = request.args.get("edit", type=int)
        edit_machine = None
        if edit_machine_id and can_manage_machines:
            edit_machine = db.execute(
                "SELECT * FROM machines WHERE id = ?", (edit_machine_id,)
            ).fetchone()
        return render_template(
            "admin_machines.html",
            machines=machines,
            edit_machine=edit_machine,
            can_manage_machines=can_manage_machines,
        )

    @app.route("/admin/machines/new", methods=["POST"])
    def admin_create_machine():
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        try:
            machine = _machine_form_data(request.form)
            db = get_db()
            cursor = db.execute(
                """
                INSERT INTO machines(
                    name, machine_type, serial_number, machine_category, commissioning_date,
                    warranty_period_months, warranty_end_date, power, power_unit,
                    protection_a, seller_data, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    machine["name"],
                    machine["machine_type"],
                    machine["serial_number"],
                    machine["machine_category"],
                    machine["commissioning_date"],
                    machine["warranty_period_months"],
                    machine["warranty_end_date"],
                    machine["power"],
                    machine["power_unit"],
                    machine["protection_a"],
                    machine["seller_data"],
                    machine["notes"],
                ),
            )
            db.commit()
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("admin_machines", new=1))
        except sqlite3.IntegrityError:
            get_db().rollback()
            flash("Maszyna o takich danych lub numerze seryjnym już istnieje.", "error")
            return redirect(url_for("admin_machines", new=1))

        machine_id = int(cursor.lastrowid)
        log_user_activity(
            "create_machine",
            f"Dodano maszynę {machine['name']} (ID {machine_id}).",
            target_type="machine",
            target_id=machine_id,
        )
        flash("Maszyna została dodana.", "success")
        return redirect(url_for("admin_machines"))

    @app.route("/admin/machines/<int:machine_id>/edit", methods=["POST"])
    def admin_edit_machine(machine_id: int):
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        try:
            machine = _machine_form_data(request.form)
            db = get_db()
            cursor = db.execute(
                """
                UPDATE machines
                SET name = ?, machine_type = ?, serial_number = ?, machine_category = ?,
                    commissioning_date = ?, warranty_period_months = ?, warranty_end_date = ?,
                    power = ?, power_unit = ?, protection_a = ?, seller_data = ?, notes = ?
                WHERE id = ?
                """,
                (
                    machine["name"],
                    machine["machine_type"],
                    machine["serial_number"],
                    machine["machine_category"],
                    machine["commissioning_date"],
                    machine["warranty_period_months"],
                    machine["warranty_end_date"],
                    machine["power"],
                    machine["power_unit"],
                    machine["protection_a"],
                    machine["seller_data"],
                    machine["notes"],
                    machine_id,
                ),
            )
            if cursor.rowcount == 0:
                flash("Maszyna nie istnieje.", "error")
                return redirect(url_for("admin_machines"))
            db.commit()
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("admin_machines", edit=machine_id))
        except sqlite3.IntegrityError:
            get_db().rollback()
            flash("Maszyna o takich danych lub numerze seryjnym już istnieje.", "error")
            return redirect(url_for("admin_machines", edit=machine_id))

        log_user_activity(
            "update_machine",
            f"Zmieniono dane maszyny {machine['name']} (ID {machine_id}).",
            target_type="machine",
            target_id=machine_id,
        )
        flash("Dane maszyny zostały zaktualizowane.", "success")
        return redirect(url_for("admin_machines"))

    @app.route("/admin/machines/export")
    def admin_export_machines():
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        db = get_db()
        machines = db.execute(
            """
            SELECT name, machine_type, serial_number, machine_category, commissioning_date,
                   warranty_period_months, warranty_end_date, power, power_unit, protection_a,
                   seller_data, notes
            FROM machines
            ORDER BY name, machine_type, serial_number
            """
        ).fetchall()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "MASZYNY"
        sheet.append(MACHINE_EXPORT_COLUMNS)
        for machine in machines:
            commissioning_date = (
                date.fromisoformat(machine["commissioning_date"])
                if machine["commissioning_date"]
                else None
            )
            warranty_end_date = (
                date.fromisoformat(machine["warranty_end_date"])
                if machine["warranty_end_date"]
                else None
            )
            sheet.append(
                [
                    machine["name"],
                    machine["machine_type"],
                    machine["serial_number"],
                    machine["machine_category"] or "",
                    commissioning_date,
                    machine["warranty_period_months"],
                    warranty_end_date,
                    machine["power"],
                    machine["power_unit"] or "",
                    machine["protection_a"],
                    machine["seller_data"] or "",
                    machine["notes"] or "",
                ]
            )
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column, width in enumerate((24, 28, 22, 24, 18, 22, 20, 12, 16, 20, 32, 40), start=1):
            sheet.column_dimensions[chr(64 + column)].width = width
        for row in sheet.iter_rows(min_row=2):
            row[4].number_format = "DD.MM.YYYY"
            row[6].number_format = "DD.MM.YYYY"

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        log_user_activity(
            "export_machines",
            "Wyeksportowano listę maszyn do pliku Excel.",
            target_type="machines",
        )
        return send_file(
            output,
            as_attachment=True,
            download_name="maszyny.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.route("/admin/machines/import", methods=["POST"])
    def admin_import_machines():
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        uploaded_file = request.files.get("file")
        if uploaded_file is None or not uploaded_file.filename:
            flash("Wybierz plik Excel do importu.", "error")
            return redirect(url_for("data_page"))
        try:
            created, updated, skipped, warnings = import_machine_workbook(uploaded_file.read())
        except (ValueError, zipfile.BadZipFile, InvalidFileException, OSError) as exc:
            flash(f"Nie udało się zaimportować maszyn: {exc}", "error")
            return redirect(url_for("data_page"))

        log_user_activity(
            "import_machines",
            f"Zaimportowano listę maszyn: dodano {created}, zaktualizowano {updated}, pominięto {skipped}.",
            target_type="machines",
        )
        flash(
            f"Import zakończony. Dodano: {created}, zaktualizowano: {updated}, pominięto: {skipped}. "
            "Dopasowane maszyny zachowują swoje dane identyfikacyjne; puste komórki nie nadpisują danych."
            + (f" Szczegóły pominięć: {'; '.join(warnings[:5])}" if warnings else ""),
            "success",
        )
        return redirect(url_for("data_page"))

    @app.route("/admin/machines/<int:machine_id>/delete", methods=["POST"])
    def admin_delete_machine(machine_id: int):
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        db = get_db()
        machine = db.execute(
            "SELECT name FROM machines WHERE id = ?", (machine_id,)
        ).fetchone()
        if machine is None:
            flash("Maszyna nie istnieje.", "error")
            return redirect(url_for("admin_machines"))

        record_ids = [
            int(row["id"])
            for row in db.execute(
                "SELECT id FROM repair_records WHERE machine_id = ?", (machine_id,)
            ).fetchall()
        ]
        import shutil

        uploads = Path(current_app_uploads())
        for record_id in record_ids:
            record_upload_dir = uploads / str(record_id)
            if record_upload_dir.exists():
                shutil.rmtree(record_upload_dir)

        with db:
            db.execute(
                "DELETE FROM attachments WHERE record_id IN "
                "(SELECT id FROM repair_records WHERE machine_id = ?)",
                (machine_id,),
            )
            db.execute("DELETE FROM repair_records WHERE machine_id = ?", (machine_id,))
            db.execute("DELETE FROM machines WHERE id = ?", (machine_id,))

        log_user_activity(
            "delete_machine",
            f"Usunięto maszynę {machine['name']} (ID {machine_id}) wraz z {len(record_ids)} wpisami.",
            target_type="machine",
            target_id=machine_id,
        )
        flash(
            f"Usunięto maszynę i {len(record_ids)} powiązanych wpisów.", "success"
        )
        return redirect(url_for("admin_machines"))

    @app.route("/admin/users/activity/export")
    def export_user_activity():
        if not session.get("role") == "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        db = get_db()
        activity = db.execute(
            "SELECT * FROM user_activity_log ORDER BY created_at DESC, id DESC"
        ).fetchall()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "REJESTR ZDARZEŃ"
        sheet.append(["Data", "Użytkownik", "Akcja", "Typ obiektu", "ID obiektu", "Szczegóły", "Adres IP"])
        for item in activity:
            sheet.append(
                [
                    format_activity_datetime(item["created_at"]),
                    item["username"] or "",
                    item["action"],
                    item["target_type"] or "",
                    item["target_id"] or "",
                    item["details"] or "",
                    item["ip_address"] or "",
                ]
            )
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        log_user_activity(
            "export_activity",
            "Wyeksportowano rejestr zdarzeń do pliku Excel.",
            target_type="activity_log",
        )
        return send_file(
            output,
            as_attachment=True,
            download_name="rejestr_zdarzen.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.route("/admin/users/<int:user_id>/delete", methods=["POST"])
    def admin_delete_user(user_id: int):
        if not session.get("role") == "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))
        try:
            delete_user(user_id)
            flash("Użytkownik został usunięty.", "success")
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("admin_users"))

    @app.route("/scheduled", methods=["GET", "POST"])
    @app.route("/planowanie", methods=["GET", "POST"])
    def scheduled_events():
        if not app.config.get("TESTING") and not user_has_permission("can_view_records"):
            flash("Brak uprawnień do przeglądania zaplanowanych zdarzeń.", "error")
            return redirect(url_for("index"))
        planned_events = get_scheduled_events("planned")
        completed_events = get_scheduled_events("done")
        calendar_view = request.args.get("calendar_view", "month")
        if calendar_view not in {"month", "year", "week"}:
            calendar_view = "month"

        current_date = datetime.now()
        current_iso_year, current_iso_week, _ = current_date.isocalendar()
        try:
            calendar_year = int(request.args.get(
                "year",
                current_iso_year if calendar_view == "week" else current_date.year,
            ))
        except (TypeError, ValueError):
            calendar_year = current_iso_year if calendar_view == "week" else current_date.year
        if not 1 <= calendar_year <= 9999:
            calendar_year = current_iso_year if calendar_view == "week" else current_date.year

        try:
            calendar_month = int(request.args.get("month", current_date.month))
        except (TypeError, ValueError):
            calendar_month = current_date.month
        if not 1 <= calendar_month <= 12:
            calendar_month = current_date.month

        month_names = (
            "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień",
        )
        events_by_date: dict[str, list[sqlite3.Row]] = {}
        dated_events = [
            (parse_datetime(event["planned_at"]), event)
            for event in (*planned_events, *completed_events)
        ]
        dated_events.sort(key=lambda item: item[0] or datetime.max)
        for planned_at, event in dated_events:
            if planned_at is not None:
                events_by_date.setdefault(planned_at.date().isoformat(), []).append(event)

        calendar_months = []
        calendar_week = current_iso_week
        calendar_week_value = ""
        week_start = None
        weekly_days = []
        if calendar_view == "week":
            requested_iso_week = request.args.get("iso_week", "")
            try:
                iso_year_text, iso_week_text = requested_iso_week.split("-W")
                calendar_year = int(iso_year_text)
                calendar_week = int(iso_week_text)
                week_start = datetime.fromisocalendar(calendar_year, calendar_week, 1).date()
            except (TypeError, ValueError):
                calendar_year = current_iso_year
                calendar_week = current_iso_week
                week_start = current_date.date() - timedelta(days=current_date.weekday())
            calendar_week_value = f"{calendar_year:04d}-W{calendar_week:02d}"
            for day_offset in range(7):
                try:
                    day = week_start + timedelta(days=day_offset)
                except OverflowError:
                    day = None
                hourly_events: list[list[dict[str, Any]]] = [[] for _ in range(24)]
                if day is not None:
                    for event in events_by_date.get(day.isoformat(), []):
                        planned_at = parse_datetime(event["planned_at"])
                        if planned_at is not None:
                            hourly_events[planned_at.hour].append({
                                "event": event,
                                "time": planned_at.strftime("%H:%M"),
                            })
                weekly_days.append({
                    "date": day,
                    "is_today": day == current_date.date(),
                    "hours": hourly_events,
                })
            try:
                previous_week_start = week_start - timedelta(days=7)
            except OverflowError:
                previous_week_start = None
            try:
                next_week_start = week_start + timedelta(days=7)
            except OverflowError:
                next_week_start = None
            previous_url = None
            if previous_week_start is not None:
                previous_iso_year, previous_iso_week, _ = previous_week_start.isocalendar()
                previous_url = url_for(
                    "scheduled_events",
                    calendar_view="week",
                    iso_week=f"{previous_iso_year:04d}-W{previous_iso_week:02d}",
                )
            next_url = None
            if next_week_start is not None and next_week_start.year <= 9999:
                next_iso_year, next_iso_week, _ = next_week_start.isocalendar()
                next_url = url_for(
                    "scheduled_events",
                    calendar_view="week",
                    iso_week=f"{next_iso_year:04d}-W{next_iso_week:02d}",
                )
            calendar_heading = f"Tydzień {calendar_week}, {calendar_year}"

        month_numbers = range(1, 13) if calendar_view == "year" else (calendar_month,)
        calendar_generator = calendar.Calendar(firstweekday=calendar.MONDAY)
        if calendar_view != "week":
            for month_number in month_numbers:
                weeks = []
                for week in calendar_generator.monthdatescalendar(calendar_year, month_number):
                    weeks.append({
                        "number": week[0].isocalendar().week,
                        "days": [
                            {
                                "date": day,
                                "in_month": day.month == month_number,
                                "is_today": day == current_date.date(),
                                "events": events_by_date.get(day.isoformat(), []) if day.month == month_number else [],
                            }
                            for day in week
                        ],
                    })
                calendar_months.append({
                    "number": month_number,
                    "name": month_names[month_number - 1],
                    "weeks": weeks,
                })

            if calendar_view == "year":
                previous_url = (
                    url_for("scheduled_events", calendar_view="year", year=calendar_year - 1, month=calendar_month)
                    if calendar_year > 1 else None
                )
                next_url = (
                    url_for("scheduled_events", calendar_view="year", year=calendar_year + 1, month=calendar_month)
                    if calendar_year < 9999 else None
                )
                calendar_heading = str(calendar_year)
            else:
                previous_year, previous_month = (
                    (calendar_year - 1, 12) if calendar_month == 1 else (calendar_year, calendar_month - 1)
                )
                next_year, next_month = (
                    (calendar_year + 1, 1) if calendar_month == 12 else (calendar_year, calendar_month + 1)
                )
                previous_url = (
                    url_for("scheduled_events", calendar_view="month", year=previous_year, month=previous_month)
                    if previous_year >= 1 else None
                )
                next_url = (
                    url_for("scheduled_events", calendar_view="month", year=next_year, month=next_month)
                    if next_year <= 9999 else None
                )
                calendar_heading = f"{month_names[calendar_month - 1]} {calendar_year}"

        return render_template(
            "scheduled_events.html",
            scheduled_events=planned_events,
            completed_events=completed_events,
            planned_count=len(planned_events),
            completed_count=len(completed_events),
            event_types=SCHEDULED_EVENT_TYPES,
            calendar_view=calendar_view,
            calendar_year=calendar_year,
            calendar_month=calendar_month,
            calendar_week=calendar_week,
            calendar_week_value=calendar_week_value,
            calendar_months=calendar_months,
            weekly_days=weekly_days,
            calendar_hours=range(24),
            calendar_heading=calendar_heading,
            calendar_weekdays=("Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"),
            previous_calendar_url=previous_url,
            next_calendar_url=next_url,
            today_calendar_url=(
                url_for(
                    "scheduled_events",
                    calendar_view="week",
                    iso_week=f"{current_iso_year:04d}-W{current_iso_week:02d}",
                )
                if calendar_view == "week"
                else url_for(
                    "scheduled_events",
                    calendar_view="month",
                    year=current_date.year,
                    month=current_date.month,
                )
                if calendar_view == "month"
                else url_for("scheduled_events", calendar_view="year", year=current_date.year)
            ),
            month_names=month_names,
        )

    @app.route("/scheduled/new", methods=["GET", "POST"])
    @app.route("/planowanie/nowe", methods=["GET", "POST"])
    def new_scheduled_event():
        if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
            flash("Brak uprawnień do planowania zdarzeń.", "error")
            return redirect(url_for("index"))
        requested_planned_at = parse_datetime(request.args.get("planned_at"))
        if request.method == "POST":
            try:
                save_scheduled_event(request.form)
                flash("Zdarzenie zostało zaplanowane.", "success")
                return redirect(url_for("scheduled_events"))
            except ValueError as exc:
                flash(str(exc), "error")
        return render_template(
            "new_scheduled_event.html",
            machine_names=list_machine_names(),
            event_types=SCHEDULED_EVENT_TYPES,
            planned_at_value=format_datetime(
                requested_planned_at.isoformat(sep=" ", timespec="minutes")
                if requested_planned_at else None
            ),
        )

    @app.route("/scheduled/<int:event_id>/done", methods=["POST"])
    @app.route("/planowanie/<int:event_id>/done", methods=["POST"])
    def scheduled_event_done(event_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
            flash("Brak uprawnień do aktualizacji zaplanowanych zdarzeń.", "error")
            return redirect(url_for("scheduled_events"))
        try:
            record_id, created, next_planned_at = complete_scheduled_event(
                event_id,
                request.form.get("expected_planned_at"),
            )
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("scheduled_events"))
        if created:
            if record_id is None:
                message = "Zdarzenie własne oznaczono jako wykonane; nie dodano go do tabeli głównej."
                if next_planned_at:
                    message += f" Następny termin: {format_datetime(next_planned_at)}."
                flash(message, "success")
            elif next_planned_at:
                flash(
                    f"Wykonano zdarzenie cykliczne i dodano rekord #{record_id} do rejestru. "
                    f"Następny termin: {format_datetime(next_planned_at)}.",
                    "success",
                )
            else:
                flash(f"Zdarzenie oznaczono jako wykonane i dodano do rejestru jako rekord #{record_id}.", "success")
        else:
            flash("Zdarzenie było już wykonane i powiązane z rejestrem.", "info")
        return redirect(url_for("scheduled_events"))

    @app.route("/scheduled/<int:event_id>/undo", methods=["POST"])
    @app.route("/planowanie/<int:event_id>/undo", methods=["POST"])
    def scheduled_event_undo(event_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
            flash("Brak uprawnień do cofania wykonania zdarzeń.", "error")
            return redirect(url_for("scheduled_events"))
        try:
            record_id, next_event_removed = undo_scheduled_event_completion(
                event_id,
                request.form.get("expected_planned_at"),
            )
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("scheduled_events"))
        message = "Cofnięto oznaczenie zdarzenia jako wykonanego."
        if record_id is not None:
            message += f" Usunięto powiązany rekord rejestru #{record_id}."
        if next_event_removed:
            message += " Usunięto również następny termin utworzony przez powtarzanie."
        flash(message, "success")
        return redirect(url_for("scheduled_events"))

    @app.route("/scheduled/<int:event_id>/delete", methods=["POST"])
    @app.route("/planowanie/<int:event_id>/delete", methods=["POST"])
    def scheduled_event_delete(event_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_delete_records"):
            flash("Brak uprawnień do usuwania zaplanowanych zdarzeń.", "error")
            return redirect(url_for("scheduled_events"))
        try:
            delete_scheduled_event(event_id)
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("scheduled_events"))
        flash("Zaplanowane zdarzenie zostało usunięte.", "success")
        return redirect(url_for("scheduled_events"))

    @app.route("/scheduled/<int:event_id>/edit", methods=["POST"])
    @app.route("/planowanie/<int:event_id>/edit", methods=["POST"])
    def scheduled_event_edit(event_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
            flash("Brak uprawnień do edycji zaplanowanych zdarzeń.", "error")
            return redirect(url_for("scheduled_events"))
        try:
            update_scheduled_event(event_id, request.form)
            flash("Zdarzenie zostało zaktualizowane.", "success")
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("scheduled_events"))

    @app.route("/api/scheduled-events", methods=["GET", "POST"])
    def api_scheduled_events():
        if request.method == "POST":
            if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
                return jsonify({"error": "Brak uprawnień do tworzenia planu."}), 403
            try:
                event_id = save_scheduled_event(request.form)
                return jsonify({"id": event_id, "status": "created"}), 201
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 400
        rows = get_scheduled_events()
        return jsonify([{
            "id": row["id"],
            "event_type": row["event_type"],
            "description": row["description"],
            "cause": row["cause"],
            "repair_method": row["repair_method"],
            "replaced_parts": row["replaced_parts"],
            "notes": row["notes"],
            "planned_at": row["planned_at"],
            "status": row["status"],
            "machine": row["name"],
            "machine_type": row["machine_type"],
            "serial_number": row["serial_number"],
            "repeat_interval": row["repeat_interval"],
            "repeat_unit": row["repeat_unit"],
        } for row in rows])

    @app.route("/api/scheduled-events/<int:event_id>", methods=["GET", "PATCH", "DELETE"])
    def api_scheduled_event_detail(event_id: int):
        event = get_db().execute(
            "SELECT se.*, m.name, m.machine_type, m.serial_number FROM scheduled_events se JOIN machines m ON m.id = se.machine_id WHERE se.id = ?",
            (event_id,),
        ).fetchone()
        if event is None or not scheduled_event_is_visible_to_current_user(event):
            return jsonify({"error": "Zdarzenie nie istnieje."}), 404
        if request.method == "GET":
            return jsonify({
                "id": event["id"],
                "event_type": event["event_type"],
                "description": event["description"],
                "cause": event["cause"],
                "repair_method": event["repair_method"],
                "replaced_parts": event["replaced_parts"],
                "notes": event["notes"],
                "planned_at": event["planned_at"],
                "status": event["status"],
                "machine": event["name"],
                "machine_type": event["machine_type"],
                "serial_number": event["serial_number"],
            })
        if request.method == "PATCH":
            if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
                return jsonify({"error": "Brak uprawnień do aktualizacji."}), 403
            try:
                update_scheduled_event(event_id, request.form or request.get_json(silent=True) or {})
                return jsonify({"status": "updated"})
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 400
        if request.method == "DELETE":
            if not app.config.get("TESTING") and not user_has_permission("can_delete_records"):
                return jsonify({"error": "Brak uprawnień do usuwania."}), 403
            try:
                delete_scheduled_event(event_id)
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 404
            return jsonify({"status": "deleted"})
        return jsonify({"error": "Nieobsługiwane żądanie."}), 405

    @app.route("/records-calendar")
    def records_calendar():
        if not app.config.get("TESTING") and not user_has_permission("can_view_records"):
            flash("Brak uprawnień do przeglądania zdarzeń.", "error")
            return redirect(url_for("index"))

        records = query_records_latest(None)
        calendar_view = request.args.get("calendar_view", "day")
        if calendar_view not in {"day", "month", "year", "week"}:
            calendar_view = "month"

        current_date = datetime.now()
        try:
            calendar_date = date.fromisoformat(request.args.get("calendar_date", ""))
        except ValueError:
            calendar_date = current_date.date()
        current_iso_year, current_iso_week, _ = current_date.isocalendar()
        try:
            calendar_year = int(request.args.get(
                "year",
                current_iso_year if calendar_view == "week" else current_date.year,
            ))
        except (TypeError, ValueError):
            calendar_year = current_iso_year if calendar_view == "week" else current_date.year
        if not 1 <= calendar_year <= 9999:
            calendar_year = current_iso_year if calendar_view == "week" else current_date.year

        try:
            calendar_month = int(request.args.get("month", current_date.month))
        except (TypeError, ValueError):
            calendar_month = current_date.month
        if not 1 <= calendar_month <= 12:
            calendar_month = current_date.month

        month_names = (
            "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień",
        )
        events_by_date: dict[str, list[sqlite3.Row]] = {}
        dated_records = [
            (parse_datetime(record["occurrence_at"]), record)
            for record in records
        ]
        dated_records.sort(key=lambda item: item[0] or datetime.max)
        for occurrence_at, record in dated_records:
            if occurrence_at is not None:
                events_by_date.setdefault(occurrence_at.date().isoformat(), []).append(record)

        calendar_months = []
        calendar_week = current_iso_week
        calendar_week_value = ""
        weekly_days = []
        calendar_slots = []
        if calendar_view == "day":
            slot_events: list[list[dict[str, Any]]] = [[] for _ in range(48)]
            for record in events_by_date.get(calendar_date.isoformat(), []):
                occurrence_at = parse_datetime(record["occurrence_at"])
                if occurrence_at is not None:
                    slot_index = occurrence_at.hour * 2 + occurrence_at.minute // 30
                    slot_events[slot_index].append({
                        "event": record,
                        "time": occurrence_at.strftime("%H:%M"),
                    })
            calendar_slots = [
                {
                    "time": f"{slot_index // 2:02d}:{'30' if slot_index % 2 else '00'}",
                    "occurrence_at": f"{calendar_date.isoformat()}T{slot_index // 2:02d}:{'30' if slot_index % 2 else '00'}",
                    "events": slot_events[slot_index],
                }
                for slot_index in range(48)
            ]
            try:
                previous_day = calendar_date - timedelta(days=1)
            except OverflowError:
                previous_day = None
            try:
                next_day = calendar_date + timedelta(days=1)
            except OverflowError:
                next_day = None
            previous_url = (
                url_for("records_calendar", calendar_view="day", calendar_date=previous_day.isoformat())
                if previous_day else None
            )
            next_url = (
                url_for("records_calendar", calendar_view="day", calendar_date=next_day.isoformat())
                if next_day else None
            )
            calendar_heading = calendar_date.strftime("%d-%m-%Y")
        elif calendar_view == "week":
            requested_iso_week = request.args.get("iso_week", "")
            try:
                iso_year_text, iso_week_text = requested_iso_week.split("-W")
                calendar_year = int(iso_year_text)
                calendar_week = int(iso_week_text)
                week_start = datetime.fromisocalendar(calendar_year, calendar_week, 1).date()
            except (TypeError, ValueError):
                calendar_year = current_iso_year
                calendar_week = current_iso_week
                week_start = current_date.date() - timedelta(days=current_date.weekday())
            calendar_week_value = f"{calendar_year:04d}-W{calendar_week:02d}"
            for day_offset in range(7):
                try:
                    day = week_start + timedelta(days=day_offset)
                except OverflowError:
                    day = None
                hourly_events: list[list[dict[str, Any]]] = [[] for _ in range(24)]
                if day is not None:
                    for record in events_by_date.get(day.isoformat(), []):
                        occurrence_at = parse_datetime(record["occurrence_at"])
                        if occurrence_at is not None:
                            hourly_events[occurrence_at.hour].append({
                                "event": record,
                                "time": occurrence_at.strftime("%H:%M"),
                            })
                weekly_days.append({
                    "date": day,
                    "is_today": day == current_date.date(),
                    "hours": hourly_events,
                })
            try:
                previous_week_start = week_start - timedelta(days=7)
            except OverflowError:
                previous_week_start = None
            try:
                next_week_start = week_start + timedelta(days=7)
            except OverflowError:
                next_week_start = None
            previous_url = None
            if previous_week_start is not None:
                previous_iso_year, previous_iso_week, _ = previous_week_start.isocalendar()
                previous_url = url_for(
                    "records_calendar",
                    calendar_view="week",
                    iso_week=f"{previous_iso_year:04d}-W{previous_iso_week:02d}",
                )
            next_url = None
            if next_week_start is not None:
                next_iso_year, next_iso_week, _ = next_week_start.isocalendar()
                next_url = url_for(
                    "records_calendar",
                    calendar_view="week",
                    iso_week=f"{next_iso_year:04d}-W{next_iso_week:02d}",
                )
            calendar_heading = f"Tydzień {calendar_week}, {calendar_year}"

        if calendar_view not in {"day", "week"}:
            month_numbers = range(1, 13) if calendar_view == "year" else (calendar_month,)
            calendar_generator = calendar.Calendar(firstweekday=calendar.MONDAY)
            for month_number in month_numbers:
                weeks = []
                for week in calendar_generator.monthdatescalendar(calendar_year, month_number):
                    weeks.append({
                        "number": week[0].isocalendar().week,
                        "days": [
                            {
                                "date": day,
                                "in_month": day.month == month_number,
                                "is_today": day == current_date.date(),
                                "events": events_by_date.get(day.isoformat(), []) if day.month == month_number else [],
                            }
                            for day in week
                        ],
                    })
                calendar_months.append({
                    "number": month_number,
                    "name": month_names[month_number - 1],
                    "weeks": weeks,
                })

            if calendar_view == "year":
                previous_url = (
                    url_for("records_calendar", calendar_view="year", year=calendar_year - 1)
                    if calendar_year > 1 else None
                )
                next_url = (
                    url_for("records_calendar", calendar_view="year", year=calendar_year + 1)
                    if calendar_year < 9999 else None
                )
                calendar_heading = str(calendar_year)
            else:
                previous_year, previous_month = (
                    (calendar_year - 1, 12) if calendar_month == 1 else (calendar_year, calendar_month - 1)
                )
                next_year, next_month = (
                    (calendar_year + 1, 1) if calendar_month == 12 else (calendar_year, calendar_month + 1)
                )
                previous_url = (
                    url_for("records_calendar", calendar_view="month", year=previous_year, month=previous_month)
                    if previous_year >= 1 else None
                )
                next_url = (
                    url_for("records_calendar", calendar_view="month", year=next_year, month=next_month)
                    if next_year <= 9999 else None
                )
                calendar_heading = f"{month_names[calendar_month - 1]} {calendar_year}"

        return render_template(
            "records_calendar.html",
            calendar_view=calendar_view,
            calendar_year=calendar_year,
            calendar_month=calendar_month,
            calendar_date=calendar_date.isoformat(),
            calendar_week=calendar_week,
            calendar_week_value=calendar_week_value,
            calendar_months=calendar_months,
            calendar_slots=calendar_slots,
            weekly_days=weekly_days,
            calendar_hours=range(24),
            calendar_heading=calendar_heading,
            calendar_weekdays=("Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"),
            previous_calendar_url=previous_url,
            next_calendar_url=next_url,
            today_calendar_url=(
                url_for(
                    "records_calendar",
                    calendar_view="day",
                    calendar_date=current_date.date().isoformat(),
                )
                if calendar_view == "day"
                else url_for(
                    "records_calendar",
                    calendar_view="week",
                    iso_week=f"{current_iso_year:04d}-W{current_iso_week:02d}",
                )
                if calendar_view == "week"
                else url_for(
                    "records_calendar",
                    calendar_view="month",
                    year=current_date.year,
                    month=current_date.month,
                )
                if calendar_view == "month"
                else url_for("records_calendar", calendar_view="year", year=current_date.year)
            ),
            month_names=month_names,
            event_types=EVENT_TYPES,
            can_create_records=app.config.get("TESTING") or user_has_permission("can_create_records"),
            can_edit_records=app.config.get("TESTING") or user_has_permission("can_edit_records"),
            can_delete_records=app.config.get("TESTING") or user_has_permission("can_delete_records"),
        )

    @app.route("/", methods=["GET", "POST"])
    def index():
        if request.method == "POST":
            if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
                flash("Brak uprawnień do dodawania rekordów.", "error")
                return redirect(url_for("index"))
            try:
                save_record(request.form)
                flash("Rekord został zapisany.", "success")
                return redirect(url_for("index"))
            except ValueError as exc:
                flash(str(exc), "error")

        raw_limit = request.args.get("limit", "20")
        if raw_limit == "all":
            limit = "all"
        else:
            try:
                limit = max(1, min(int(raw_limit or 20), 500))
            except (ValueError, TypeError):
                limit = 20
        due_events = get_due_scheduled_events()
        for event in due_events:
            flash(
                f"Maszyna {event['name']} ma zaplanowane zdarzenie {event['event_type'].lower()} na {format_datetime(event['planned_at'])}.",
                "info",
            )
        return render_template(
            "index.html",
            latest_records=query_records_latest(None if limit == "all" else limit),
            dashboard_kpis=query_dashboard_kpis(),
            event_types=EVENT_TYPES,
            limit=limit,
            due_events=due_events,
        )

    @app.route("/new", methods=["GET", "POST"])
    def new_record():
        if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
            flash("Brak uprawnie\u0144 do dodawania rekord\u00f3w.", "error")
            return redirect(url_for("index"))
        return_calendar_date = request.values.get("return_calendar_date", "")
        try:
            return_calendar_date = date.fromisoformat(return_calendar_date).isoformat()
        except ValueError:
            return_calendar_date = ""
        form_error = ""
        field_errors: dict[str, str] = {}
        form_values = request.form if request.method == "POST" else {}
        if request.method == "POST":
            try:
                record_id = save_record(request.form)
                for f in request.files.getlist("attachments"):
                    if f and f.filename:
                        save_attachment(record_id, f)
                flash("Rekord został zapisany.", "success")
                if return_calendar_date:
                    return redirect(url_for(
                        "records_calendar",
                        calendar_view="day",
                        calendar_date=return_calendar_date,
                    ))
                return redirect(url_for("index"))
            except ValueError as exc:
                form_error = str(exc)
                if "Data wystąpienia" in form_error:
                    field_errors["occurrence_at"] = form_error
                elif "Data naprawienia" in form_error:
                    field_errors["repaired_at"] = form_error
                elif "Maszyna, typ i numer" in form_error:
                    for field_name in ("machine_name", "machine_type", "serial_number"):
                        field_errors[field_name] = form_error
        return render_template(
            "new_record.html",
            machine_names=list_machine_names(),
            event_types=EVENT_TYPES,
            initial_occurrence_at=format_datetime(request.args.get("occurrence_at")),
            return_calendar_date=return_calendar_date,
            form_error=form_error,
            field_errors=field_errors,
            form_values=form_values,
            labor_rate=get_labor_rate(),
        )

    @app.route("/import", methods=["GET", "POST"])
    def import_data():
        if request.method == "GET":
            return redirect(url_for("data_page"), code=301)
        if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
            flash("Brak uprawnie\u0144 do importu danych.", "error")
            return redirect(url_for("data_page"))
        upload = request.files.get("file")
        if not upload or not upload.filename:
            flash("Wybierz plik Excel do importu.", "error")
            return redirect(url_for("data_page"))
        imported, skipped, attachments_added, attachments_skipped = import_workbook(upload.read(), upload.filename or "")
        log_user_activity(
            "import_records",
            f"Zaimportowano {imported}, pominięto {skipped} rekordów z pliku {upload.filename or 'bez nazwy'}.",
            target_type="import",
        )
        message = f"Zaimportowano {imported} rekordów, pominięto {skipped}."
        if attachments_added or attachments_skipped:
            message += f" Załączniki: dodano {attachments_added}, pominięto {attachments_skipped}."
        flash(message, "success")
        return redirect(url_for("index"))

    @app.route("/data", methods=["GET", "POST"])
    def data_page():
        if request.method == "POST":
            if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
                flash("Brak uprawnie\u0144 do importu danych.", "error")
                return redirect(url_for("data_page"))
            upload = request.files.get("file")
            if not upload or not upload.filename:
                flash("Wybierz plik Excel do importu.", "error")
                return redirect(url_for("data_page"))
            imported, skipped, attachments_added, attachments_skipped = import_workbook(upload.read(), upload.filename or "")
            log_user_activity(
                "import_records",
                f"Zaimportowano {imported}, pominięto {skipped} rekordów z pliku {upload.filename or 'bez nazwy'}.",
                target_type="import",
            )
            message = f"Zaimportowano {imported} rekordów, pominięto {skipped}."
            if attachments_added or attachments_skipped:
                message += f" Załączniki: dodano {attachments_added}, pominięto {attachments_skipped}."
            flash(message, "success")
            return redirect(url_for("index"))
        return render_template("import_export.html")

    @app.route("/export")
    def export_data():
        if not app.config.get("TESTING") and session.get("role") == "guest":
            flash("Konto gościa nie ma uprawnień do eksportu danych.", "error")
            return redirect(url_for("filter_page"))

        filters = {
            "machine": request.args.get("machine", "").strip(),
            "machine_type": request.args.get("machine_type", "").strip(),
            "serial_number": request.args.get("serial_number", "").strip(),
            "event_type": request.args.get("event_type", "").strip(),
            "occurrence_from": format_date(request.args.get("occurrence_from", "").strip()),
            "occurrence_to": format_date(request.args.get("occurrence_to", "").strip()),
            "description_terms": request.args.get("description_terms", "").strip(),
        }
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "ZBIORCZE"
        sheet.append(EXPORT_COLUMNS)
        for record in query_records(filters):
            sheet.append(
                [
                    record["id"],
                    record["record_year"],
                    record["record_week"],
                    record["week_sequence"],
                    record["name"],
                    record["machine_type"],
                    record["serial_number"],
                    record["event_type"],
                    record["description"],
                    record["cause"],
                    record["repair_method"],
                    record["replaced_parts"],
                    record["participants"],
                    format_datetime(record["occurrence_at"]),
                    format_datetime(record["repaired_at"]),
                    record["machine_mtg"],
                    record["downtime_hours"],
                    record["labor_cost"],
                    record["material_cost"],
                    record["spare_parts_cost"],
                    record["service_cost"],
                    record["total_cost"],
                    record["notes"],
                ]
            )
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        log_user_activity(
            "export_records",
            "Wyeksportowano rekordy do pliku Excel.",
            target_type="export",
        )
        return send_file(
            output,
            as_attachment=True,
            download_name="tygodniowy_raport_napraw.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.route("/export/bundle")
    def export_records_bundle():
        if not app.config.get("TESTING") and session.get("role") == "guest":
            flash("Konto gościa nie ma uprawnień do eksportu danych.", "error")
            return redirect(url_for("filter_page"))

        records = query_records({
            "machine": "", "machine_type": "", "serial_number": "", "event_type": "",
            "occurrence_from": "", "occurrence_to": "", "description_terms": "",
        })
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "ZBIORCZE"
        sheet.append(EXPORT_COLUMNS)
        for record in records:
            sheet.append([
                record["id"], record["record_year"], record["record_week"], record["week_sequence"],
                record["name"], record["machine_type"], record["serial_number"], record["event_type"],
                record["description"], record["cause"], record["repair_method"], record["replaced_parts"],
                record["participants"], format_datetime(record["occurrence_at"]),
                format_datetime(record["repaired_at"]), record["machine_mtg"], record["downtime_hours"],
                record["labor_cost"], record["material_cost"], record["spare_parts_cost"],
                record["service_cost"], record["total_cost"], record["notes"],
            ])
        workbook_output = io.BytesIO()
        workbook.save(workbook_output)

        manifest = {"format": "repair-records-bundle", "version": 1, "attachments": []}
        upload_root = Path(current_app_uploads())
        record_ids = [int(record["id"]) for record in records]
        if record_ids:
            placeholders = ",".join("?" for _ in record_ids)
            attachments = get_db().execute(
                f"SELECT record_id, filename, original_name, content_type FROM attachments "
                f"WHERE record_id IN ({placeholders}) ORDER BY id",
                record_ids,
            ).fetchall()
        else:
            attachments = []

        bundle_output = io.BytesIO()
        with zipfile.ZipFile(bundle_output, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("records.xlsx", workbook_output.getvalue())
            for attachment in attachments:
                member_name = f"attachments/{int(attachment['record_id'])}/{Path(attachment['filename']).name}"
                file_path = upload_root / str(attachment["record_id"]) / Path(attachment["filename"]).name
                if not file_path.is_file():
                    continue
                archive.write(file_path, member_name)
                manifest["attachments"].append({
                    "record_id": int(attachment["record_id"]),
                    "filename": Path(attachment["filename"]).name,
                    "original_name": attachment["original_name"],
                    "content_type": attachment["content_type"],
                    "path": member_name,
                })
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))

        bundle_output.seek(0)
        log_user_activity("export_records", "Wyeksportowano zdarzenia wraz z załącznikami.", target_type="export")
        return send_file(
            bundle_output,
            as_attachment=True,
            download_name="zdarzenia_z_zalacznikami.zip",
            mimetype="application/zip",
        )

    @app.route("/admin/export/full-instance-backup")
    def export_full_instance_backup():
        if not app.config.get("TESTING") and session.get("role") != "admin":
            flash("Pełną kopię bazy może pobrać tylko administrator.", "error")
            return redirect(url_for("data_page"))

        snapshot_directory = tempfile.TemporaryDirectory(prefix="instance-backup-")
        snapshot_path = Path(snapshot_directory.name) / "repair_reports.sqlite3"
        source_db = sqlite3.connect(current_app_database(), timeout=60)
        snapshot_db = sqlite3.connect(str(snapshot_path))
        try:
            source_db.backup(snapshot_db, pages=256, sleep=0.1)
        finally:
            snapshot_db.close()
            source_db.close()

        chunks: queue.Queue = queue.Queue(maxsize=8)
        cancelled = threading.Event()
        errors: list[BaseException] = []
        zip_writer = _QueueZipWriter(chunks, cancelled)
        uploads_root = Path(current_app_uploads())

        def produce_archive() -> None:
            try:
                with zipfile.ZipFile(zip_writer, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as archive:
                    archive.write(snapshot_path, "instance/repair_reports.sqlite3")
                    if uploads_root.is_dir():
                        for file_path in uploads_root.rglob("*"):
                            if cancelled.is_set():
                                break
                            if file_path.is_file():
                                relative_path = file_path.relative_to(uploads_root).as_posix()
                                archive.write(file_path, f"instance/uploads/{relative_path}")
            except BaseException as exc:
                errors.append(exc)
            finally:
                if not cancelled.is_set():
                    try:
                        chunks.put(None, timeout=1)
                    except queue.Full:
                        pass

        producer = threading.Thread(target=produce_archive, name="instance-backup-zip", daemon=True)
        producer.start()

        def stream_archive():
            try:
                while True:
                    try:
                        chunk = chunks.get(timeout=0.1)
                    except queue.Empty:
                        if not producer.is_alive():
                            break
                        continue
                    if chunk is None:
                        break
                    yield chunk
                if errors:
                    raise errors[0]
            finally:
                cancelled.set()
                producer.join(timeout=2)
                snapshot_directory.cleanup()

        log_user_activity(
            "export_full_backup",
            "Rozpoczęto pobieranie pełnej kopii bazy i załączników.",
            target_type="backup",
        )
        response = Response(stream_archive(), mimetype="application/zip")
        response.headers["Content-Disposition"] = (
            f'attachment; filename="instance_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip"'
        )
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.route("/filter")
    def filter_page():
        filters = {
            "machine": request.args.get("machine", "").strip(),
            "machine_type": request.args.get("machine_type", "").strip(),
            "serial_number": request.args.get("serial_number", "").strip(),
            "event_type": request.args.get("event_type", "").strip(),
            "occurrence_from": format_date(request.args.get("occurrence_from", "").strip()),
            "occurrence_to": format_date(request.args.get("occurrence_to", "").strip()),
            "description_terms": request.args.get("description_terms", "").strip(),
        }
        filter_active = any(filters.values())
        return render_template(
            "filter_export.html",
            filtered_records=query_records(filters) if filter_active else [],
            filter_active=filter_active,
            machine_names=list_machine_names(),
            event_types=EVENT_TYPES,
            filters=filters,
        )

    @app.route("/api/machines")
    def machine_lookup():
        machine_name = request.args.get("machine", "").strip()
        machine_type = request.args.get("machine_type", "").strip()
        serial_number = request.args.get("serial_number", "").strip()
        return jsonify(
            {
                "types": list_machine_types(machine_name),
                "serial_numbers": list_serial_numbers(machine_name, machine_type),
                "previous_entries": list_previous_entries(machine_name, machine_type, serial_number),
            }
        )

    @app.route("/api/record/<int:record_id>")
    def record_details_api(record_id: int):
        db = get_db()
        row = db.execute(
            """
            SELECT r.*, m.name, m.machine_type, m.serial_number
            FROM repair_records r
            JOIN machines m ON m.id = r.machine_id
            WHERE r.id = ?
            """,
            (record_id,),
        ).fetchone()
        if not row:
            return jsonify({"error": "Rekord nie istnieje."}), 404
        record = dict(row)
        attachments = db.execute(
            "SELECT id, filename, original_name, content_type FROM attachments WHERE record_id = ? ORDER BY created_at",
            (record_id,),
        ).fetchall()
        record["attachments"] = [dict(a) for a in attachments]
        log_user_activity(
            "view_record",
            f"Wyświetlono rekord #{record_id}.",
            target_type="record",
            target_id=record_id,
        )
        return jsonify(record)

    @app.route("/record/<int:record_id>/pdf")
    def record_pdf(record_id: int):
        """Export one record as a monochrome PDF, including image attachments."""
        db = get_db()
        row = db.execute(
            """
            SELECT r.*, m.name, m.machine_type, m.serial_number
            FROM repair_records r JOIN machines m ON m.id = r.machine_id
            WHERE r.id = ?
            """,
            (record_id,),
        ).fetchone()
        if not row:
            return "Rekord nie istnieje.", 404

        attachments = db.execute(
            "SELECT filename, original_name, content_type FROM attachments WHERE record_id = ? ORDER BY created_at",
            (record_id,),
        ).fetchall()
        record = dict(row)
        output = io.BytesIO()

        # Prefer an installed Windows font so Polish characters are preserved.
        font_path = next((Path(p) for p in (
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            r"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ) if Path(p).exists()), None)
        font_name = "Helvetica"
        if font_path:
            font_name = "ReportFont"
            pdfmetrics.registerFont(TTFont(font_name, str(font_path)))

        styles = getSampleStyleSheet()
        normal = ParagraphStyle("record-normal", parent=styles["Normal"], fontName=font_name,
                                fontSize=8.5, leading=11, textColor=colors.black)
        heading = ParagraphStyle("record-heading", parent=normal, fontSize=15, leading=18,
                                 spaceAfter=5, alignment=TA_LEFT)
        label = ParagraphStyle("record-label", parent=normal, fontName=font_name,
                       fontSize=8, leading=10, textColor=colors.black)

        def text(value: Any) -> str:
            return str(value if value not in (None, "") else "—").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

        fields = [
            ("Maszyna", record["name"]), ("Typ", record["machine_type"]), ("Numer", record["serial_number"]),
            ("Rok / tydzień / Lp.", f'{record["record_year"]} / W{record["record_week"]} / {record["week_sequence"]}'),
            ("Rodzaj zdarzenia", record["event_type"]), ("MTG maszyny", record["machine_mtg"]),
            ("Data wystąpienia", format_datetime(record["occurrence_at"])), ("Data naprawienia", format_datetime(record["repaired_at"])),
            ("Opis", record["description"]), ("Przyczyna", record["cause"]), ("Sposób naprawy", record["repair_method"]),
            ("Wymienione części", record["replaced_parts"]), ("Osoby biorące udział", record["participants"]),
            ("Uwagi", record["notes"]), ("Pracochłonność [h]", record["labor_cost"]), ("Materiał PLN", record["material_cost"]),
            ("Części zamienne PLN", record["spare_parts_cost"]), ("Serwis PLN", record["service_cost"]),
            ("Koszt netto PLN", record["total_cost"]), ("Wykluczenie z produkcji [h]", record["downtime_hours"]),
        ]
        data = [[Paragraph(f"<b>{text(name)}</b>", label), Paragraph(text(value), normal)] for name, value in fields]
        table = Table(data, colWidths=[38 * mm, 140 * mm], repeatRows=0)
        table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.7, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.black),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eeeeee")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))

        if record["name"] and record["machine_type"]== "-":
            head_text=f"Zdarzenie #{record_id} : {text(record['name'])}"
        else:
            head_text=f"Zdarzenie #{record_id} : {text(record['name'])} — {text(record['machine_type'])}"

        story = [Paragraph(head_text, heading), table]

        image_files = []
        for attachment in attachments:
            if not attachment["content_type"].startswith("image/"):
                continue
            path = Path(current_app_uploads()) / str(record_id) / attachment["filename"]
            if path.exists():
                image_files.append((path, attachment["original_name"]))
        if image_files:
            story += [PageBreak(), Paragraph("Załączniki — zdjęcia", heading)]
            for path, original_name in image_files:
                image = PdfImage(str(path), width=170 * mm, height=105 * mm, kind="proportional")
                story += [Paragraph(text(original_name), normal), Spacer(1, 2 * mm), image, Spacer(1, 6 * mm)]
        SimpleDocTemplate(output, pagesize=A4, rightMargin=16 * mm, leftMargin=16 * mm,
                          topMargin=15 * mm, bottomMargin=15 * mm,
                          title=f"Rekord {record_id}").build(story, canvasmaker=PageNumberCanvas)
        output.seek(0)
        log_user_activity(
            "export_record_pdf",
            f"Wyeksportowano rekord #{record_id} do PDF.",
            target_type="record",
            target_id=record_id,
        )
        return send_file(output, as_attachment=True, download_name=f"rekord_{record_id}_{record['name']}.pdf", mimetype="application/pdf")

    @app.route("/api/record/<int:record_id>", methods=["PATCH"])
    def update_record_api(record_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
            return jsonify({"error": "Brak uprawnie\u0144."}), 403
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "Brak danych."}), 400
        try:
            update_record(record_id, data)
            log_user_activity(
                "update_record",
                f"Zaktualizowano rekord #{record_id}.",
                target_type="record",
                target_id=record_id,
            )
            return jsonify({"ok": True})
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400

    @app.route("/api/record/<int:record_id>", methods=["DELETE"])
    def delete_record_api(record_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_delete_records"):
            return jsonify({"error": "Brak uprawnie\u0144."}), 403
        import shutil

        db = get_db()
        if not db.execute("SELECT id FROM repair_records WHERE id = ?", (record_id,)).fetchone():
            return jsonify({"error": "Rekord nie istnieje."}), 404
        upload_dir = Path(current_app_uploads()) / str(record_id)
        if upload_dir.exists():
            shutil.rmtree(str(upload_dir))
        db.execute("DELETE FROM attachments WHERE record_id = ?", (record_id,))
        db.execute("DELETE FROM repair_records WHERE id = ?", (record_id,))
        db.commit()
        log_user_activity(
            "delete_record",
            f"Usunięto rekord #{record_id}.",
            target_type="record",
            target_id=record_id,
        )
        return jsonify({"ok": True})

    @app.route("/api/record/<int:record_id>/attachments", methods=["POST"])
    def add_record_attachments(record_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_edit_records"):
            return jsonify({"error": "Brak uprawnie\u0144."}), 403
        db = get_db()
        if not db.execute("SELECT id FROM repair_records WHERE id = ?", (record_id,)).fetchone():
            return jsonify({"error": "Rekord nie istnieje."}), 404
        for f in request.files.getlist("attachments"):
            if f and f.filename:
                save_attachment(record_id, f)
        return jsonify({"ok": True})

    @app.route("/api/attachment/<int:attachment_id>", methods=["DELETE"])
    def delete_attachment_api(attachment_id: int):
        if not app.config.get("TESTING") and not user_has_permission("can_delete_records"):
            return jsonify({"error": "Brak uprawnie\u0144."}), 403
        db = get_db()
        row = db.execute("SELECT * FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
        if not row:
            return jsonify({"error": "Załącznik nie istnieje."}), 404
        file_path = Path(current_app_uploads()) / str(row["record_id"]) / row["filename"]
        if file_path.exists():
            file_path.unlink()
        db.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
        db.commit()
        log_user_activity(
            "delete_attachment",
            f"Usunięto załącznik {row['original_name']} z rekordu #{row['record_id']}.",
            target_type="record",
            target_id=int(row["record_id"]),
        )
        return jsonify({"ok": True})

    @app.route("/attachments/<int:record_id>/<path:filename>")
    def serve_attachment(record_id: int, filename: str):
        row = get_db().execute(
            "SELECT filename FROM attachments WHERE record_id = ? AND filename = ?",
            (record_id, filename),
        ).fetchone()
        if not row:
            return "Not found", 404
        upload_dir = Path(current_app_uploads()) / str(record_id)
        return send_from_directory(str(upload_dir), filename)

    @app.route("/record/<int:record_id>/attachments.zip")
    def download_record_attachments(record_id: int):
        attachments = get_db().execute(
            "SELECT filename, original_name FROM attachments WHERE record_id = ? ORDER BY created_at",
            (record_id,),
        ).fetchall()
        if not attachments:
            return "Załączniki nie istnieją.", 404

        output = io.BytesIO()
        upload_dir = Path(current_app_uploads()) / str(record_id)
        used_names: set[str] = set()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for attachment in attachments:
                file_path = upload_dir / attachment["filename"]
                if not file_path.is_file():
                    return "Plik załącznika nie istnieje.", 404

                archive_name = Path(attachment["original_name"].replace("\\", "/")).name
                if not archive_name:
                    archive_name = attachment["filename"]
                stem, suffix = Path(archive_name).stem, Path(archive_name).suffix
                unique_name = archive_name
                counter = 1
                while unique_name.casefold() in used_names:
                    unique_name = f"{stem}_{counter}{suffix}"
                    counter += 1
                used_names.add(unique_name.casefold())
                archive.write(file_path, arcname=unique_name)

        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=f"zalaczniki_rekordu_{record_id}.zip",
            mimetype="application/zip",
        )

    return app


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(current_app_database())
        g.db.row_factory = sqlite3.Row
    return g.db


def current_app_database() -> str:
    from flask import current_app

    return current_app.config["DATABASE"]


def current_app_uploads() -> str:
    from flask import current_app

    return current_app.config["UPLOAD_DIR"]


def save_attachment(record_id: int, file_storage: Any) -> None:
    from werkzeug.utils import secure_filename
    import mimetypes

    original_name = file_storage.filename or "file"
    safe_name = secure_filename(original_name)
    if not safe_name:
        return

    upload_dir = Path(current_app_uploads()) / str(record_id)
    upload_dir.mkdir(parents=True, exist_ok=True)

    dest = upload_dir / safe_name
    counter = 1
    while dest.exists():
        stem, suffix = Path(safe_name).stem, Path(safe_name).suffix
        dest = upload_dir / f"{stem}_{counter}{suffix}"
        counter += 1

    file_storage.save(str(dest))

    content_type = (
        file_storage.content_type
        or mimetypes.guess_type(safe_name)[0]
        or "application/octet-stream"
    )
    db = get_db()
    db.execute(
        """
        INSERT INTO attachments(record_id, filename, original_name, content_type, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (record_id, dest.name, original_name, content_type, datetime.now(UTC).isoformat(timespec="seconds")),
    )
    db.commit()
    log_user_activity(
        "add_attachment",
        f"Dodano załącznik {original_name} do rekordu #{record_id}.",
        target_type="record",
        target_id=record_id,
    )


def init_db() -> None:
    db = get_db()
    db.executescript(SCHEMA)
    machine_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(machines)").fetchall()
    }
    for column, column_type in (
        ("machine_category", "TEXT"),
        ("commissioning_date", "TEXT"),
        ("warranty_period_months", "INTEGER"),
        ("warranty_end_date", "TEXT"),
        ("power", "REAL"),
        ("power_unit", "TEXT"),
        ("protection_a", "TEXT"),
        ("seller_data", "TEXT"),
        ("notes", "TEXT"),
    ):
        if column not in machine_columns:
            db.execute(f"ALTER TABLE machines ADD COLUMN {column} {column_type}")
    scheduled_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(scheduled_events)").fetchall()
    }
    for column, column_type in (
        ("created_by_user_id", "INTEGER"),
        ("repair_record_id", "INTEGER"),
        ("cause", "TEXT"),
        ("repair_method", "TEXT"),
        ("replaced_parts", "TEXT"),
        ("recurrence_next_event_id", "INTEGER"),
    ):
        if column not in scheduled_columns:
            db.execute(f"ALTER TABLE scheduled_events ADD COLUMN {column} {column_type}")
    activity_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(user_activity_log)").fetchall()
    }
    if "ip_address" not in activity_columns:
        db.execute("ALTER TABLE user_activity_log ADD COLUMN ip_address TEXT")
    try:
        db.execute("ALTER TABLE repair_records ADD COLUMN notes TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists
    try:
        db.execute("ALTER TABLE repair_records ADD COLUMN participants TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists
    duplicate_serials = db.execute(
        """
        SELECT serial_number
        FROM machines
        WHERE lower(trim(serial_number)) NOT IN ('', 'brak', '-', 'n/a', 'nie dotyczy')
        GROUP BY serial_number
        HAVING COUNT(*) > 1
        """
    ).fetchall()
    for duplicate in duplicate_serials:
        machines = db.execute(
            """
            SELECT m.id, COUNT(r.id) AS record_count
            FROM machines m
            LEFT JOIN repair_records r ON r.machine_id = m.id
            WHERE m.serial_number = ?
            GROUP BY m.id
            ORDER BY record_count DESC, m.id ASC
            """,
            (duplicate["serial_number"],),
        ).fetchall()
        canonical_id = int(machines[0]["id"])
        for machine in machines[1:]:
            db.execute(
                "UPDATE repair_records SET machine_id = ? WHERE machine_id = ?",
                (canonical_id, int(machine["id"])),
            )
            db.execute(
                "UPDATE scheduled_events SET machine_id = ? WHERE machine_id = ?",
                (canonical_id, int(machine["id"])),
            )
            db.execute("DELETE FROM machines WHERE id = ?", (int(machine["id"]),))
    db.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_machines_serial_number
        ON machines(serial_number)
        WHERE lower(trim(serial_number)) NOT IN ('', 'brak', '-', 'n/a', 'nie dotyczy')
        """
    )
    db.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_scheduled_events_repair_record_id
        ON scheduled_events(repair_record_id)
        WHERE repair_record_id IS NOT NULL
        """
    )
    materialize_completed_scheduled_events()
    db.commit()


def is_placeholder_serial(serial_number: str) -> bool:
    return serial_number.strip().lower() in {"", "brak", "-", "n/a", "nie dotyczy"}


def get_or_create_machine(name: str, machine_type: str, serial_number: str) -> int:
    if not all([name, machine_type, serial_number]):
        raise ValueError("Maszyna, typ i numer są wymagane.")
    db = get_db()
    if is_placeholder_serial(serial_number):
        existing = db.execute(
            """
            SELECT id FROM machines
            WHERE name = ? AND machine_type = ? AND serial_number = ?
            """,
            (name, machine_type, serial_number),
        ).fetchone()
    else:
        existing = db.execute(
            "SELECT id FROM machines WHERE serial_number = ?",
            (serial_number,),
        ).fetchone()
    if existing:
        return int(existing["id"])
    cursor = db.execute(
        """
        INSERT INTO machines(name, machine_type, serial_number)
        VALUES (?, ?, ?)
        """,
        (name, machine_type, serial_number),
    )
    db.commit()
    return int(cursor.lastrowid)


def next_week_sequence(record_year: int, record_week: int) -> int:
    db = get_db()
    row = db.execute(
        """
        SELECT COALESCE(MAX(week_sequence), 0) AS max_sequence
        FROM repair_records
        WHERE record_year = ? AND record_week = ?
        """,
        (record_year, record_week),
    ).fetchone()
    return int(row["max_sequence"]) + 1


def create_record_from_scheduled_event(
    event: sqlite3.Row,
    repaired_at: datetime | None,
) -> int:
    if event["event_type"] == "Własne":
        raise ValueError("Zdarzenia własne nie mogą trafiać do tabeli głównej.")
    occurrence_at = parse_datetime(event["planned_at"])
    if occurrence_at is None:
        raise ValueError(f"Zdarzenie planowane #{event['id']} ma nieprawidłową datę.")
    record_year, record_week, _ = occurrence_at.isocalendar()
    cursor = get_db().execute(
        """
        INSERT INTO repair_records(
            record_year, record_week, week_sequence, machine_id, event_type,
            description, cause, repair_method, replaced_parts, participants, occurrence_at,
            repaired_at, machine_mtg, downtime_hours, labor_cost, material_cost,
            spare_parts_cost, service_cost, total_cost, notes, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, '', ?, 0, 0, 0, 0, 0, ?, ?)
        """,
        (
            record_year,
            record_week,
            next_week_sequence(record_year, record_week),
            int(event["machine_id"]),
            str(event["event_type"]),
            str(event["description"]),
            event["cause"],
            event["repair_method"],
            event["replaced_parts"],
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            calculate_downtime_hours(occurrence_at, repaired_at),
            event["notes"],
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    return int(cursor.lastrowid)


def materialize_completed_scheduled_events() -> None:
    db = get_db()
    events = db.execute(
        """
        SELECT * FROM scheduled_events
        WHERE status = 'done' AND repair_record_id IS NULL AND event_type != 'Własne'
        ORDER BY id
        """
    ).fetchall()
    for event in events:
        record_id = create_record_from_scheduled_event(event, None)
        db.execute(
            "UPDATE scheduled_events SET repair_record_id = ? WHERE id = ?",
            (record_id, int(event["id"])),
        )


def complete_scheduled_event(
    event_id: int,
    expected_planned_at: str | None = None,
) -> tuple[int | None, bool, str | None]:
    db = get_db()
    db.execute("BEGIN IMMEDIATE")
    event = db.execute(
        "SELECT * FROM scheduled_events WHERE id = ?",
        (event_id,),
    ).fetchone()
    if event is None:
        db.rollback()
        raise ValueError(f"Zaplanowane zdarzenie #{event_id} nie istnieje.")
    if not scheduled_event_is_visible_to_current_user(event):
        db.rollback()
        raise ValueError(f"Zaplanowane zdarzenie #{event_id} nie istnieje.")
    if expected_planned_at:
        expected = parse_datetime(expected_planned_at)
        current = parse_datetime(event["planned_at"])
        if expected is None or current is None or expected != current:
            db.rollback()
            raise ValueError("Termin tego zdarzenia został już zaktualizowany. Odśwież harmonogram.")
    if event["repair_record_id"] is not None:
        record_id = int(event["repair_record_id"])
        db.commit()
        return record_id, False, None

    if event["status"] == "done" and event["event_type"] == "Własne":
        db.commit()
        return None, False, None

    completed_at = datetime.now().replace(second=0, microsecond=0)
    record_id = (
        None
        if event["event_type"] == "Własne"
        else create_record_from_scheduled_event(event, completed_at)
    )
    next_planned_at = next_scheduled_occurrence(
        parse_datetime(event["planned_at"]),
        int(event["repeat_interval"]),
        str(event["repeat_unit"]),
    )
    updated_at = datetime.now(UTC).isoformat(timespec="seconds")
    if next_planned_at:
        completed_occurrence = db.execute(
            """
            INSERT INTO scheduled_events(
                machine_id, repair_record_id, event_type, description, cause,
                repair_method, replaced_parts, notes, planned_at, repeat_interval,
                repeat_unit, status, recurrence_next_event_id, created_at, updated_at,
                created_by_user_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'done', ?, ?, ?, ?)
            """,
            (
                int(event["machine_id"]),
                record_id,
                event["event_type"],
                event["description"],
                event["cause"],
                event["repair_method"],
                event["replaced_parts"],
                event["notes"],
                event["planned_at"],
                int(event["repeat_interval"]),
                event["repeat_unit"],
                event_id,
                updated_at,
                updated_at,
                event["created_by_user_id"],
            ),
        )
        db.execute(
            """
            UPDATE scheduled_events
            SET recurrence_next_event_id = ?
            WHERE recurrence_next_event_id = ? AND id != ?
            """,
            (int(completed_occurrence.lastrowid), event_id, int(completed_occurrence.lastrowid)),
        )
        db.execute(
            """
            UPDATE scheduled_events
            SET planned_at = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                next_planned_at.isoformat(sep=" ", timespec="minutes"),
                updated_at,
                event_id,
            ),
        )
    else:
        db.execute(
            """
            UPDATE scheduled_events
            SET status = 'done', repair_record_id = ?, updated_at = ?
            WHERE id = ?
            """,
            (record_id, updated_at, event_id),
        )
    db.commit()
    if record_id is not None:
        log_user_activity(
            "create_record",
            f"Utworzono rekord #{record_id} z zaplanowanego zdarzenia #{event_id}.",
            target_type="record",
            target_id=record_id,
        )
    completion_details = (
        f"Zrealizowano zdarzenie #{event_id} jako rekord rejestru #{record_id}."
        if record_id is not None
        else f"Zrealizowano zdarzenie własne #{event_id} bez dodawania go do rejestru."
    )
    log_user_activity(
        "complete_scheduled_event",
        completion_details,
        target_type="scheduled_event",
        target_id=event_id,
    )
    return (
        record_id,
        True,
        next_planned_at.isoformat(sep=" ", timespec="minutes") if next_planned_at else None,
    )


def undo_scheduled_event_completion(
    event_id: int,
    expected_planned_at: str | None = None,
) -> tuple[int | None, bool]:
    import shutil

    db = get_db()
    db.execute("BEGIN IMMEDIATE")
    event = db.execute(
        "SELECT * FROM scheduled_events WHERE id = ?",
        (event_id,),
    ).fetchone()
    if event is None or not scheduled_event_is_visible_to_current_user(event):
        db.rollback()
        raise ValueError(f"Wykonane zdarzenie #{event_id} nie istnieje.")
    if event["status"] != "done":
        db.rollback()
        raise ValueError("To zdarzenie nie jest oznaczone jako wykonane.")
    if expected_planned_at:
        expected = parse_datetime(expected_planned_at)
        current = parse_datetime(event["planned_at"])
        if expected is None or current is None or expected != current:
            db.rollback()
            raise ValueError("Termin tego zdarzenia został już zaktualizowany. Odśwież harmonogram.")

    next_event_id = event["recurrence_next_event_id"]
    next_planned_at = next_scheduled_occurrence(
        parse_datetime(event["planned_at"]),
        int(event["repeat_interval"]),
        str(event["repeat_unit"]),
    )
    if next_event_id is None and next_planned_at is not None:
        # Older recurring completions predate the explicit link; recover only an unambiguous match.
        candidates = db.execute(
            """
            SELECT id FROM scheduled_events
            WHERE status = 'planned' AND machine_id = ? AND planned_at = ?
              AND event_type = ? AND description = ? AND cause IS ?
              AND repair_method IS ? AND replaced_parts IS ? AND notes IS ?
              AND repeat_interval = ? AND repeat_unit = ? AND created_by_user_id IS ?
            """,
            (
                int(event["machine_id"]),
                next_planned_at.isoformat(sep=" ", timespec="minutes"),
                event["event_type"],
                event["description"],
                event["cause"],
                event["repair_method"],
                event["replaced_parts"],
                event["notes"],
                int(event["repeat_interval"]),
                event["repeat_unit"],
                event["created_by_user_id"],
            ),
        ).fetchall()
        if len(candidates) > 1:
            db.rollback()
            raise ValueError("Nie można jednoznacznie ustalić następnego terminu. Nie cofnięto wykonania.")
        if candidates:
            next_event_id = int(candidates[0]["id"])

    if next_event_id is not None:
        next_event = db.execute(
            "SELECT * FROM scheduled_events WHERE id = ?",
            (int(next_event_id),),
        ).fetchone()
        if next_event is None:
            next_event_id = None
        else:
            expected_next_at = (
                next_planned_at.isoformat(sep=" ", timespec="minutes")
                if next_planned_at is not None else None
            )
            if (
                next_event["status"] != "planned"
                or next_event["planned_at"] != expected_next_at
                or any(
                    next_event[column] != event[column]
                    for column in (
                        "machine_id", "event_type", "description", "cause",
                        "repair_method", "replaced_parts", "notes",
                        "repeat_interval", "repeat_unit", "created_by_user_id",
                    )
                )
            ):
                db.rollback()
                raise ValueError(
                    "Następny termin tego zdarzenia został zmieniony. "
                    "Nie cofnięto wykonania ani nie usunięto żadnych danych."
                )
            db.execute("DELETE FROM scheduled_events WHERE id = ?", (int(next_event_id),))

    record_id = int(event["repair_record_id"]) if event["repair_record_id"] is not None else None
    if record_id is not None:
        db.execute("DELETE FROM attachments WHERE record_id = ?", (record_id,))
        db.execute("DELETE FROM repair_records WHERE id = ?", (record_id,))
    updated_at = datetime.now(UTC).isoformat(timespec="seconds")
    db.execute(
        """
        UPDATE scheduled_events
        SET status = 'planned', repair_record_id = NULL,
            recurrence_next_event_id = NULL, updated_at = ?
        WHERE id = ?
        """,
        (updated_at, event_id),
    )
    db.commit()

    if record_id is not None:
        record_upload_dir = Path(current_app_uploads()) / str(record_id)
        if record_upload_dir.exists():
            shutil.rmtree(record_upload_dir)
        log_user_activity(
            "delete_record",
            f"Usunięto rekord #{record_id} po cofnięciu wykonania zdarzenia #{event_id}.",
            target_type="record",
            target_id=record_id,
        )
    log_user_activity(
        "undo_complete_scheduled_event",
        f"Cofnięto wykonanie zdarzenia #{event_id}."
        + (f" Usunięto następny termin #{next_event_id}." if next_event_id is not None else "")
        + (f" Usunięto rekord rejestru #{record_id}." if record_id is not None else ""),
        target_type="scheduled_event",
        target_id=event_id,
    )
    return record_id, next_event_id is not None


def save_record(form: dict[str, Any]) -> int:
    occurrence_at = parse_datetime(form.get("occurrence_at"))
    if not occurrence_at:
        raise ValueError("Data wystąpienia jest wymagana.")
    repaired_at = parse_datetime(form.get("repaired_at"))
    record_year, record_week, _ = occurrence_at.isocalendar()
    machine_id = get_or_create_machine(*machine_key_from_form(form))
    downtime_hours = calculate_downtime_hours(occurrence_at, repaired_at)
    labor_cost = as_float(form.get("labor_cost"))
    material_cost = as_float(form.get("material_cost"))
    spare_parts_cost = as_float(form.get("spare_parts_cost"))
    service_cost = as_float(form.get("service_cost"))
    total_cost = round((labor_cost * get_labor_rate()) + material_cost + spare_parts_cost + service_cost, 2)
    db = get_db()
    cursor = db.execute(
        """
        INSERT INTO repair_records(
            record_year, record_week, week_sequence, machine_id, event_type,
            description, cause, repair_method, replaced_parts, participants, occurrence_at,
            repaired_at, machine_mtg, downtime_hours, labor_cost, material_cost,
            spare_parts_cost, service_cost, total_cost, notes, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record_year,
            record_week,
            next_week_sequence(record_year, record_week),
            machine_id,
            str(form.get("event_type", "")).strip() or "Inne",
            str(form.get("description", "")).strip(),
            str(form.get("cause", "")).strip(),
            str(form.get("repair_method", "")).strip(),
            str(form.get("replaced_parts", "")).strip(),
            str(form.get("participants", "")).strip() or None,
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(form.get("machine_mtg", "")).strip(),
            downtime_hours,
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            str(form.get("notes", "")).strip() or None,
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    log_user_activity(
        "create_record",
        f"Utworzono rekord #{cursor.lastrowid} dla maszyny {str(form.get('machine_name', '')).strip()}.",
        target_type="record",
        target_id=int(cursor.lastrowid),
    )
    return int(cursor.lastrowid)


def next_scheduled_occurrence(
    planned_at: datetime | None,
    interval: int,
    unit: str,
) -> datetime | None:
    if planned_at is None:
        return None
    if interval <= 0 or not unit or unit == "none":
        return None
    if unit == "days":
        return planned_at + timedelta(days=interval)
    if unit == "weeks":
        return planned_at + timedelta(weeks=interval)
    if unit != "months":
        return None
    month_index = planned_at.year * 12 + planned_at.month - 1 + interval
    year, month = divmod(month_index, 12)
    month += 1
    day = min(planned_at.day, calendar.monthrange(year, month)[1])
    return planned_at.replace(year=year, month=month, day=day)


def get_scheduled_events(status: str | None = None) -> list[sqlite3.Row]:
    db = get_db()
    query = """
        SELECT se.*, m.name, m.machine_type, m.serial_number
        FROM scheduled_events se
        JOIN machines m ON m.id = se.machine_id
        WHERE (se.event_type != 'Własne' OR se.created_by_user_id = ?)
    """
    user_id = current_user_id()
    params: list[Any] = [user_id if user_id is not None else -1]
    if status is not None:
        query += " AND se.status = ?"
        params.append(status)
    query += " ORDER BY se.planned_at ASC, se.id DESC"
    return list(db.execute(query, params).fetchall())


def get_due_scheduled_events() -> list[sqlite3.Row]:
    now = datetime.now(UTC)
    db = get_db()
    rows = db.execute(
        """
        SELECT se.*, m.name, m.machine_type, m.serial_number
        FROM scheduled_events se
        JOIN machines m ON m.id = se.machine_id
        WHERE se.status = 'planned' AND date(se.planned_at) <= date(?)
          AND (se.event_type != 'Własne' OR se.created_by_user_id = ?)
        ORDER BY se.planned_at ASC, se.id DESC
        """,
        (now.date().isoformat(), current_user_id() if current_user_id() is not None else -1),
    ).fetchall()
    return list(rows)


def save_scheduled_event(form: dict[str, Any]) -> int:
    planned_at = parse_datetime(form.get("planned_at")) or parse_datetime(form.get("date"))
    if not planned_at:
        raise ValueError("Data planowanego zdarzenia jest wymagana.")
    machine_name = str(form.get("machine_name", "")).strip()
    machine_type = str(form.get("machine_type", "")).strip()
    serial_number = str(form.get("serial_number", "")).strip()
    if not all([machine_name, machine_type, serial_number]):
        raise ValueError("Maszyna, typ i numer seryjny są wymagane.")
    repeat_interval = int(form.get("repeat_interval") or 0)
    repeat_unit = str(form.get("repeat_unit") or "none").strip().lower()
    event_type = str(form.get("event_type", "")).strip() or "Interwencja"
    if event_type not in SCHEDULED_EVENT_TYPES or event_type == "Awaria":
        raise ValueError("Wybierz prawidłowy rodzaj zdarzenia planowanego.")
    creator_user_id = current_user_id()
    if event_type == "Własne" and creator_user_id is None:
        raise ValueError("Aby zaplanować zdarzenie własne, zaloguj się na konto użytkownika.")
    if repeat_interval < 0:
        raise ValueError("Interwał powtarzania musi być nieujemny.")
    if repeat_unit not in {"none", "days", "weeks", "months"}:
        repeat_unit = "none"
    db = get_db()
    machine_id = get_or_create_machine(machine_name, machine_type, serial_number)
    cursor = db.execute(
        """
        INSERT INTO scheduled_events(
            machine_id, created_by_user_id, event_type, description, cause, repair_method, replaced_parts, notes, planned_at,
            repeat_interval, repeat_unit, status, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            machine_id,
            creator_user_id,
            event_type,
            str(form.get("description", "")).strip(),
            str(form.get("cause", "")).strip() or None,
            str(form.get("repair_method", "")).strip() or None,
            str(form.get("replaced_parts", "")).strip() or None,
            str(form.get("notes", "")).strip() or None,
            planned_at.isoformat(sep=" ", timespec="minutes"),
            repeat_interval,
            repeat_unit,
            str(form.get("status") or "planned").strip() or "planned",
            datetime.now(UTC).isoformat(timespec="seconds"),
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    log_user_activity(
        "create_scheduled_event",
        f"Zaplanowano zdarzenie #{cursor.lastrowid} dla maszyny {machine_name}.",
        target_type="scheduled_event",
        target_id=int(cursor.lastrowid),
    )
    return int(cursor.lastrowid)


def update_scheduled_event(event_id: int, form: dict[str, Any]) -> None:
    event = get_db().execute(
        "SELECT * FROM scheduled_events WHERE id = ?",
        (event_id,),
    ).fetchone()
    if event is None or not scheduled_event_is_visible_to_current_user(event):
        raise ValueError(f"Zaplanowane zdarzenie #{event_id} nie istnieje.")
    description_value = str(form.get("description", "")).strip()
    if not description_value:
        raise ValueError("Opis zdarzenia jest wymagany.")
    planned_at = parse_datetime(form.get("planned_at")) or parse_datetime(form.get("date"))
    if not planned_at:
        raise ValueError("Data planowanego zdarzenia jest wymagana.")
    repeat_interval = max(0, int(form.get("repeat_interval") or 0))
    repeat_unit = str(form.get("repeat_unit") or "none").strip().lower()
    if repeat_unit not in {"none", "days", "weeks", "months"}:
        repeat_unit = "none"
    event_type = str(form.get("event_type", "")).strip() or "Interwencja"
    if event_type not in SCHEDULED_EVENT_TYPES or event_type == "Awaria":
        raise ValueError("Wybierz prawidłowy rodzaj zdarzenia planowanego.")
    created_by_user_id = event["created_by_user_id"]
    if event_type == "Własne" and created_by_user_id is None:
        created_by_user_id = current_user_id()
        if created_by_user_id is None:
            raise ValueError("Aby zaplanować zdarzenie własne, zaloguj się na konto użytkownika.")
    machine_name = str(form.get("machine_name", "")).strip()
    if machine_name:
        machine_type = str(form.get("machine_type", "")).strip()
        serial_number = str(form.get("serial_number", "")).strip()
        if not all([machine_type, serial_number]):
            raise ValueError("Typ i numer seryjny są wymagane, jeśli zmieniasz maszynę.")
        machine_id = get_or_create_machine(machine_name, machine_type, serial_number)
        db = get_db()
        db.execute(
            "UPDATE scheduled_events SET machine_id = ?, created_by_user_id = ?, event_type = ?, description = ?, cause = ?, repair_method = ?, replaced_parts = ?, notes = ?, planned_at = ?, repeat_interval = ?, repeat_unit = ?, updated_at = ? WHERE id = ?",
            (
                machine_id,
                created_by_user_id,
                event_type,
                str(form.get("description", "")).strip(),
                str(form.get("cause", "")).strip() or None,
                str(form.get("repair_method", "")).strip() or None,
                str(form.get("replaced_parts", "")).strip() or None,
                str(form.get("notes", "")).strip() or None,
                planned_at.isoformat(sep=" ", timespec="minutes"),
                repeat_interval,
                repeat_unit,
                datetime.now(UTC).isoformat(timespec="seconds"),
                event_id,
            ),
        )
    else:
        db = get_db()
        db.execute(
            "UPDATE scheduled_events SET created_by_user_id = ?, event_type = ?, description = ?, cause = ?, repair_method = ?, replaced_parts = ?, notes = ?, planned_at = ?, repeat_interval = ?, repeat_unit = ?, updated_at = ? WHERE id = ?",
            (
                created_by_user_id,
                event_type,
                str(form.get("description", "")).strip(),
                str(form.get("cause", "")).strip() or None,
                str(form.get("repair_method", "")).strip() or None,
                str(form.get("replaced_parts", "")).strip() or None,
                str(form.get("notes", "")).strip() or None,
                planned_at.isoformat(sep=" ", timespec="minutes"),
                repeat_interval,
                repeat_unit,
                datetime.now(UTC).isoformat(timespec="seconds"),
                event_id,
            ),
        )
    db.commit()
    log_user_activity(
        "update_scheduled_event",
        f"Zaktualizowano zaplanowane zdarzenie #{event_id}.",
        target_type="scheduled_event",
        target_id=event_id,
    )


def delete_scheduled_event(event_id: int) -> None:
    db = get_db()
    event = db.execute(
        "SELECT * FROM scheduled_events WHERE id = ?",
        (event_id,),
    ).fetchone()
    if event is None or not scheduled_event_is_visible_to_current_user(event):
        raise ValueError(f"Zaplanowane zdarzenie #{event_id} nie istnieje.")
    db.execute("DELETE FROM scheduled_events WHERE id = ?", (event_id,))
    db.commit()
    log_user_activity(
        "delete_scheduled_event",
        f"Usunięto zaplanowane zdarzenie #{event_id}.",
        target_type="scheduled_event",
        target_id=event_id,
    )


def query_records_latest(limit: int | str | None = 20) -> list[sqlite3.Row]:
    db = get_db()
    query = """
        SELECT r.*, m.name, m.machine_type, m.serial_number,
            (SELECT COUNT(*) FROM attachments a WHERE a.record_id = r.id) AS attachment_count
        FROM repair_records r
        JOIN machines m ON m.id = r.machine_id
        ORDER BY r.occurrence_at DESC, r.id DESC
    """
    params: list[Any] = []
    if limit is not None and limit != "all":
        query += " LIMIT ?"
        params.append(int(limit))
    return list(db.execute(query, params).fetchall())


def query_dashboard_kpis() -> dict[str, int | float]:
    """Summarize records from the current calendar month, independent of the table limit."""
    now = datetime.now()
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if month_start.month == 12:
        next_month = month_start.replace(year=month_start.year + 1, month=1)
    else:
        next_month = month_start.replace(month=month_start.month + 1)

    row = get_db().execute(
        """
        SELECT
            SUM(CASE WHEN lower(event_type) = lower('Awaria') THEN 1 ELSE 0 END) AS failures,
            COALESCE(SUM(downtime_hours), 0) AS downtime,
            COALESCE(SUM(total_cost), 0) AS cost,
            SUM(CASE WHEN repaired_at IS NULL OR repaired_at = '' THEN 1 ELSE 0 END) AS unrepaired
        FROM repair_records
        WHERE occurrence_at >= ? AND occurrence_at < ?
        """,
        (month_start.isoformat(sep=" ", timespec="minutes"), next_month.isoformat(sep=" ", timespec="minutes")),
    ).fetchone()
    return {
        "failures": row["failures"] or 0,
        "downtime": row["downtime"],
        "cost": row["cost"],
        "unrepaired": row["unrepaired"] or 0,
    }


def query_records(filters: dict[str, str]) -> list[sqlite3.Row]:
    db = get_db()
    query = """
        SELECT
            r.*,
            m.name,
            m.machine_type,
            m.serial_number,
            (SELECT COUNT(*) FROM attachments a WHERE a.record_id = r.id) AS attachment_count
        FROM repair_records r
        JOIN machines m ON m.id = r.machine_id
        WHERE 1 = 1
    """
    params: list[Any] = []
    if filters.get("machine"):
        query += " AND m.name = ?"
        params.append(filters["machine"])
    if filters.get("machine_type"):
        query += " AND m.machine_type = ?"
        params.append(filters["machine_type"])
    if filters.get("serial_number"):
        query += " AND m.serial_number = ?"
        params.append(filters["serial_number"])
    if filters.get("event_type"):
        query += " AND r.event_type = ?"
        params.append(filters["event_type"])
    if filters.get("occurrence_from"):
        query += " AND date(r.occurrence_at) >= date(?)"
        occurrence_from = parse_datetime(filters["occurrence_from"])
        params.append(occurrence_from.date().isoformat() if occurrence_from else filters["occurrence_from"])
    if filters.get("occurrence_to"):
        query += " AND date(r.occurrence_at) <= date(?)"
        occurrence_to = parse_datetime(filters["occurrence_to"])
        params.append(occurrence_to.date().isoformat() if occurrence_to else filters["occurrence_to"])
    description_terms = [term.strip() for term in filters.get("description_terms", "").split(";") if term.strip()]
    if description_terms:
        text_fields = ["r.description", "r.cause", "r.repair_method", "r.replaced_parts", "r.notes"]
        term_clauses = []
        for term in description_terms:
            field_conditions = " OR ".join(
                f"LOWER(COALESCE({f}, '')) LIKE ?" for f in text_fields
            )
            term_clauses.append(f"({field_conditions})")
            params.extend([f"%{term.lower()}%"] * len(text_fields))
        query += " AND (" + " OR ".join(term_clauses) + ")"
    query += " ORDER BY r.occurrence_at DESC, r.id DESC"
    return list(db.execute(query, params).fetchall())


def list_machine_names() -> list[str]:
    db = get_db()
    rows = db.execute("SELECT DISTINCT name FROM machines ORDER BY name").fetchall()
    return [str(row["name"]) for row in rows]


def list_machine_types(machine_name: str) -> list[str]:
    if not machine_name:
        return []
    db = get_db()
    rows = db.execute(
        """
        SELECT DISTINCT machine_type
        FROM machines
        WHERE name = ?
        ORDER BY machine_type
        """,
        (machine_name,),
    ).fetchall()
    return [str(row["machine_type"]) for row in rows]


def list_serial_numbers(machine_name: str, machine_type: str) -> list[str]:
    if not machine_name or not machine_type:
        return []
    db = get_db()
    rows = db.execute(
        """
        SELECT DISTINCT serial_number
        FROM machines
        WHERE name = ? AND machine_type = ?
        ORDER BY serial_number
        """,
        (machine_name, machine_type),
    ).fetchall()
    return [str(row["serial_number"]) for row in rows]


def list_previous_entries(machine_name: str, machine_type: str = "", serial_number: str = "") -> list[dict[str, Any]]:
    if not machine_name:
        return []
    db = get_db()
    query = """
        SELECT
            r.id,
            r.occurrence_at,
            r.event_type,
            r.description,
            r.cause,
            r.repair_method,
            r.replaced_parts,
            m.machine_type,
            m.serial_number
        FROM repair_records r
        JOIN machines m ON m.id = r.machine_id
        WHERE m.name = ?
    """
    params: list[Any] = [machine_name]
    if machine_type:
        query += " AND m.machine_type = ?"
        params.append(machine_type)
    if serial_number:
        query += " AND m.serial_number = ?"
        params.append(serial_number)
    query += " ORDER BY r.occurrence_at DESC, r.id DESC"
    rows = db.execute(query, params).fetchall()
    return [
        {
            "id": int(row["id"]),
            "occurrence_at": format_datetime(row["occurrence_at"]),
            "event_type": row["event_type"],
            "description": row["description"],
            "cause": row["cause"] or "",
            "repair_method": row["repair_method"] or "",
            "replaced_parts": row["replaced_parts"] or "",
            "machine_type": row["machine_type"],
            "serial_number": row["serial_number"],
        }
        for row in rows
    ]


def update_record(record_id: int, data: dict[str, Any]) -> None:
    db = get_db()
    if not db.execute("SELECT id FROM repair_records WHERE id = ?", (record_id,)).fetchone():
        raise ValueError(f"Rekord #{record_id} nie istnieje.")
    occurrence_at = parse_datetime(data.get("occurrence_at"))
    if not occurrence_at:
        raise ValueError("Data wystąpienia jest wymagana.")
    repaired_at = parse_datetime(data.get("repaired_at"))
    machine_id = get_or_create_machine(
        str(data.get("machine_name", "")).strip(),
        str(data.get("machine_type", "")).strip(),
        str(data.get("serial_number", "")).strip(),
    )
    downtime_hours = calculate_downtime_hours(occurrence_at, repaired_at)
    labor_cost = as_float(data.get("labor_cost"))
    material_cost = as_float(data.get("material_cost"))
    spare_parts_cost = as_float(data.get("spare_parts_cost"))
    service_cost = as_float(data.get("service_cost"))
    total_cost = round(
        (labor_cost * get_labor_rate()) + material_cost + spare_parts_cost + service_cost, 2
    )
    db.execute(
        """
        UPDATE repair_records SET
            machine_id = ?, event_type = ?, description = ?, cause = ?,
            repair_method = ?, replaced_parts = ?, participants = ?, occurrence_at = ?,
            repaired_at = ?, machine_mtg = ?, downtime_hours = ?,
            labor_cost = ?, material_cost = ?, spare_parts_cost = ?,
            service_cost = ?, total_cost = ?, notes = ?
        WHERE id = ?
        """,
        (
            machine_id,
            str(data.get("event_type", "")).strip() or "Inne",
            str(data.get("description", "")).strip(),
            str(data.get("cause", "")).strip(),
            str(data.get("repair_method", "")).strip(),
            str(data.get("replaced_parts", "")).strip(),
            str(data.get("participants", "")).strip() or None,
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(data.get("machine_mtg", "")).strip(),
            downtime_hours,
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            str(data.get("notes", "")).strip() or None,
            record_id,
        ),
    )
    db.commit()


def import_workbook(binary_content: bytes, filename: str = "") -> tuple[int, int, int, int]:
    """Import an Excel file or an application bundle with its attachments."""
    if binary_content[:2] == b"PK":
        try:
            with zipfile.ZipFile(io.BytesIO(binary_content)) as archive:
                names = set(archive.namelist())
                if "manifest.json" in names and "records.xlsx" in names:
                    return _import_records_bundle(archive)
        except zipfile.BadZipFile:
            pass
        imported, skipped = _import_xlsx_bytes(binary_content)
        return imported, skipped, 0, 0
    if binary_content[:4] == b"\xd0\xcf\x11\xe0":
        imported, skipped = _import_xls_bytes(binary_content)
        return imported, skipped, 0, 0
    # Fall back to file extension when magic bytes are ambiguous
    if filename.lower().endswith(".xls"):
        imported, skipped = _import_xls_bytes(binary_content)
    else:
        imported, skipped = _import_xlsx_bytes(binary_content)
    return imported, skipped, 0, 0


def _import_records_bundle(archive: zipfile.ZipFile) -> tuple[int, int, int, int]:
    from pathlib import PurePosixPath

    try:
        manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError, KeyError) as exc:
        raise ValueError("Nieprawidłowy manifest paczki importu.") from exc
    if manifest.get("format") != "repair-records-bundle" or manifest.get("version") != 1:
        raise ValueError("Nieobsługiwany format paczki importu.")
    manifest_attachments = manifest.get("attachments", [])
    if not isinstance(manifest_attachments, list):
        raise ValueError("Nieprawidłowa lista załączników w paczce importu.")

    attachment_data = []
    archive_names = set(archive.namelist())
    for item in manifest_attachments:
        if not isinstance(item, dict):
            raise ValueError("Nieprawidłowy wpis załącznika w paczce importu.")
        try:
            source_record_id = int(item["record_id"])
            member_name = str(item["path"])
            pure_path = PurePosixPath(member_name)
            if (pure_path.is_absolute() or ".." in pure_path.parts or len(pure_path.parts) != 3
                    or pure_path.parts[0] != "attachments"):
                raise ValueError
            if member_name not in archive_names:
                raise ValueError
            content = archive.read(member_name)
            attachment_data.append((source_record_id, item, content))
        except (KeyError, TypeError, ValueError, zipfile.BadZipFile) as exc:
            raise ValueError("Paczka zawiera nieprawidłowy lub brakujący załącznik.") from exc

    workbook = load_workbook(io.BytesIO(archive.read("records.xlsx")), data_only=True)
    sheet = workbook["ZBIORCZE"] if "ZBIORCZE" in workbook.sheetnames else workbook.active
    record_mapping: dict[int, int] = {}
    imported, skipped = _process_import_rows(
        list(sheet.iter_rows(values_only=True)), source_record_mapping=record_mapping
    )
    attachments_added = 0
    attachments_skipped = 0
    for source_record_id, item, content in attachment_data:
        target_record_id = record_mapping.get(source_record_id)
        if target_record_id is None:
            attachments_skipped += 1
            continue
        if save_imported_attachment(
            target_record_id,
            content,
            str(item.get("original_name") or item.get("filename") or "file"),
            str(item.get("content_type") or "application/octet-stream"),
        ):
            attachments_added += 1
        else:
            attachments_skipped += 1
    return imported, skipped, attachments_added, attachments_skipped


def _import_xlsx_bytes(binary_content: bytes) -> tuple[int, int]:
    workbook = load_workbook(io.BytesIO(binary_content), data_only=True)
    sheet = workbook["ZBIORCZE"] if "ZBIORCZE" in workbook.sheetnames else workbook.active
    return _process_import_rows(list(sheet.iter_rows(values_only=True)))


def _import_xls_bytes(binary_content: bytes) -> tuple[int, int]:
    try:
        import xlrd
    except ImportError as exc:
        raise ValueError(
            "Biblioteka xlrd jest wymagana do importu plików .xls. "
            "Zainstaluj ją: pip install xlrd"
        ) from exc

    wb = xlrd.open_workbook(file_contents=binary_content)
    sheet_names = wb.sheet_names()
    sheet = wb.sheet_by_name("ZBIORCZE") if "ZBIORCZE" in sheet_names else wb.sheet_by_index(0)

    rows: list[tuple[Any, ...]] = []
    for row_idx in range(sheet.nrows):
        row_data: list[Any] = []
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    dt_tuple = xlrd.xldate_as_tuple(cell.value, wb.datemode)
                    row_data.append(
                        datetime(int(dt_tuple[0]), int(dt_tuple[1]), int(dt_tuple[2]),
                                 int(dt_tuple[3]), int(dt_tuple[4]), int(dt_tuple[5]))
                        if dt_tuple[0] else None
                    )
                except Exception:
                    row_data.append(cell.value)
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                row_data.append(None)
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                v = cell.value
                try:
                    iv = int(v)
                    row_data.append(iv if iv == v else v)
                except (ValueError, OverflowError):
                    row_data.append(v)
            else:
                row_data.append(cell.value)
        rows.append(tuple(row_data))
    return _process_import_rows(rows)


def _process_import_rows(
    rows: list[tuple[Any, ...]],
    source_record_mapping: dict[int, int] | None = None,
) -> tuple[int, int]:
    if not rows:
        return 0, 0
    # Find first row that contains at least 3 recognised column headers
    header_row_index = 0
    for i, row in enumerate(rows[:10]):
        matches = sum(
            1 for h in row
            if h is not None and HEADER_ALIASES.get(normalize_header(h))
        )
        if matches >= 3:
            header_row_index = i
            break

    header_map: dict[str, int] = {}
    for index, header in enumerate(rows[header_row_index]):
        if header is None:
            continue
        canonical_name = HEADER_ALIASES.get(normalize_header(header))
        if canonical_name:
            header_map[canonical_name] = index

    imported = 0
    skipped = 0
    for values in rows[header_row_index + 1:]:
        if not any(value not in (None, "") for value in values):
            continue
        occurrence_at = parse_datetime(row_value(values, header_map, "Data wystąpienia", None))
        if not occurrence_at:
            skipped += 1
            continue
        record = {
            "machine_name": row_value(values, header_map, "Maszyna", "") or "",
            "machine_type": row_value(values, header_map, "Typ", "") or "",
            "serial_number": row_value(values, header_map, "Numer", "") or "",
            "event_type": row_value(values, header_map, "Rodzaj zdarzenia", "Inne") or "Inne",
            "description": row_value(values, header_map, "Opis", "") or "",
            "cause": row_value(values, header_map, "Przyczyna", "") or "",
            "repair_method": row_value(values, header_map, "Sposób naprawy", "") or "",
            "replaced_parts": row_value(values, header_map, "Wymienione części", "") or "",
            "participants": row_value(values, header_map, "Osoby biorące udział w interwencji", "") or "",
            "occurrence_at": occurrence_at,
            "repaired_at": parse_datetime(row_value(values, header_map, "Data naprawienia/ interwencji", None)),
            "machine_mtg": row_value(values, header_map, "MTG maszyny", "") or "",
            "labor_cost": row_value(values, header_map, "Praco chłonność [80PLN/h]", 0) or 0,
            "material_cost": row_value(values, header_map, "Materiał własny PLN", 0) or 0,
            "spare_parts_cost": row_value(values, header_map, "Części zamienne PLN", 0) or 0,
            "service_cost": row_value(values, header_map, "Serwis PLN", 0) or 0,
            "notes": row_value(values, header_map, "Uwagi", "") or "",
        }
        try:
            record_id, created = save_imported_record(record)
            source_record_id = row_value(values, header_map, "Nr", None)
            try:
                if source_record_mapping is not None and source_record_id not in (None, ""):
                    source_record_mapping[int(float(source_record_id))] = record_id
            except (TypeError, ValueError, OverflowError):
                pass
            if created:
                imported += 1
            else:
                skipped += 1
        except ValueError:
            skipped += 1
    return imported, skipped


def save_imported_record(record: dict[str, Any]) -> tuple[int, bool]:
    occurrence_at = record["occurrence_at"]
    repaired_at = record.get("repaired_at")
    if not isinstance(occurrence_at, datetime):
        raise ValueError("Nieprawidłowa data wystąpienia w imporcie.")
    record_year, record_week, _ = occurrence_at.isocalendar()
    machine_id = get_or_create_machine(
        str(record.get("machine_name", "")).strip(),
        str(record.get("machine_type", "")).strip(),
        str(record.get("serial_number", "")).strip(),
    )
    labor_cost = as_float(record.get("labor_cost"))
    material_cost = as_float(record.get("material_cost"))
    spare_parts_cost = as_float(record.get("spare_parts_cost"))
    service_cost = as_float(record.get("service_cost"))
    total_cost = round((labor_cost * get_labor_rate()) + material_cost + spare_parts_cost + service_cost, 2)
    db = get_db()
    occurrence_value = occurrence_at.isoformat(sep=" ", timespec="minutes")
    repaired_value = repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None
    event_type = str(record.get("event_type", "Inne")).strip() or "Inne"
    description = str(record.get("description", "")).strip()
    cause = str(record.get("cause", "")).strip()
    repair_method = str(record.get("repair_method", "")).strip()
    replaced_parts = str(record.get("replaced_parts", "")).strip()
    participants = str(record.get("participants", "")).strip() or None
    machine_mtg = str(record.get("machine_mtg", "")).strip()
    notes = str(record.get("notes", "")).strip() or None

    duplicate = db.execute(
        """
        SELECT id FROM repair_records
        WHERE machine_id = ? AND occurrence_at = ? AND event_type = ?
          AND description = ? AND COALESCE(cause, '') = ?
          AND COALESCE(repair_method, '') = ? AND COALESCE(replaced_parts, '') = ?
          AND COALESCE(participants, '') = ? AND COALESCE(repaired_at, '') = ?
          AND COALESCE(machine_mtg, '') = ? AND COALESCE(notes, '') = ?
          AND labor_cost = ? AND material_cost = ? AND spare_parts_cost = ?
          AND service_cost = ?
        LIMIT 1
        """,
        (
            machine_id, occurrence_value, event_type, description, cause,
            repair_method, replaced_parts, participants or "", repaired_value or "",
            machine_mtg, notes or "", labor_cost, material_cost, spare_parts_cost,
            service_cost,
        ),
    ).fetchone()
    if duplicate:
        return int(duplicate["id"]), False

    cursor = db.execute(
        """
        INSERT INTO repair_records(
            record_year, record_week, week_sequence, machine_id, event_type,
            description, cause, repair_method, replaced_parts, participants, occurrence_at,
            repaired_at, machine_mtg, downtime_hours, labor_cost, material_cost,
            spare_parts_cost, service_cost, total_cost, notes, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record_year,
            record_week,
            next_week_sequence(record_year, record_week),
            machine_id,
            event_type,
            description,
            cause,
            repair_method,
            replaced_parts,
            participants,
            occurrence_value,
            repaired_value,
            machine_mtg,
            calculate_downtime_hours(occurrence_at, repaired_at),
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            notes,
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    log_user_activity(
        "create_record",
        f"Utworzono rekord #{cursor.lastrowid} podczas importu.",
        target_type="record",
        target_id=int(cursor.lastrowid),
    )
    return int(cursor.lastrowid), True


def save_imported_attachment(record_id: int, content: bytes, original_name: str, content_type: str) -> bool:
    from werkzeug.utils import secure_filename

    safe_name = secure_filename(original_name) or "file"
    db = get_db()
    upload_dir = Path(current_app_uploads()) / str(record_id)
    upload_dir.mkdir(parents=True, exist_ok=True)
    for existing in db.execute(
        "SELECT filename, original_name FROM attachments WHERE record_id = ?",
        (record_id,),
    ).fetchall():
        if existing["original_name"] != original_name:
            continue
        existing_path = upload_dir / existing["filename"]
        if existing_path.is_file() and existing_path.read_bytes() == content:
            return False

    destination = upload_dir / safe_name
    stem, suffix = Path(safe_name).stem, Path(safe_name).suffix
    counter = 1
    while destination.exists():
        destination = upload_dir / f"{stem}_{counter}{suffix}"
        counter += 1
    destination.write_bytes(content)
    try:
        db.execute(
            "INSERT INTO attachments(record_id, filename, original_name, content_type, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (record_id, destination.name, original_name, content_type,
             datetime.now(UTC).isoformat(timespec="seconds")),
        )
        db.commit()
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    return True


if __name__ == "__main__":
    import os

    create_app().run(
        host="0.0.0.0",
        port=5001,
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
    )
