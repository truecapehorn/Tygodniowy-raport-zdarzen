

from __future__ import annotations

import io
import sqlite3
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from openpyxl import Workbook, load_workbook


EVENT_TYPES = [
    "Awaria",
    "Przegląd",
    "Modernizacja",
    "Interwencja",
    "Inne",
]

LABOR_RATE_PLN_PER_HOUR = 80

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
    "Data wystąpienia",
    "Data naprawienia/ interwencji",
    "MTG maszyny",
    "Wykluczenie z produkcji [rob.godzina]",
    "Praco chłonność [80PLN/h]",
    "Materiał własny PLN",
    "Części zamienne PLN",
    "Serwis PLN",
    "Koszt brutto naprawy PLN",
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
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS machines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    machine_type TEXT NOT NULL,
    serial_number TEXT NOT NULL,
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
    occurrence_at TEXT NOT NULL,
    repaired_at TEXT,
    machine_mtg TEXT,
    downtime_hours REAL NOT NULL DEFAULT 0,
    labor_cost REAL NOT NULL DEFAULT 0,
    material_cost REAL NOT NULL DEFAULT 0,
    spare_parts_cost REAL NOT NULL DEFAULT 0,
    service_cost REAL NOT NULL DEFAULT 0,
    total_cost REAL NOT NULL DEFAULT 0,
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
    return parsed.strftime("%Y-%m-%d %H:%M")


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


def calculate_downtime_hours(occurrence_at: datetime, repaired_at: datetime | None) -> float:
    if not repaired_at or repaired_at <= occurrence_at:
        return 0.0
    return round((repaired_at - occurrence_at).total_seconds() / 3600, 2)


def machine_key_from_form(form: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(form.get("machine_name", "")).strip(),
        str(form.get("machine_type", "")).strip(),
        str(form.get("serial_number", "")).strip(),
    )


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    app = Flask(__name__)
    default_db = Path(app.instance_path) / "repair_reports.sqlite3"
    app.config.from_mapping(
        SECRET_KEY="dev",
        DATABASE=str(default_db),
        UPLOAD_DIR=str(Path(app.instance_path) / "uploads"),
    )
    if test_config:
        app.config.update(test_config)

    Path(app.config["DATABASE"]).parent.mkdir(parents=True, exist_ok=True)
    Path(app.config["UPLOAD_DIR"]).mkdir(parents=True, exist_ok=True)

    @app.before_request
    def before_request() -> None:
        get_db()

    @app.teardown_appcontext
    def close_db(_: BaseException | None) -> None:
        connection = g.pop("db", None)
        if connection is not None:
            connection.close()

    with app.app_context():
        init_db()

    @app.route("/", methods=["GET", "POST"])
    def index():
        if request.method == "POST":
            try:
                save_record(request.form)
                flash("Rekord został zapisany.", "success")
                return redirect(url_for("index"))
            except ValueError as exc:
                flash(str(exc), "error")

        try:
            limit = max(1, min(int(request.args.get("limit", 20) or 20), 500))
        except (ValueError, TypeError):
            limit = 20
        filters = {
            "machine": request.args.get("machine", "").strip(),
            "machine_type": request.args.get("machine_type", "").strip(),
            "serial_number": request.args.get("serial_number", "").strip(),
            "event_type": request.args.get("event_type", "").strip(),
            "occurrence_from": request.args.get("occurrence_from", "").strip(),
            "occurrence_to": request.args.get("occurrence_to", "").strip(),
            "description_terms": request.args.get("description_terms", "").strip(),
        }
        filter_active = any(filters.values())
        return render_template(
            "index.html",
            latest_records=query_records_latest(limit),
            filtered_records=query_records(filters) if filter_active else [],
            filter_active=filter_active,
            machine_names=list_machine_names(),
            event_types=EVENT_TYPES,
            filters=filters,
            limit=limit,
        )

    @app.route("/new", methods=["GET", "POST"])
    def new_record():
        if request.method == "POST":
            try:
                record_id = save_record(request.form)
                for f in request.files.getlist("attachments"):
                    if f and f.filename:
                        save_attachment(record_id, f)
                flash("Rekord został zapisany.", "success")
                return redirect(url_for("index"))
            except ValueError as exc:
                flash(str(exc), "error")
        return render_template(
            "new_record.html",
            machine_names=list_machine_names(),
            event_types=EVENT_TYPES,
        )

    @app.route("/import", methods=["GET", "POST"])
    def import_data():
        if request.method == "GET":
            return redirect(url_for("data_page"), code=301)
        upload = request.files.get("file")
        if not upload or not upload.filename:
            flash("Wybierz plik Excel do importu.", "error")
            return redirect(url_for("data_page"))
        imported = import_workbook(upload.read(), upload.filename or "")
        flash(f"Zaimportowano {imported} rekordów.", "success")
        return redirect(url_for("index"))

    @app.route("/data", methods=["GET", "POST"])
    def data_page():
        if request.method == "POST":
            upload = request.files.get("file")
            if not upload or not upload.filename:
                flash("Wybierz plik Excel do importu.", "error")
                return redirect(url_for("data_page"))
            imported = import_workbook(upload.read(), upload.filename or "")
            flash(f"Zaimportowano {imported} rekordów.", "success")
            return redirect(url_for("index"))
        return render_template("import_export.html")

    @app.route("/export")
    def export_data():
        filters = {
            "machine": request.args.get("machine", "").strip(),
            "machine_type": request.args.get("machine_type", "").strip(),
            "serial_number": request.args.get("serial_number", "").strip(),
            "event_type": request.args.get("event_type", "").strip(),
            "occurrence_from": request.args.get("occurrence_from", "").strip(),
            "occurrence_to": request.args.get("occurrence_to", "").strip(),
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
                    format_datetime(record["occurrence_at"]),
                    format_datetime(record["repaired_at"]),
                    record["machine_mtg"],
                    record["downtime_hours"],
                    record["labor_cost"],
                    record["material_cost"],
                    record["spare_parts_cost"],
                    record["service_cost"],
                    record["total_cost"],
                ]
            )
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name="tygodniowy_raport_napraw.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.route("/api/machines")
    def machine_lookup():
        machine_name = request.args.get("machine", "").strip()
        machine_type = request.args.get("machine_type", "").strip()
        return jsonify(
            {
                "types": list_machine_types(machine_name),
                "serial_numbers": list_serial_numbers(machine_name, machine_type),
                "previous_entries": list_previous_entries(machine_name),
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
        return jsonify(record)

    @app.route("/api/record/<int:record_id>", methods=["PATCH"])
    def update_record_api(record_id: int):
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "Brak danych."}), 400
        try:
            update_record(record_id, data)
            return jsonify({"ok": True})
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400

    @app.route("/api/record/<int:record_id>", methods=["DELETE"])
    def delete_record_api(record_id: int):
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
        return jsonify({"ok": True})

    @app.route("/api/record/<int:record_id>/attachments", methods=["POST"])
    def add_record_attachments_api(record_id: int):
        if not get_db().execute("SELECT id FROM repair_records WHERE id = ?", (record_id,)).fetchone():
            return jsonify({"error": "Rekord nie istnieje."}), 404
        for f in request.files.getlist("attachments"):
            if f and f.filename:
                save_attachment(record_id, f)
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


def init_db() -> None:
    db = get_db()
    db.executescript(SCHEMA)
    db.commit()


def get_or_create_machine(name: str, machine_type: str, serial_number: str) -> int:
    if not all([name, machine_type, serial_number]):
        raise ValueError("Maszyna, typ i numer są wymagane.")
    db = get_db()
    existing = db.execute(
        """
        SELECT id FROM machines
        WHERE name = ? AND machine_type = ? AND serial_number = ?
        """,
        (name, machine_type, serial_number),
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
    total_cost = round((labor_cost * LABOR_RATE_PLN_PER_HOUR) + material_cost + spare_parts_cost + service_cost, 2)
    db = get_db()
    cursor = db.execute(
        """
        INSERT INTO repair_records(
            record_year, record_week, week_sequence, machine_id, event_type,
            description, cause, repair_method, replaced_parts, occurrence_at,
            repaired_at, machine_mtg, downtime_hours, labor_cost, material_cost,
            spare_parts_cost, service_cost, total_cost, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(form.get("machine_mtg", "")).strip(),
            downtime_hours,
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    return int(cursor.lastrowid)


def query_records_latest(limit: int = 20) -> list[sqlite3.Row]:
    db = get_db()
    return list(
        db.execute(
            """
            SELECT r.*, m.name, m.machine_type, m.serial_number
            FROM repair_records r
            JOIN machines m ON m.id = r.machine_id
            ORDER BY r.occurrence_at DESC, r.id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    )


def query_records(filters: dict[str, str]) -> list[sqlite3.Row]:
    db = get_db()
    query = """
        SELECT
            r.*,
            m.name,
            m.machine_type,
            m.serial_number
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
        params.append(filters["occurrence_from"])
    if filters.get("occurrence_to"):
        query += " AND date(r.occurrence_at) <= date(?)"
        params.append(filters["occurrence_to"])
    description_terms = [term.strip() for term in filters.get("description_terms", "").split(";") if term.strip()]
    if description_terms:
        text_fields = ["r.description", "r.cause", "r.repair_method", "r.replaced_parts"]
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


def list_previous_entries(machine_name: str) -> list[dict[str, Any]]:
    if not machine_name:
        return []
    db = get_db()
    rows = db.execute(
        """
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
        ORDER BY r.occurrence_at DESC, r.id DESC
        """,
        (machine_name,),
    ).fetchall()
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
        (labor_cost * LABOR_RATE_PLN_PER_HOUR) + material_cost + spare_parts_cost + service_cost, 2
    )
    db.execute(
        """
        UPDATE repair_records SET
            machine_id = ?, event_type = ?, description = ?, cause = ?,
            repair_method = ?, replaced_parts = ?, occurrence_at = ?,
            repaired_at = ?, machine_mtg = ?, downtime_hours = ?,
            labor_cost = ?, material_cost = ?, spare_parts_cost = ?,
            service_cost = ?, total_cost = ?
        WHERE id = ?
        """,
        (
            machine_id,
            str(data.get("event_type", "")).strip() or "Inne",
            str(data.get("description", "")).strip(),
            str(data.get("cause", "")).strip(),
            str(data.get("repair_method", "")).strip(),
            str(data.get("replaced_parts", "")).strip(),
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(data.get("machine_mtg", "")).strip(),
            downtime_hours,
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            record_id,
        ),
    )
    db.commit()


def import_workbook(binary_content: bytes, filename: str = "") -> int:
    """Detect format (XLS vs XLSX) and dispatch to the appropriate importer."""
    if binary_content[:2] == b"PK":
        return _import_xlsx_bytes(binary_content)
    if binary_content[:4] == b"\xd0\xcf\x11\xe0":
        return _import_xls_bytes(binary_content)
    # Fall back to file extension when magic bytes are ambiguous
    if filename.lower().endswith(".xls"):
        return _import_xls_bytes(binary_content)
    return _import_xlsx_bytes(binary_content)


def _import_xlsx_bytes(binary_content: bytes) -> int:
    workbook = load_workbook(io.BytesIO(binary_content), data_only=True)
    sheet = workbook["ZBIORCZE"] if "ZBIORCZE" in workbook.sheetnames else workbook.active
    return _process_import_rows(list(sheet.iter_rows(values_only=True)))


def _import_xls_bytes(binary_content: bytes) -> int:
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


def _process_import_rows(rows: list[tuple[Any, ...]]) -> int:
    if not rows:
        return 0
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
    for values in rows[header_row_index + 1:]:
        if not any(value not in (None, "") for value in values):
            continue
        occurrence_at = parse_datetime(row_value(values, header_map, "Data wystąpienia", None))
        if not occurrence_at:
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
            "occurrence_at": occurrence_at,
            "repaired_at": parse_datetime(row_value(values, header_map, "Data naprawienia/ interwencji", None)),
            "machine_mtg": row_value(values, header_map, "MTG maszyny", "") or "",
            "labor_cost": row_value(values, header_map, "Praco chłonność [80PLN/h]", 0) or 0,
            "material_cost": row_value(values, header_map, "Materiał własny PLN", 0) or 0,
            "spare_parts_cost": row_value(values, header_map, "Części zamienne PLN", 0) or 0,
            "service_cost": row_value(values, header_map, "Serwis PLN", 0) or 0,
        }
        try:
            save_imported_record(record)
            imported += 1
        except ValueError:
            continue
    return imported


def save_imported_record(record: dict[str, Any]) -> int:
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
    total_cost = round((labor_cost * LABOR_RATE_PLN_PER_HOUR) + material_cost + spare_parts_cost + service_cost, 2)
    db = get_db()
    cursor = db.execute(
        """
        INSERT INTO repair_records(
            record_year, record_week, week_sequence, machine_id, event_type,
            description, cause, repair_method, replaced_parts, occurrence_at,
            repaired_at, machine_mtg, downtime_hours, labor_cost, material_cost,
            spare_parts_cost, service_cost, total_cost, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record_year,
            record_week,
            next_week_sequence(record_year, record_week),
            machine_id,
            str(record.get("event_type", "Inne")).strip() or "Inne",
            str(record.get("description", "")).strip(),
            str(record.get("cause", "")).strip(),
            str(record.get("repair_method", "")).strip(),
            str(record.get("replaced_parts", "")).strip(),
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(record.get("machine_mtg", "")).strip(),
            calculate_downtime_hours(occurrence_at, repaired_at),
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            datetime.now(UTC).isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    return int(cursor.lastrowid)


if __name__ == "__main__":
    import os

    create_app().run(
        host="0.0.0.0",
        port=5001,
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
    )
