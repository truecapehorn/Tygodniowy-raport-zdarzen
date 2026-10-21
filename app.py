

from __future__ import annotations

import io
import json
import secrets
import sqlite3
from datetime import datetime, timezone
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
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
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
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(id)
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
        "MAX_UPLOAD_SIZE_MB": 50,
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
        INSERT INTO user_activity_log(user_id, username, action, target_type, target_id, details, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            username,
            action,
            target_type,
            target_id,
            details,
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


def get_user_history(limit: int = 20, offset: int = 0) -> list[sqlite3.Row]:
    db = get_db()
    return list(
        db.execute(
            "SELECT * FROM user_activity_log ORDER BY created_at DESC, id DESC LIMIT ? OFFSET ?",
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
            "current_username": username or "gość",
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
        activity_per_page = 20
        activity_total = db.execute("SELECT COUNT(*) AS total FROM user_activity_log").fetchone()["total"]
        activity_pages = max(1, (int(activity_total) + activity_per_page - 1) // activity_per_page)
        activity_page = min(activity_page, activity_pages)
        return render_template(
            "admin_users.html",
            users=users,
            edit_user=edit_user,
            recent_activity=get_user_history(activity_per_page, (activity_page - 1) * activity_per_page),
            activity_page=activity_page,
            activity_pages=activity_pages,
        )

    @app.route("/admin/machines")
    def admin_machines():
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
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
        if edit_machine_id:
            edit_machine = db.execute(
                "SELECT * FROM machines WHERE id = ?", (edit_machine_id,)
            ).fetchone()
        return render_template(
            "admin_machines.html", machines=machines, edit_machine=edit_machine
        )

    @app.route("/admin/machines/<int:machine_id>/edit", methods=["POST"])
    def admin_edit_machine(machine_id: int):
        if session.get("role") != "admin":
            flash("Brak uprawnień administratora.", "error")
            return redirect(url_for("index"))

        name = request.form.get("name", "").strip()
        machine_type = request.form.get("machine_type", "").strip()
        serial_number = request.form.get("serial_number", "").strip()
        if not all((name, machine_type, serial_number)):
            flash("Nazwa, typ i numer seryjny są wymagane.", "error")
            return redirect(url_for("admin_machines", edit=machine_id))

        db = get_db()
        try:
            cursor = db.execute(
                """
                UPDATE machines
                SET name = ?, machine_type = ?, serial_number = ?
                WHERE id = ?
                """,
                (name, machine_type, serial_number, machine_id),
            )
            if cursor.rowcount == 0:
                flash("Maszyna nie istnieje.", "error")
                return redirect(url_for("admin_machines"))
            db.commit()
        except sqlite3.IntegrityError:
            db.rollback()
            flash("Maszyna o takich danych lub numerze seryjnym już istnieje.", "error")
            return redirect(url_for("admin_machines", edit=machine_id))

        log_user_activity(
            "update_machine",
            f"Zmieniono dane maszyny {name} (ID {machine_id}).",
            target_type="machine",
            target_id=machine_id,
        )
        flash("Dane maszyny zostały zaktualizowane.", "success")
        return redirect(url_for("admin_machines"))

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
        sheet.append(["Data", "Użytkownik", "Akcja", "Typ obiektu", "ID obiektu", "Szczegóły"])
        for item in activity:
            sheet.append(
                [
                    format_activity_datetime(item["created_at"]),
                    item["username"] or "",
                    item["action"],
                    item["target_type"] or "",
                    item["target_id"] or "",
                    item["details"] or "",
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
        return render_template(
            "index.html",
            latest_records=query_records_latest(None if limit == "all" else limit),
            event_types=EVENT_TYPES,
            limit=limit,
        )

    @app.route("/new", methods=["GET", "POST"])
    def new_record():
        if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
            flash("Brak uprawnie\u0144 do dodawania rekord\u00f3w.", "error")
            return redirect(url_for("index"))
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
        if not app.config.get("TESTING") and not user_has_permission("can_create_records"):
            flash("Brak uprawnie\u0144 do importu danych.", "error")
            return redirect(url_for("data_page"))
        upload = request.files.get("file")
        if not upload or not upload.filename:
            flash("Wybierz plik Excel do importu.", "error")
            return redirect(url_for("data_page"))
        imported = import_workbook(upload.read(), upload.filename or "")
        log_user_activity(
            "import_records",
            f"Zaimportowano {imported} rekordów z pliku {upload.filename or 'bez nazwy'}.",
            target_type="import",
        )
        flash(f"Zaimportowano {imported} rekordów.", "success")
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
            imported = import_workbook(upload.read(), upload.filename or "")
            log_user_activity(
                "import_records",
                f"Zaimportowano {imported} rekordów z pliku {upload.filename or 'bez nazwy'}.",
                target_type="import",
            )
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
        story = [Paragraph(f"Zdarzenie #{record_id} — {text(record['name'])}", heading), table]
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
        return send_file(output, as_attachment=True, download_name=f"rekord_{record_id}.pdf", mimetype="application/pdf")

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
            db.execute("DELETE FROM machines WHERE id = ?", (int(machine["id"]),))
    db.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_machines_serial_number
        ON machines(serial_number)
        WHERE lower(trim(serial_number)) NOT IN ('', 'brak', '-', 'n/a', 'nie dotyczy')
        """
    )
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
            str(record.get("event_type", "Inne")).strip() or "Inne",
            str(record.get("description", "")).strip(),
            str(record.get("cause", "")).strip(),
            str(record.get("repair_method", "")).strip(),
            str(record.get("replaced_parts", "")).strip(),
            str(record.get("participants", "")).strip() or None,
            occurrence_at.isoformat(sep=" ", timespec="minutes"),
            repaired_at.isoformat(sep=" ", timespec="minutes") if repaired_at else None,
            str(record.get("machine_mtg", "")).strip(),
            calculate_downtime_hours(occurrence_at, repaired_at),
            labor_cost,
            material_cost,
            spare_parts_cost,
            service_cost,
            total_cost,
            str(record.get("notes", "")).strip() or None,
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
    return int(cursor.lastrowid)


if __name__ == "__main__":
    import os

    create_app().run(
        host="0.0.0.0",
        port=5001,
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
    )
