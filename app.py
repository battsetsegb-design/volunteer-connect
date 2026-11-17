import io
import json
import os
import re
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path

from flask import (
    Flask, abort, flash, redirect, render_template, request,
    send_file, session, url_for,
)
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
DATABASE = BASE_DIR / "Register.db"
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")
SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "change-this-development-secret")
app.secret_key = SECRET_KEY
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true",
)

EMAIL_PATTERN = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)
ADMIN_PASSWORD_HASH = generate_password_hash(ADMIN_PASSWORD)
REGISTRATION_MIGRATION_COLUMNS = {
    "activity_interests": "TEXT NOT NULL DEFAULT '[]'",
    "availability": "TEXT NOT NULL DEFAULT ''",
    "motivation": "TEXT NOT NULL DEFAULT ''",
    "privacy_consent": "INTEGER NOT NULL DEFAULT 0",
    "consent_at": "TEXT",
}
ACTIVITY_CHOICES = (
    "Байгаль орчин, мод тарих", "Хүүхэд, боловсрол", "Ахмад настанд туслах",
    "Орчны цэвэрлэгээ", "Амьтны хамгаалал", "Бусад",
)
AVAILABILITY_CHOICES = (
    "Ажлын өдрийн өглөө", "Ажлын өдрийн өдөр", "Ажлын өдрийн орой",
    "Амралтын өдөр", "Цагийн хуваарь уян хатан",
)


def _connect_existing_database():
    """Open the existing database without creating an empty replacement."""
    if not DATABASE.is_file():
        raise FileNotFoundError(f"Өгөгдлийн сан олдсонгүй: {DATABASE}")
    connection = sqlite3.connect(DATABASE.resolve().as_uri() + "?mode=rw", uri=True, timeout=30)
    connection.row_factory = sqlite3.Row
    return connection


def backup_database():
    source = _connect_existing_database()
    backup_path = DATABASE.with_name(
        f"{DATABASE.stem}.backup-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}"
        f"-{secrets.token_hex(4)}{DATABASE.suffix}"
    )
    destination = None
    try:
        destination = sqlite3.connect(backup_path)
        source.backup(destination)
        destination.commit()
    except Exception:
        if destination is not None:
            destination.close()
        source.close()
        if backup_path.exists():
            backup_path.unlink()
        raise
    else:
        destination.close()
        source.close()
    return backup_path


def migrate_database():
    connection = _connect_existing_database()
    try:
        table_exists = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", ("registrations",)
        ).fetchone()
        if table_exists is None:
            raise RuntimeError("Register.db дотор registrations хүснэгт олдсонгүй.")
        existing = {row["name"] for row in connection.execute("PRAGMA table_info(registrations)")}
        missing = {name: definition for name, definition in REGISTRATION_MIGRATION_COLUMNS.items() if name not in existing}
        if not missing:
            return None
        backup_path = backup_database()
        connection.execute("BEGIN IMMEDIATE")
        for name, definition in missing.items():
            connection.execute(f"ALTER TABLE registrations ADD COLUMN {name} {definition}")
        connection.commit()
        app.logger.info("Migration columns added: %s; backup: %s", ", ".join(missing), backup_path)
        return backup_path
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def get_db():
    return _connect_existing_database()


def new_registration_token():
    token = secrets.token_urlsafe(32)
    session["registration_token"] = token
    return token


def get_admin_csrf_token():
    token = session.get("admin_csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["admin_csrf_token"] = token
    return token


def validate_admin_csrf_token():
    submitted = request.form.get("csrf_token", "")
    expected = session.get("admin_csrf_token", "")
    if not submitted or not expected or not secrets.compare_digest(submitted, expected):
        abort(400, description="Хүсэлтийн хамгаалалтын token хүчингүй байна. Хуудсаа шинэчлээд дахин оролдоно уу.")


def decode_interests(raw_value):
    try:
        decoded = json.loads(raw_value or "[]")
    except (TypeError, json.JSONDecodeError):
        return []
    return [str(item) for item in decoded] if isinstance(decoded, list) else []


def admin_required():
    if not session.get("is_admin"):
        return redirect(url_for("admin_login", next=request.path))
    return None


@app.route("/", methods=["GET", "POST"])
def index():
    values = {
        "last_name": "", "first_name": "", "phone": "", "email": "",
        "activity_interests": [], "availability": "", "motivation": "",
        "privacy_consent": False,
    }
    errors = []
    if request.method == "POST":
        values = {
            "last_name": request.form.get("last_name", "").strip(),
            "first_name": request.form.get("first_name", "").strip(),
            "phone": request.form.get("phone", "").strip(),
            "email": request.form.get("email", "").strip(),
            "activity_interests": request.form.getlist("activity_interests"),
            "availability": request.form.get("availability", "").strip(),
            "motivation": request.form.get("motivation", "").strip(),
            "privacy_consent": request.form.get("privacy_consent") == "yes",
        }
        submitted_token = request.form.get("registration_token", "")
        expected_token = session.pop("registration_token", None)
        if not (submitted_token and expected_token and secrets.compare_digest(submitted_token, expected_token)):
            errors.append("Маягтыг дахин илгээхээс өмнө хуудсаа шинэчилнэ үү.")
        if not values["last_name"]:
            errors.append("Овгоо оруулна уу.")
        elif len(values["last_name"]) > 100:
            errors.append("Овог 100 тэмдэгтээс урт байж болохгүй.")
        if not values["first_name"]:
            errors.append("Нэрээ оруулна уу.")
        elif len(values["first_name"]) > 100:
            errors.append("Нэр 100 тэмдэгтээс урт байж болохгүй.")
        digits = re.sub(r"\D", "", values["phone"])
        if not re.fullmatch(r"[0-9+()\-\s]{6,20}", values["phone"]) or not 6 <= len(digits) <= 15:
            errors.append("Утасны дугаараа зөв оруулна уу.")
        if len(values["email"]) > 254 or not EMAIL_PATTERN.fullmatch(values["email"]):
            errors.append("Имэйл хаягийн форматаа шалгана уу. Жишээ: ner@company.mn")
        selected = values["activity_interests"]
        if not selected:
            errors.append("Сонирхож буй ажлын чиглэлээс дор хаяж нэгийг сонгоно уу.")
        elif any(item not in ACTIVITY_CHOICES for item in selected):
            errors.append("Сонгосон ажлын чиглэлийн утга буруу байна.")
        elif len(selected) != len(set(selected)):
            errors.append("Ажлын чиглэлийг давхар сонгож болохгүй.")
        if values["availability"] not in AVAILABILITY_CHOICES:
            errors.append("Оролцох боломжтой цагаа сонгоно уу.")
        if not values["motivation"]:
            errors.append("Сайн дурын ажил хийх болсон шалтгаанаа бичнэ үү.")
        elif len(values["motivation"]) > 2000:
            errors.append("Шалтгаан 2000 тэмдэгтээс урт байж болохгүй.")
        if not values["privacy_consent"]:
            errors.append("Хувийн мэдээлэл ашиглах зөвшөөрөл өгөх шаардлагатай.")
        if not errors:
            consent_at = datetime.now().astimezone().isoformat(timespec="seconds")
            try:
                with get_db() as connection:
                    connection.execute(
                        """INSERT INTO registrations (last_name, first_name, phone, email, created_at,
                           activity_interests, availability, motivation, privacy_consent, consent_at)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (values["last_name"], values["first_name"], values["phone"], values["email"],
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                         json.dumps(selected, ensure_ascii=False), values["availability"],
                         values["motivation"], 1, consent_at),
                    )
                return redirect(url_for("index", success="1"))
            except sqlite3.Error:
                app.logger.exception("Бүртгэлийг SQLite-д хадгалж чадсангүй.")
                errors.append("Бүртгэх үед алдаа гарлаа. Таны оруулсан мэдээллийг хадгалсангүй.")
    return render_template(
        "index.html", values=values, errors=errors,
        success=request.args.get("success") == "1",
        activity_choices=ACTIVITY_CHOICES, availability_choices=AVAILABILITY_CHOICES,
        registration_token=new_registration_token(),
    )


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if session.get("is_admin"):
        return redirect(url_for("admin"))
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if secrets.compare_digest(username, ADMIN_USERNAME) and check_password_hash(ADMIN_PASSWORD_HASH, password):
            session.clear()
            session["is_admin"] = True
            next_url = request.args.get("next", "")
            if next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)
            return redirect(url_for("admin"))
        flash("Нэвтрэх нэр эсвэл нууц үг буруу байна.", "error")
    return render_template("login.html")


def get_activity_counts(connection):
    rows = connection.execute("SELECT activity_interests FROM registrations").fetchall()
    counts = {choice: 0 for choice in ACTIVITY_CHOICES}
    for row in rows:
        for interest in set(decode_interests(row["activity_interests"])):
            if interest in counts:
                counts[interest] += 1
    return counts


@app.route("/admin", methods=["GET"])
def admin():
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect
    search = request.args.get("q", "").strip()[:100]
    with get_db() as connection:
        total_count = connection.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
        activity_counts = get_activity_counts(connection)
        registrations = connection.execute(
            """SELECT id, last_name, first_name, phone, email, created_at,
                      activity_interests, availability
               FROM registrations ORDER BY id DESC"""
        ).fetchall()
    if search:
        # SQLite LIKE-ийн Unicode case-insensitive ажиллагаа хязгаарлагдмал.
        # Python str.casefold() нь Монгол кирилл болон латин үсгийн хувилбаруудыг
        # case-insensitive, хэсэгчилсэн байдлаар шалгана. Хайлтын зайг strip хийсэн.
        needle = search.casefold()
        search_fields = ("last_name", "first_name", "phone", "email")
        registrations = [
            row for row in registrations
            if any(needle in (row[field] or "").casefold() for field in search_fields)
        ]
    return render_template(
        "admin.html", registrations=registrations, total_count=total_count,
        activity_counts=activity_counts, search=search,
        csrf_token=get_admin_csrf_token(), detail=None,
    )


@app.route("/admin/registration/<int:registration_id>", methods=["GET"])
def admin_registration_detail(registration_id):
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect
    with get_db() as connection:
        detail = connection.execute(
            """SELECT id, last_name, first_name, phone, email, created_at,
                      activity_interests, availability, motivation, privacy_consent, consent_at
               FROM registrations WHERE id = ?""", (registration_id,),
        ).fetchone()
        registrations = connection.execute(
            """SELECT id, last_name, first_name, phone, email, created_at,
                      activity_interests, availability FROM registrations ORDER BY id DESC"""
        ).fetchall()
        total_count = connection.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
        activity_counts = get_activity_counts(connection)
    if detail is None:
        abort(404)
    return render_template(
        "admin.html", registrations=registrations, total_count=total_count,
        activity_counts=activity_counts, search="", csrf_token=get_admin_csrf_token(),
        detail=detail, detail_interests=decode_interests(detail["activity_interests"]),
    )


@app.route("/admin/registration/<int:registration_id>/delete", methods=["POST"])
def admin_delete_registration(registration_id):
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect
    validate_admin_csrf_token()
    with get_db() as connection:
        cursor = connection.execute("DELETE FROM registrations WHERE id = ?", (registration_id,))
        deleted = cursor.rowcount
    if not deleted:
        abort(404)
    flash("Бүртгэлийг амжилттай устгалаа.", "success")
    return redirect(url_for("admin"))


@app.route("/admin/export", methods=["GET"])
def export_excel():
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect
    with get_db() as connection:
        registrations = connection.execute(
            """SELECT id, last_name, first_name, phone, email, created_at,
                      activity_interests, availability, motivation, privacy_consent, consent_at
               FROM registrations ORDER BY id ASC"""
        ).fetchall()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Бүртгэл"
    sheet.append(["№", "Овог", "Нэр", "Утасны дугаар", "И-мэйл", "Сонирхож буй чиглэл",
                  "Оролцох боломжтой цаг", "Шалтгаан", "Хувийн мэдээлэл ашиглах зөвшөөрөл",
                  "Зөвшөөрөл өгсөн огноо", "Бүртгэсэн огноо"])
    for row in registrations:
        sheet.append([row["id"], row["last_name"], row["first_name"], row["phone"], row["email"],
                      ", ".join(decode_interests(row["activity_interests"])), row["availability"],
                      row["motivation"], "Тийм" if row["privacy_consent"] else "Үгүй",
                      row["consent_at"] or "", row["created_at"]])
    for column, width in {"A":8,"B":22,"C":22,"D":22,"E":32,"F":38,"G":28,"H":50,"I":30,"J":26,"K":22}.items():
        sheet.column_dimensions[column].width = width
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Register.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    if not session.get("is_admin"):
        return redirect(url_for("admin_login"))
    session.clear()
    flash("Системээс гарлаа.", "success")
    return redirect(url_for("admin_login"))


migrate_database()

if __name__ == "__main__":
    app.run(debug=True)
