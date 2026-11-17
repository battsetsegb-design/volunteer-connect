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

# Локал хөгжүүлэлтийн анхдагч нууц үг. Байршуулахдаа орчны хувьсагчаар солино.
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")
SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "change-this-development-secret")

app.secret_key = SECRET_KEY
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # HTTPS-тэй production орчинд SESSION_COOKIE_SECURE=true гэж тохируулна.
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true",
)

EMAIL_PATTERN = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)

# Нууц үгийг plaintext хэлбэрээр хадгалахгүй.
ADMIN_PASSWORD_HASH = generate_password_hash(ADMIN_PASSWORD)

# Эдгээр нь зөвхөн нэмэгдэх баганууд. Одоо байгаа бүртгэлийн багануудыг
# өөрчлөхгүй бөгөөд хуучин мөрүүд default утгатай хэвээр үлдэнэ.
REGISTRATION_MIGRATION_COLUMNS = {
    "activity_interests": "TEXT NOT NULL DEFAULT '[]'",
    "availability": "TEXT NOT NULL DEFAULT ''",
    "motivation": "TEXT NOT NULL DEFAULT ''",
    "privacy_consent": "INTEGER NOT NULL DEFAULT 0",
    "consent_at": "TEXT",
}

ACTIVITY_CHOICES = (
    "Байгаль орчин, мод тарих",
    "Хүүхэд, боловсрол",
    "Ахмад настанд туслах",
    "Орчны цэвэрлэгээ",
    "Амьтны хамгаалал",
    "Бусад",
)

AVAILABILITY_CHOICES = (
    "Ажлын өдрийн өглөө",
    "Ажлын өдрийн өдөр",
    "Ажлын өдрийн орой",
    "Амралтын өдөр",
    "Цагийн хуваарь уян хатан",
)


def _connect_existing_database():
    """Зөвхөн байгаа DB-г нээнэ; байхгүй бол шинээр хоосон файл үүсгэхгүй."""
    if not DATABASE.is_file():
        raise FileNotFoundError(
            f"Өгөгдлийн сан олдсонгүй: {DATABASE}. "
            "Одоо ашиглаж буй Register.db файлыг төслийн хавтсанд байрлуулна уу."
        )

    database_uri = DATABASE.resolve().as_uri() + "?mode=rw"
    connection = sqlite3.connect(database_uri, uri=True, timeout=30)
    connection.row_factory = sqlite3.Row
    return connection


def backup_database():
    """Одоо байгаа SQLite санг SQLite backup API ашиглан хуулбарлана."""
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
    """Хуучин өгөгдлийг хадгалж, шинэ Volunteer Connect багануудыг нэмнэ.

    Багана бүрийг PRAGMA table_info-оор шалгадаг тул дахин ажиллуулахад
    өөрчлөлт давхардахгүй. Бодит schema-г мэдэхгүй эсвэл хүснэгт байхгүй
    тохиолдолд шинэ/хоосон DB үүсгэхийн оронд алдаа өгч зогсоно.
    """
    connection = _connect_existing_database()
    backup_path = None
    try:
        table_exists = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            ("registrations",),
        ).fetchone()
        if table_exists is None:
            raise RuntimeError(
                "Register.db дотор registrations хүснэгт олдсонгүй; "
                "өгөгдлийн санг өөрчлөлгүй үлдээлээ."
            )

        existing_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(registrations)")
        }
        missing_columns = {
            name: definition
            for name, definition in REGISTRATION_MIGRATION_COLUMNS.items()
            if name not in existing_columns
        }

        # Бүх багана аль хэдийн байгаа бол backup дахин үүсгэхгүй, DB-г өөрчлөхгүй.
        if not missing_columns:
            return None

        # Өөрчлөлт хийхийн өмнөх DB-г тусдаа файлд бүрэн хуулна.
        backup_path = backup_database()

        connection.execute("BEGIN IMMEDIATE")
        for name, definition in missing_columns.items():
            # Багана/definition нь дээрх тогтмол dict-ээс авсан тул SQL-д хэрэглэж болно.
            connection.execute(
                f"ALTER TABLE registrations ADD COLUMN {name} {definition}"
            )
        connection.commit()
        app.logger.info(
            "SQLite migration амжилттай. Нэмсэн багана: %s; backup: %s",
            ", ".join(missing_columns),
            backup_path,
        )
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


def admin_required():
    if not session.get("is_admin"):
        return redirect(url_for("admin_login", next=request.path))
    return None


@app.route("/", methods=["GET", "POST"])
def index():
    values = {
        "last_name": "",
        "first_name": "",
        "phone": "",
        "email": "",
        "activity_interests": [],
        "availability": "",
        "motivation": "",
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
        valid_token = bool(
            submitted_token
            and expected_token
            and secrets.compare_digest(submitted_token, expected_token)
        )
        if not valid_token:
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
        if not re.fullmatch(r"[0-9+()\-\s]{6,20}", values["phone"]) or not (
            6 <= len(digits) <= 15
        ):
            errors.append("Утасны дугаараа зөв оруулна уу.")

        if (
            len(values["email"]) > 254
            or not EMAIL_PATTERN.fullmatch(values["email"])
        ):
            errors.append("Имэйл хаягийн форматаа шалгана уу. Жишээ: ner@company.mn")

        selected_interests = values["activity_interests"]
        if not selected_interests:
            errors.append("Сонирхож буй ажлын чиглэлээс дор хаяж нэгийг сонгоно уу.")
        elif any(choice not in ACTIVITY_CHOICES for choice in selected_interests):
            errors.append("Сонгосон ажлын чиглэлийн утга буруу байна.")
        elif len(selected_interests) != len(set(selected_interests)):
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
                        """
                        INSERT INTO registrations (
                            last_name, first_name, phone, email, created_at,
                            activity_interests, availability, motivation,
                            privacy_consent, consent_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            values["last_name"],
                            values["first_name"],
                            values["phone"],
                            values["email"],
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            json.dumps(selected_interests, ensure_ascii=False),
                            values["availability"],
                            values["motivation"],
                            1,
                            consent_at,
                        ),
                    )
                return redirect(url_for("index", success="1"))
            except sqlite3.Error:
                app.logger.exception("Бүртгэлийг SQLite-д хадгалж чадсангүй.")
                errors.append("Бүртгэх үед алдаа гарлаа. Таны оруулсан мэдээллийг хадгалсангүй.")

    # POST алдаатай байсан ч хэрэглэгч засварлаад дахин илгээх шинэ token олгоно.
    registration_token = new_registration_token()
    return render_template(
        "index.html",
        values=values,
        errors=errors,
        success=request.args.get("success") == "1",
        activity_choices=ACTIVITY_CHOICES,
        availability_choices=AVAILABILITY_CHOICES,
        registration_token=registration_token,
    )


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if session.get("is_admin"):
        return redirect(url_for("admin"))

    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")

        if (
            secrets.compare_digest(username, ADMIN_USERNAME)
            and check_password_hash(ADMIN_PASSWORD_HASH, password)
        ):
            session.clear()
            session["is_admin"] = True
            next_url = request.args.get("next", "")
            # Гаднын URL руу redirect хийхээс хамгаална.
            if next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)
            return redirect(url_for("admin"))

        flash("Нэвтрэх нэр эсвэл нууц үг буруу байна.", "error")

    return render_template("login.html")


@app.route("/admin", methods=["GET"])
def admin():
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect

    with get_db() as connection:
        registrations = connection.execute(
            """
            SELECT id, last_name, first_name, phone, email, created_at
            FROM registrations
            ORDER BY id DESC
            """
        ).fetchall()

    return render_template("admin.html", registrations=registrations)


@app.route("/admin/export", methods=["GET"])
def export_excel():
    auth_redirect = admin_required()
    if auth_redirect:
        return auth_redirect

    with get_db() as connection:
        registrations = connection.execute(
            """
            SELECT id, last_name, first_name, phone, email, created_at
            FROM registrations
            ORDER BY id ASC
            """
        ).fetchall()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Бүртгэл"
    sheet.append(["№", "Овог", "Нэр", "Утасны дугаар", "Имэйл", "Бүртгэсэн огноо"])

    for row in registrations:
        sheet.append([
            row["id"],
            row["last_name"],
            row["first_name"],
            row["phone"],
            row["email"],
            row["created_at"],
        ])

    for column, width in {"A": 8, "B": 22, "C": 22, "D": 22, "E": 32, "F": 22}.items():
        sheet.column_dimensions[column].width = width

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Register.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    if not session.get("is_admin"):
        return redirect(url_for("admin_login"))
    session.clear()
    flash("Системээс гарлаа.", "success")
    return redirect(url_for("admin_login"))


# Байгаа DB-д migration хийнэ. DB байхгүй эсвэл schema тохирохгүй бол
# автоматаар хоосон DB үүсгэхгүй, алдаатайгаар ил тод зогсоно.
migrate_database()

if __name__ == "__main__":
    app.run(debug=True)
