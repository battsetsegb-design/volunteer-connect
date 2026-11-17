import io
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


def get_db():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with get_db() as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS registrations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                last_name TEXT NOT NULL,
                first_name TEXT NOT NULL,
                phone TEXT NOT NULL,
                email TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)


def admin_required():
    if not session.get("is_admin"):
        return redirect(url_for("admin_login", next=request.path))
    return None


@app.route("/", methods=["GET", "POST"])
def index():
    values = {"last_name": "", "first_name": "", "phone": "", "email": ""}
    errors = []

    if request.method == "POST":
        values = {
            "last_name": request.form.get("last_name", "").strip(),
            "first_name": request.form.get("first_name", "").strip(),
            "phone": request.form.get("phone", "").strip(),
            "email": request.form.get("email", "").strip(),
        }

        if not values["last_name"]:
            errors.append("Овгоо оруулна уу.")
        if not values["first_name"]:
            errors.append("Нэрээ оруулна уу.")

        digits = re.sub(r"\D", "", values["phone"])
        if not re.fullmatch(r"[0-9+()\-\s]{6,20}", values["phone"]) or not (
            6 <= len(digits) <= 15
        ):
            errors.append("Утасны дугаараа зөв оруулна уу.")

        if not EMAIL_PATTERN.fullmatch(values["email"]):
            errors.append("Имэйл хаягийн форматаа шалгана уу. Жишээ: ner@company.mn")

        if not errors:
            try:
                with get_db() as connection:
                    connection.execute(
                        """
                        INSERT INTO registrations
                            (last_name, first_name, phone, email, created_at)
                        VALUES (?, ?, ?, ?, ?)
                        """,
                        (
                            values["last_name"],
                            values["first_name"],
                            values["phone"],
                            values["email"],
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        ),
                    )
                return redirect(url_for("index", success="1"))
            except sqlite3.Error:
                app.logger.exception("Бүртгэлийг SQLite-д хадгалж чадсангүй.")
                errors.append("Бүртгэх үед алдаа гарлаа. Дахин оролдоно уу.")

    return render_template(
        "index.html",
        values=values,
        errors=errors,
        success=request.args.get("success") == "1",
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


init_db()

if __name__ == "__main__":
    app.run(debug=True)