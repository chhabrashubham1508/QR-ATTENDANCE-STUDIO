import os 
import io
import csv
import base64
import secrets
import sqlite3
import hmac
import hashlib
import time
from functools import wraps
from datetime import datetime, timedelta, timezone

import qrcode
from openpyxl import Workbook
from openpyxl.styles import Font
from flask import (
    Flask, render_template, request, redirect, url_for, flash, jsonify, Response, abort
)
from flask_login import (
    LoginManager, UserMixin, login_user, login_required, logout_user, current_user
)
from flask_wtf.csrf import CSRFProtect
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(
    __name__,
    static_folder=STATIC_DIR,
    template_folder=TEMPLATES_DIR,
    static_url_path="/static"
)
app.config['SECRET_KEY'] = 'qr_india'
app.config["WTF_CSRF_TIME_LIMIT"] = 7200

csrf = CSRFProtect(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

DB = os.path.join(BASE_DIR, "attendance.db")

# Dynamic QR configuration: 10 seconds validity window with 3s grace window for network latency
QR_TIMEOUT = 10
QR_GRACE_PERIOD = 3
IST = timezone(timedelta(hours=5, minutes=30), "IST")
MAX_DURATION_MIN = 720          # 12 hours safety cap
ARCHIVED_REPORT_DAYS = 7

def parse_duration(form):
    """Read the duration dropdown and return the selected duration in minutes."""
    choice = form.get("duration", "60")
    if choice == "custom":
        raw_hours = form.get("custom_hours", "").strip()
        raw_minutes = form.get("custom_minutes", "").strip()
        if (raw_hours and not raw_hours.isdigit()) or (raw_minutes and not raw_minutes.isdigit()):
            raise ValueError("Enter custom hours and minutes as whole numbers.")
        hours = int(raw_hours or 0)
        minutes = int(raw_minutes or 0)
        if minutes > 59:
            raise ValueError("Custom minutes must be between 0 and 59.")
        minutes += hours * 60
    elif choice in ("30", "60", "120"):
        minutes = int(choice)
    else:
        raise ValueError("Invalid duration selected.")
    if not 1 <= minutes <= MAX_DURATION_MIN:
        raise ValueError(f"Duration must be between 1 and {MAX_DURATION_MIN} minutes.")
    return minutes

def close_expired_sessions(con, now=None):
    """Auto-close every active session whose lecture time has ended."""
    now = now or int(datetime.now(timezone.utc).timestamp())
    con.execute("""
        DELETE FROM dynamic_tokens
        WHERE session_id IN (SELECT id FROM sessions WHERE active = 1 AND ends_at <= ?)
    """, (now,))
    con.execute("UPDATE sessions SET active = 0 WHERE active = 1 AND ends_at <= ?", (now,))

def make_dynamic_token():
    """Generate a secure, dynamic, cryptographically random token for QR generation."""
    return secrets.token_urlsafe(32)

def generate_qr_base64(data_payload):
    """Generate a base64 encoded PNG data string for a QR code."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=3,
    )
    qr.add_data(data_payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0f172a", back_color="#ffffff")
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG")
    return base64.b64encode(buffer.getvalue()).decode("ascii")

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con

def init_db():
    with db() as con:
        # Check and migrate users check constraint if needed
        user_table = con.execute("SELECT sql FROM sqlite_master WHERE name='users'").fetchone()
        if user_table and "'teacher'" not in user_table["sql"]:
            con.execute("PRAGMA foreign_keys = OFF")
            con.execute("BEGIN TRANSACTION")
            con.execute("""
            CREATE TABLE IF NOT EXISTS users_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'teacher', 'student'))
            )
            """)
            con.execute("INSERT OR IGNORE INTO users_new SELECT id, username, password_hash, role FROM users")
            con.execute("DROP TABLE users")
            con.execute("ALTER TABLE users_new RENAME TO users")
            con.commit()
            con.execute("PRAGMA foreign_keys = ON")

        con.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','teacher','student'))
        );

        CREATE TABLE IF NOT EXISTS teachers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            department TEXT NOT NULL DEFAULT 'Computer Science',
            email TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            roll_no TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            token TEXT UNIQUE NOT NULL,
            created_at INTEGER NOT NULL,
            expires_at INTEGER NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS dynamic_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
            token TEXT UNIQUE NOT NULL,
            created_at INTEGER NOT NULL,
            expires_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_dynamic_tokens_lookup ON dynamic_tokens(token, expires_at);

        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
            student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
            marked_at INTEGER NOT NULL,
            UNIQUE(session_id, student_id)
        );

        CREATE TABLE IF NOT EXISTS archived_attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_title TEXT NOT NULL,
            roll_no TEXT NOT NULL,
            student_name TEXT NOT NULL,
            marked_at INTEGER NOT NULL,
            archived_at INTEGER NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_archived_attendance_retention
        ON archived_attendance(archived_at);

        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            author_name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'teacher',
            message TEXT NOT NULL,
            created_at INTEGER NOT NULL
        );
        """)

        cols = {r["name"] for r in con.execute("PRAGMA table_info(sessions)")}
        if "duration_minutes" not in cols:
            con.execute("ALTER TABLE sessions ADD COLUMN duration_minutes INTEGER NOT NULL DEFAULT 60")
        if "started_at" not in cols:
            con.execute("ALTER TABLE sessions ADD COLUMN started_at INTEGER")
            con.execute("UPDATE sessions SET started_at = created_at WHERE started_at IS NULL")
        if "ends_at" not in cols:
            con.execute("ALTER TABLE sessions ADD COLUMN ends_at INTEGER")
            con.execute("UPDATE sessions SET ends_at = started_at + duration_minutes * 60 WHERE ends_at IS NULL")

        # Seed admin teacher profile if missing
        admin_row = con.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()
        if admin_row:
            existing_teacher = con.execute("SELECT id FROM teachers WHERE user_id = ?", (admin_row["id"],)).fetchone()
            if not existing_teacher:
                con.execute("""
                    INSERT INTO teachers(user_id, name, department, email)
                    VALUES (?, 'System Administrator', 'Administration', 'admin@institution.edu')
                """, (admin_row["id"],))

# Initialize database schema on startup
init_db()

class User(UserMixin):
    def __init__(self, row, profile=None):
        self.id = str(row["id"])
        self.username = row["username"]
        self.role = row["role"]
        self.display_name = profile["name"] if profile and "name" in profile.keys() else row["username"]

@login_manager.user_loader
def load_user(user_id):
    with db() as con:
        row = con.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return None
        profile = None
        if row["role"] in ("admin", "teacher"):
            profile = con.execute("SELECT name FROM teachers WHERE user_id = ?", (user_id,)).fetchone()
        elif row["role"] == "student":
            profile = con.execute("SELECT name FROM students WHERE user_id = ?", (user_id,)).fetchone()
    return User(row, profile)

def role_required(*roles):
    def decorator(view):
        @wraps(view)
        @login_required
        def wrapped(*args, **kwargs):
            if current_user.role not in roles:
                abort(403)
            return view(*args, **kwargs)
        return wrapped
    return decorator

def create_account(username, password, role, name=None, roll_no=None, department=None, email=None):
    if not username.strip() or len(password) < 8:
        raise ValueError("Username is required; password must be at least 8 characters.")

    with db() as con:
        cur = con.execute(
            "INSERT INTO users(username, password_hash, role) VALUES (?, ?, ?)",
            (username.strip(), generate_password_hash(password), role)
        )
        user_id = cur.lastrowid
        if role == "student":
            if not name or not roll_no:
                raise ValueError("Student name and roll number are required.")
            con.execute(
                "INSERT INTO students(user_id, roll_no, name) VALUES (?, ?, ?)",
                (user_id, roll_no.strip(), name.strip())
            )
        elif role in ("teacher", "admin"):
            t_name = name.strip() if name else username.strip()
            con.execute(
                "INSERT INTO teachers(user_id, name, department, email) VALUES (?, ?, ?, ?)",
                (user_id, t_name, department.strip() if department else "Computer Science", email.strip() if email else "")
            )

def generate_session_token(con, session_id):
    """Generate and store a new dynamic token valid for QR_TIMEOUT seconds."""
    now = int(datetime.now(timezone.utc).timestamp())
    token = make_dynamic_token()
    expires_at = now + QR_TIMEOUT
    con.execute("""
        UPDATE sessions 
        SET token = ?, created_at = ?, expires_at = ?, active = 1
        WHERE id = ?
    """, (token, now, expires_at, session_id))
    con.execute("""
        INSERT OR IGNORE INTO dynamic_tokens(session_id, token, created_at, expires_at)
        VALUES (?, ?, ?, ?)
    """, (session_id, token, now, expires_at + QR_GRACE_PERIOD))
    # Purge old expired tokens older than 1 hour to keep DB lean
    con.execute("DELETE FROM dynamic_tokens WHERE expires_at < ?", (now - 3600,))
    return token, expires_at

@app.template_filter("datetimeformat")
def datetimeformat(value):
    if not value:
        return ""
    try:
        return datetime.fromtimestamp(int(value), IST).strftime("%d %b %Y, %I:%M %p IST")
    except Exception:
        return str(value)

@app.template_filter("timeago")
def timeago_filter(value):
    if not value:
        return ""
    try:
        now = int(datetime.now(timezone.utc).timestamp())
        diff = max(0, now - int(value))
        if diff < 60:
            return "just now"
        elif diff < 3600:
            return f"{diff // 60}m ago"
        elif diff < 86400:
            return f"{diff // 3600}h ago"
        else:
            return f"{diff // 86400}d ago"
    except Exception:
        return ""

@app.context_processor
def inject_clock():
    return {
        "namespace_now": int(datetime.now(timezone.utc).timestamp())
    }

@app.cli.command("init-demo")
def init_demo():
    """Create the database and sample accounts for testing."""
    init_db()
    with db() as con:
        exists = con.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
    if not exists:
        create_account("admin", "Admin123!", "admin", name="Administrator", department="Management")
    with db() as con:
        exists = con.execute("SELECT id FROM users WHERE username = ?", ("teacher1",)).fetchone()
    if not exists:
        create_account("teacher1", "Teacher123!", "teacher", name="Prof. Alan Turing", department="Computer Science", email="turing@college.edu")
    with db() as con:
        exists = con.execute("SELECT id FROM users WHERE username = ?", ("student1",)).fetchone()
    if not exists:
        create_account(
            "student1", "Student123!", "student",
            name="Demo Student", roll_no="BT001"
        )
    print("Demo accounts ready. Change their passwords before real use.")


@app.route("/")
def index():
    if current_user.is_authenticated:
        endpoint = "admin_dashboard" if current_user.role in ("admin", "teacher") else "student_dashboard"
        return redirect(url_for(endpoint))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with db() as con:
            row = con.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()

        if row and check_password_hash(row["password_hash"], password):
            profile = None
            if row["role"] in ("admin", "teacher"):
                profile = con.execute("SELECT name FROM teachers WHERE user_id = ?", (row["id"],)).fetchone()
            elif row["role"] == "student":
                profile = con.execute("SELECT name FROM students WHERE user_id = ?", (row["id"],)).fetchone()

            login_user(User(row, profile))
            return redirect(url_for(
                "admin_dashboard" if row["role"] in ("admin", "teacher")
                else "student_dashboard"
            ))
        flash("Invalid username or password.")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# --- Admin & Teacher Interactive Dashboard ---
@app.route("/admin")
@role_required("admin", "teacher")
def admin_dashboard():
    with db() as con:
        close_expired_sessions(con)
        students_count = con.execute("SELECT COUNT(*) AS n FROM students").fetchone()["n"]
        teachers_count = con.execute("SELECT COUNT(*) AS n FROM teachers").fetchone()["n"]
        sessions_count = con.execute("SELECT COUNT(*) AS n FROM sessions").fetchone()["n"]
        active_sessions_count = con.execute("SELECT COUNT(*) AS n FROM sessions WHERE active = 1").fetchone()["n"]
        attendance_count = con.execute("SELECT COUNT(*) AS n FROM attendance").fetchone()["n"]

        recent = con.execute("""
            SELECT s.id, s.title, s.created_at, s.expires_at, s.active,
                   COUNT(a.id) AS total
            FROM sessions s LEFT JOIN attendance a ON a.session_id = s.id
            GROUP BY s.id ORDER BY s.id DESC LIMIT 25
        """).fetchall()

        # Trend Data: Attendance counts by date (last 14 active days)
        trend_rows = con.execute("""
            SELECT date(datetime(marked_at, 'unixepoch')) AS day, COUNT(*) AS count
            FROM attendance
            GROUP BY day ORDER BY day ASC LIMIT 14
        """).fetchall()

        today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        if trend_rows:
            chart_trend_labels = [r["day"] for r in trend_rows]
            chart_trend_data = [r["count"] for r in trend_rows]
        else:
            chart_trend_labels = [today_str]
            chart_trend_data = [0]

        # Lecture Bar Chart Data (last 8 sessions)
        session_chart_rows = list(reversed(recent[:8]))
        if session_chart_rows:
            chart_session_labels = [s["title"][:22] for s in session_chart_rows]
            chart_session_data = [s["total"] for s in session_chart_rows]
        else:
            chart_session_labels = ["No lectures yet"]
            chart_session_data = [0]

        closed_count = max(0, sessions_count - active_sessions_count)

        # Classroom Chat & Announcements
        messages = con.execute("""
            SELECT m.*, u.username
            FROM messages m
            JOIN users u ON u.id = m.user_id
            ORDER BY m.id DESC LIMIT 30
        """).fetchall()

    return render_template(
        "admin.html",
        students_count=students_count,
        teachers_count=teachers_count,
        sessions_count=sessions_count,
        active_sessions_count=active_sessions_count,
        attendance_count=attendance_count,
        recent=recent,
        chart_trend_labels=chart_trend_labels,
        chart_trend_data=chart_trend_data,
        chart_session_labels=chart_session_labels,
        chart_session_data=chart_session_data,
        active_count=active_sessions_count,
        closed_count=closed_count,
        messages=messages
    )


# --- Student Management ---
@app.route("/admin/students", methods=["GET", "POST"])
@role_required("admin", "teacher")
def manage_students():
    if request.method == "POST":
        try:
            create_account(
                username=request.form.get("username", ""),
                password=request.form.get("password", ""),
                role="student",
                name=request.form.get("name", ""),
                roll_no=request.form.get("roll_no", "")
            )
            flash("Student account created successfully.")
        except (ValueError, sqlite3.IntegrityError) as exc:
            flash(f"Could not create student: {exc}")
        return redirect(url_for("manage_students"))

    with db() as con:
        students = con.execute("""
            SELECT s.id, s.roll_no, s.name, u.username, u.id AS user_id,
                   COUNT(a.id) AS attendance_count
            FROM students s
            JOIN users u ON u.id = s.user_id
            LEFT JOIN attendance a ON a.student_id = s.id
            GROUP BY s.id
            ORDER BY s.roll_no ASC
        """).fetchall()
    return render_template("students.html", students=students)


@app.route("/admin/students/<int:student_id>/edit", methods=["GET", "POST"])
@role_required("admin", "teacher")
def edit_student(student_id):
    with db() as con:
        st = con.execute("""
            SELECT s.*, u.username, u.id AS user_id
            FROM students s JOIN users u ON u.id = s.user_id
            WHERE s.id = ?
        """, (student_id,)).fetchone()
        if not st:
            abort(404)

        if request.method == "POST":
            name = request.form.get("name", "").strip()
            roll_no = request.form.get("roll_no", "").strip()
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")

            if not name or not roll_no or not username:
                flash("Name, Roll Number, and Username are required.")
                return redirect(url_for("edit_student", student_id=student_id))

            try:
                con.execute("UPDATE students SET name = ?, roll_no = ? WHERE id = ?", (name, roll_no, student_id))
                if password:
                    if len(password) < 8:
                        flash("Password must be at least 8 characters.")
                        return redirect(url_for("edit_student", student_id=student_id))
                    con.execute(
                        "UPDATE users SET username = ?, password_hash = ? WHERE id = ?",
                        (username, generate_password_hash(password), st["user_id"])
                    )
                else:
                    con.execute("UPDATE users SET username = ? WHERE id = ?", (username, st["user_id"]))
                flash(f"Student '{name}' updated successfully.")
                return redirect(url_for("manage_students"))
            except sqlite3.IntegrityError as exc:
                flash(f"Update failed (duplicate roll number or username): {exc}")
                return redirect(url_for("edit_student", student_id=student_id))

    return render_template("edit_student.html", student=st)


@app.route("/admin/students/<int:student_id>/delete", methods=["POST"])
@role_required("admin", "teacher")
def delete_student(student_id):
    with db() as con:
        st = con.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()
        if not st:
            flash("Student not found.")
            return redirect(url_for("manage_students"))

        con.execute("DELETE FROM attendance WHERE student_id = ?", (student_id,))
        con.execute("DELETE FROM students WHERE id = ?", (student_id,))
        con.execute("DELETE FROM users WHERE id = ?", (st["user_id"],))

    flash(f"Student '{st['name']}' ({st['roll_no']}) deleted successfully.")
    return redirect(url_for("manage_students"))


# --- Teacher Management ---
@app.route("/admin/teachers", methods=["GET", "POST"])
@role_required("admin", "teacher")
def manage_teachers():
    if request.method == "POST":
        try:
            create_account(
                username=request.form.get("username", ""),
                password=request.form.get("password", ""),
                role="teacher",
                name=request.form.get("name", ""),
                department=request.form.get("department", "Computer Science"),
                email=request.form.get("email", "")
            )
            flash("Teacher account created successfully.")
        except (ValueError, sqlite3.IntegrityError) as exc:
            flash(f"Could not create teacher: {exc}")
        return redirect(url_for("manage_teachers"))

    with db() as con:
        teachers = con.execute("""
            SELECT t.id, t.name, t.department, t.email, u.username, u.role, u.id AS user_id
            FROM teachers t
            JOIN users u ON u.id = t.user_id
            ORDER BY t.name ASC
        """).fetchall()
    return render_template("teachers.html", teachers=teachers)


@app.route("/admin/teachers/<int:teacher_id>/edit", methods=["GET", "POST"])
@role_required("admin", "teacher")
def edit_teacher(teacher_id):
    with db() as con:
        t = con.execute("""
            SELECT t.*, u.username, u.id AS user_id
            FROM teachers t JOIN users u ON u.id = t.user_id
            WHERE t.id = ?
        """, (teacher_id,)).fetchone()
        if not t:
            abort(404)

        if request.method == "POST":
            name = request.form.get("name", "").strip()
            department = request.form.get("department", "").strip()
            email = request.form.get("email", "").strip()
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")

            if not name or not username:
                flash("Name and Username are required.")
                return redirect(url_for("edit_teacher", teacher_id=teacher_id))

            try:
                con.execute(
                    "UPDATE teachers SET name = ?, department = ?, email = ? WHERE id = ?",
                    (name, department, email, teacher_id)
                )
                if password:
                    if len(password) < 8:
                        flash("Password must be at least 8 characters.")
                        return redirect(url_for("edit_teacher", teacher_id=teacher_id))
                    con.execute(
                        "UPDATE users SET username = ?, password_hash = ? WHERE id = ?",
                        (username, generate_password_hash(password), t["user_id"])
                    )
                else:
                    con.execute("UPDATE users SET username = ? WHERE id = ?", (username, t["user_id"]))
                flash(f"Teacher '{name}' updated successfully.")
                return redirect(url_for("manage_teachers"))
            except sqlite3.IntegrityError as exc:
                flash(f"Update failed (duplicate username): {exc}")
                return redirect(url_for("edit_teacher", teacher_id=teacher_id))

    return render_template("edit_teacher.html", teacher=t)


@app.route("/admin/teachers/<int:teacher_id>/delete", methods=["POST"])
@role_required("admin", "teacher")
def delete_teacher(teacher_id):
    with db() as con:
        t = con.execute("SELECT * FROM teachers WHERE id = ?", (teacher_id,)).fetchone()
        if not t:
            flash("Teacher not found.")
            return redirect(url_for("manage_teachers"))

        if t["user_id"] == int(current_user.id):
            flash("You cannot delete your own account while logged in.")
            return redirect(url_for("manage_teachers"))

        con.execute("DELETE FROM teachers WHERE id = ?", (teacher_id,))
        con.execute("DELETE FROM users WHERE id = ?", (t["user_id"],))

    flash(f"Teacher '{t['name']}' deleted successfully.")
    return redirect(url_for("manage_teachers"))


# --- Lecture / Session Management ---
@app.route("/admin/session/new", methods=["POST"])
@role_required("admin", "teacher")
def new_session():
    title = request.form.get("title", "").strip()
    if not title:
        flash("Enter a session or lecture title.")
        return redirect(url_for("admin_dashboard"))

    try:
        duration_minutes = parse_duration(request.form)
    except ValueError as exc:
        flash(str(exc))
        return redirect(url_for("admin_dashboard"))

    now = int(datetime.now(timezone.utc).timestamp())
    token = make_dynamic_token()
    expires_at = now + QR_TIMEOUT
    ends_at = now + duration_minutes * 60

    with db() as con:
        close_expired_sessions(con, now)
        cur = con.execute("""
            INSERT INTO sessions(
                title, token, created_at, expires_at, active,
                duration_minutes, started_at, ends_at
            )
            VALUES (?, ?, ?, ?, 1, ?, ?, ?)
        """, (title, token, now, expires_at, duration_minutes, now, ends_at))
        session_id = cur.lastrowid
        con.execute("""
            INSERT OR IGNORE INTO dynamic_tokens(session_id, token, created_at, expires_at)
            VALUES (?, ?, ?, ?)
        """, (session_id, token, now, expires_at + QR_GRACE_PERIOD))

    return redirect(url_for("show_session", session_id=session_id))


@app.route("/admin/session/<int:session_id>")
@role_required("admin", "teacher")
def show_session(session_id):
    now = int(datetime.now(timezone.utc).timestamp())

    with db() as con:
        close_expired_sessions(con, now)
        s = con.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        if not s:
            abort(404)

        # If session is active and current QR slice expired, rotate token
        if s["active"] == 1 and s["expires_at"] <= now:
            token, expires_at = generate_session_token(con, session_id)
            s = con.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()

        count = con.execute(
            "SELECT COUNT(*) AS total FROM attendance WHERE session_id = ?", (session_id,)
        ).fetchone()["total"]

    qr_data = ""
    if s["active"] == 1:
        qr_data = generate_qr_base64(s["token"])

    remaining = max(0, s["expires_at"] - now) if s["active"] == 1 else 0

    return render_template(
        "session.html",
        s=s,
        qr_data=qr_data,
        timeout=QR_TIMEOUT,
        remaining=remaining,
        attendance_count=count
    )


@app.route("/admin/session/<int:session_id>/edit", methods=["GET", "POST"])
@role_required("admin", "teacher")
def edit_session(session_id):
    with db() as con:
        s = con.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        if not s:
            abort(404)

        if request.method == "POST":
            title = request.form.get("title", "").strip()
            active = 1 if request.form.get("active") in ("1", "on", "true") else 0
            if not title:
                flash("Lecture title cannot be empty.")
                return redirect(url_for("edit_session", session_id=session_id))

            con.execute("UPDATE sessions SET title = ?, active = ? WHERE id = ?", (title, active, session_id))
            flash(f"Lecture '{title}' updated successfully.")
            return redirect(url_for("admin_dashboard"))

    return render_template("edit_session.html", session=s)


@app.route("/admin/session/<int:session_id>/delete", methods=["POST"])
@role_required("admin", "teacher")
def delete_session(session_id):
    with db() as con:
        s = con.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        if not s:
            flash("Lecture session not found.")
            return redirect(url_for("admin_dashboard"))

        con.execute("DELETE FROM dynamic_tokens WHERE session_id = ?", (session_id,))
        archived_at = int(datetime.now(timezone.utc).timestamp())
        con.execute("""
            INSERT INTO archived_attendance(
                session_title, roll_no, student_name, marked_at, archived_at
            )
            SELECT s.title, st.roll_no, st.name, a.marked_at, ?
            FROM attendance a
            JOIN sessions s ON s.id = a.session_id
            JOIN students st ON st.id = a.student_id
            WHERE a.session_id = ?
        """, (archived_at, session_id))
        con.execute("DELETE FROM sessions WHERE id = ?", (session_id,))

    flash(f"Lecture '{s['title']}' deleted successfully.")
    return redirect(url_for("admin_dashboard"))


@app.route("/api/session/<int:session_id>/qr")
@role_required("admin", "teacher")
def session_qr_feed(session_id):
    now = int(datetime.now(timezone.utc).timestamp())
    with db() as con:
        close_expired_sessions(con, now)
        s = con.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        if not s:
            return jsonify(ok=False, message="Session not found."), 404

        if not s["active"]:
            return jsonify(ok=False, active=False, message="Session is closed.")

        if s["expires_at"] <= now:
            token, expires_at = generate_session_token(con, session_id)
        else:
            token = s["token"]
            expires_at = s["expires_at"]

        count = con.execute(
            "SELECT COUNT(*) AS total FROM attendance WHERE session_id = ?", (session_id,)
        ).fetchone()["total"]

    qr_b64 = generate_qr_base64(token)
    return jsonify(
        ok=True,
        active=True,
        qr_data=qr_b64,
        expires_at=expires_at,
        remaining=max(0, expires_at - now),
        timeout=QR_TIMEOUT,
        attendance_count=count
    )


@app.route("/admin/session/<int:session_id>/close", methods=["POST"])
@role_required("admin", "teacher")
def close_session(session_id):
    with db() as con:
        con.execute("UPDATE sessions SET active = 0 WHERE id = ?", (session_id,))
        con.execute("DELETE FROM dynamic_tokens WHERE session_id = ?", (session_id,))
    flash("Attendance session closed.")
    return redirect(url_for("admin_dashboard"))


# --- Classroom Chat & Discussion API ---
@app.route("/api/chat/post", methods=["POST"])
@login_required
def post_chat_message():
    message = request.form.get("message", "").strip()
    if not message:
        flash("Message cannot be empty.")
        return redirect(request.referrer or url_for("admin_dashboard"))

    now = int(datetime.now(timezone.utc).timestamp())
    author = current_user.display_name or current_user.username

    with db() as con:
        con.execute("""
            INSERT INTO messages(user_id, author_name, role, message, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (int(current_user.id), author, current_user.role, message, now))

    flash("Announcement posted.")
    return redirect(request.referrer or url_for("admin_dashboard"))


@app.route("/api/chat/<int:message_id>/delete", methods=["POST"])
@login_required
def delete_chat_message(message_id):
    with db() as con:
        msg = con.execute("SELECT * FROM messages WHERE id = ?", (message_id,)).fetchone()
        if not msg:
            flash("Message not found.")
            return redirect(request.referrer or url_for("admin_dashboard"))

        if current_user.role not in ("admin", "teacher") and msg["user_id"] != int(current_user.id):
            abort(403)

        con.execute("DELETE FROM messages WHERE id = ?", (message_id,))

    flash("Message deleted.")
    return redirect(request.referrer or url_for("admin_dashboard"))


# --- Student Area ---
@app.route("/student")
@role_required("student")
def student_dashboard():
    with db() as con:
        student = con.execute(
            "SELECT * FROM students WHERE user_id = ?",
            (int(current_user.id),)
        ).fetchone()

        records = con.execute("""
            SELECT s.title, a.marked_at
            FROM attendance a
            JOIN sessions s ON s.id = a.session_id
            WHERE a.student_id = ?
            ORDER BY a.marked_at DESC LIMIT 20
        """, (student["id"],)).fetchall()

        # Fetch recent classroom announcements
        messages = con.execute("""
            SELECT m.*, u.username
            FROM messages m
            JOIN users u ON u.id = m.user_id
            ORDER BY m.id DESC LIMIT 20
        """).fetchall()

    return render_template("student.html", student=student, records=records, messages=messages)


@app.route("/api/attendance", methods=["POST"])
@role_required("student")
def mark_attendance():
    data = request.get_json(silent=True) or {}
    token = data.get("token", "").strip()
    now = int(datetime.now(timezone.utc).timestamp())

    if not token:
        return jsonify(ok=False, message="No QR code data received."), 400

    with db() as con:
        close_expired_sessions(con, now)
        student = con.execute(
            "SELECT id FROM students WHERE user_id = ?",
            (int(current_user.id),)
        ).fetchone()

        if not student:
            return jsonify(ok=False, message="Student profile not found."), 403

        # Validate token against active sessions: either current session token or valid unexpired dynamic token
        s = con.execute("""
            SELECT s.id, s.title
            FROM sessions s
            WHERE s.active = 1
              AND s.ends_at > ?
              AND (
                  (s.token = ? AND s.expires_at + ? >= ?)
                  OR EXISTS (
                      SELECT 1 FROM dynamic_tokens dt
                      WHERE dt.session_id = s.id
                        AND dt.token = ?
                        AND dt.expires_at >= ?
                  )
              )
        """, (now, token, QR_GRACE_PERIOD, now, token, now)).fetchone()

        if not s:
            return jsonify(
                ok=False, message="QR code expired or invalid. Please scan the current code on screen."
            ), 400

        cur = con.execute("""
            INSERT OR IGNORE INTO attendance(session_id, student_id, marked_at)
            VALUES (?, ?, ?)
        """, (s["id"], student["id"], now))

        if cur.rowcount == 0:
            return jsonify(
                ok=True, message=f"Attendance already recorded for '{s['title']}'."
            )

    return jsonify(ok=True, message=f"Attendance marked successfully for '{s['title']}'!")


@app.route("/admin/report")
@role_required("admin", "teacher")
def report():
    archive_cutoff = int(datetime.now(timezone.utc).timestamp()) - ARCHIVED_REPORT_DAYS * 86400
    with db() as con:
        con.execute("DELETE FROM archived_attendance WHERE archived_at < ?", (archive_cutoff,))
        rows = con.execute("""
            SELECT s.title, st.roll_no, st.name, a.marked_at
            FROM attendance a
            JOIN sessions s ON s.id = a.session_id
            JOIN students st ON st.id = a.student_id
            UNION ALL
            SELECT session_title AS title, roll_no, student_name AS name, marked_at
            FROM archived_attendance
            ORDER BY marked_at DESC
        """).fetchall()
    return render_template("report.html", records=rows)


@app.route("/admin/report.csv")
@role_required("admin", "teacher")
def report_csv():
    archive_cutoff = int(datetime.now(timezone.utc).timestamp()) - ARCHIVED_REPORT_DAYS * 86400
    with db() as con:
        con.execute("DELETE FROM archived_attendance WHERE archived_at < ?", (archive_cutoff,))
        rows = con.execute("""
            SELECT s.title, st.roll_no, st.name, a.marked_at
            FROM attendance a
            JOIN sessions s ON s.id = a.session_id
            JOIN students st ON st.id = a.student_id
            UNION ALL
            SELECT session_title AS title, roll_no, student_name AS name, marked_at
            FROM archived_attendance
            ORDER BY marked_at DESC
        """).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Session", "Roll Number", "Student", "Timestamp (UTC)"])
    for r in rows:
        writer.writerow([
            r["title"], r["roll_no"], r["name"],
            datetime.fromtimestamp(r["marked_at"], timezone.utc).isoformat()
        ])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=attendance.csv"}
    )


@app.route("/admin/report.xlsx")
@role_required("admin", "teacher")
def report_excel():
    archive_cutoff = int(datetime.now(timezone.utc).timestamp()) - ARCHIVED_REPORT_DAYS * 86400
    with db() as con:
        con.execute("DELETE FROM archived_attendance WHERE archived_at < ?", (archive_cutoff,))
        rows = con.execute("""
            SELECT s.title, st.roll_no, st.name, a.marked_at
            FROM attendance a
            JOIN sessions s ON s.id = a.session_id
            JOIN students st ON st.id = a.student_id
            UNION ALL
            SELECT session_title AS title, roll_no, student_name AS name, marked_at
            FROM archived_attendance
            ORDER BY marked_at DESC
        """).fetchall()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Attendance"
    headers = ("Lecture", "Roll Number", "Student Name", "Attendance Date", "Attendance Time (IST)")
    worksheet.append(headers)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for row in rows:
        attended_at = datetime.fromtimestamp(row["marked_at"], IST)
        worksheet.append((
            row["title"],
            row["roll_no"],
            row["name"],
            attended_at.date(),
            attended_at.time().replace(tzinfo=None),
        ))

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column, width in {"A": 32, "B": 18, "C": 28, "D": 18, "E": 24}.items():
        worksheet.column_dimensions[column].width = width
    for cell in worksheet["D"][1:]:
        cell.number_format = "dd-mmm-yyyy"
    for cell in worksheet["E"][1:]:
        cell.number_format = "hh:mm:ss AM/PM"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    filename = f"attendance_report_{datetime.now(IST).strftime('%Y%m%d_%H%M%S_IST')}.xlsx"
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# if __name__ == "__main__":
#     init_db()
#     app.run(host='0.0.0.0',port=5000)
if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, ssl_context="adhoc")