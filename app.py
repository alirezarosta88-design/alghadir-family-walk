import os
import re
import io
import hmac
import secrets
from datetime import timedelta
from functools import wraps

from flask import (
    Flask, request, redirect, session, render_template_string,
    send_file, flash, abort, url_for
)
from supabase import create_client
from openpyxl import Workbook


# ---------------------------------------------------------
# 1) SETTINGS
# ---------------------------------------------------------

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
FLASK_SECRET_KEY = os.environ.get("FLASK_SECRET_KEY")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")
FAMILY_ACCESS_PASSWORD = os.environ.get("FAMILY_ACCESS_PASSWORD")

required_settings = {
    "SUPABASE_URL": SUPABASE_URL,
    "SUPABASE_SERVICE_ROLE_KEY": SUPABASE_SERVICE_ROLE_KEY,
    "FLASK_SECRET_KEY": FLASK_SECRET_KEY,
    "ADMIN_PASSWORD": ADMIN_PASSWORD,
    "FAMILY_ACCESS_PASSWORD": FAMILY_ACCESS_PASSWORD,
}

missing_settings = [key for key, value in required_settings.items() if not value]
if missing_settings:
    raise RuntimeError(
        "Missing Render environment variables: " + ", ".join(missing_settings)
    )

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
app.permanent_session_lifetime = timedelta(hours=12)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="Lax",
    MAX_CONTENT_LENGTH=2 * 1024 * 1024,
)

db = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)


# ---------------------------------------------------------
# 2) SECURITY HELPERS
# ---------------------------------------------------------

def csrf_token():
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_urlsafe(32)
    return session["_csrf_token"]


def check_csrf():
    sent = request.form.get("_csrf", "")
    saved = session.get("_csrf_token", "")
    return bool(sent and saved and hmac.compare_digest(sent, saved))


def require_csrf():
    if not check_csrf():
        abort(400, description="درخواست نامعتبر است. صفحه را تازه‌سازی کنید.")


def family_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if not session.get("family_phone"):
            return redirect(url_for("home"))
        return function(*args, **kwargs)
    return wrapper


def admin_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin"))
        return function(*args, **kwargs)
    return wrapper


def page(template, **context):
    context.setdefault("token", csrf_token())
    body = render_template_string(template, **context)
    return render_template_string(BASE, content=body)


def safe_excel(value):
    """Prevent spreadsheet formula injection in exported text."""
    value = "" if value is None else str(value)
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


# ---------------------------------------------------------
# 3) SHARED DESIGN
# ---------------------------------------------------------

BASE = r"""
<!doctype html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#08172d">
    <title>ALGHADIR WALK | الغدیر واک</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        :root {
            --navy: #07162b;
            --navy2: #102744;
            --gold: #e8bd62;
            --gold2: #f7df9c;
            --cream: #fffaf0;
            --muted: #aebbd0;
            --teal: #28b7a5;
            --red: #d85b69;
            --purple: #8b78dc;
        }
        * { box-sizing: border-box; font-family: "Vazirmatn", Tahoma, sans-serif; }
        body {
            margin: 0;
            min-height: 100vh;
            color: var(--cream);
            background:
                radial-gradient(ellipse at 12% 8%, rgba(232,189,98,.18), transparent 32%),
                radial-gradient(ellipse at 90% 75%, rgba(40,183,165,.13), transparent 30%),
                linear-gradient(135deg, #061225 0%, #0b1e37 52%, #102744 100%);
        }
        a { color: var(--gold2); text-decoration: none; }
        .topbar {
            width: min(1100px, 94%);
            margin: 24px auto 18px;
            padding: 19px 24px;
            border: 1px solid rgba(232,189,98,.35);
            border-radius: 22px;
            background: linear-gradient(120deg, rgba(255,255,255,.09), rgba(255,255,255,.025));
            box-shadow: 0 15px 45px rgba(0,0,0,.22);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 15px;
        }
        .brand { display: flex; align-items: center; gap: 13px; }
        .brand-mark {
            width: 48px; height: 48px; border-radius: 16px;
            display: grid; place-items: center;
            color: var(--navy); font-size: 25px; font-weight: 900;
            background: linear-gradient(145deg, var(--gold2), #c99438);
            box-shadow: 0 5px 22px rgba(232,189,98,.24);
        }
        .brand-en { color: var(--gold2); font-size: 17px; font-weight: 900; letter-spacing: 2px; }
        .brand-fa { color: white; font-size: 13px; opacity: .9; }
        .top-tag {
            color: var(--gold2); border: 1px solid rgba(232,189,98,.35);
            border-radius: 30px; padding: 8px 13px; font-size: 12px;
        }
        main { width: min(1000px, 94%); margin: 0 auto 45px; }
        .hero { text-align: center; padding: 24px 10px 19px; }
        .hero .eyebrow { color: var(--gold); letter-spacing: 3px; font-size: 11px; font-weight: 800; }
        .hero h1 { margin: 8px 0 5px; font-size: clamp(27px, 6vw, 43px); font-weight: 900; color: white; }
        .hero p { color: var(--muted); margin: 0; font-size: 14px; }
        .card {
            max-width: 760px; margin: 15px auto; padding: clamp(20px, 4vw, 34px);
            border: 1px solid rgba(232,189,98,.23); border-radius: 25px;
            background: linear-gradient(145deg, rgba(255,255,255,.105), rgba(255,255,255,.045));
            box-shadow: 0 20px 55px rgba(0,0,0,.25);
            backdrop-filter: blur(12px);
        }
        .card.wide { max-width: 100%; }
        .section-kicker { color: var(--gold); font-size: 11px; letter-spacing: 2px; font-weight: 800; }
        h2 { color: white; margin: 7px 0 18px; font-size: 23px; }
        h3 { color: var(--gold2); margin: 25px 0 12px; font-size: 17px; }
        p { line-height: 1.9; }
        .muted { color: var(--muted); font-size: 13px; line-height: 1.9; }
        label { display: block; margin: 15px 0 7px; color: #f4e6c8; font-size: 13px; font-weight: 600; }
        input, textarea, select {
            width: 100%; min-height: 49px; padding: 11px 14px;
            border: 1px solid rgba(232,189,98,.3); border-radius: 13px;
            outline: none; color: white; background: rgba(3,13,27,.55);
            font-size: 14px; transition: .2s;
        }
        input:focus, textarea:focus, select:focus {
            border-color: var(--gold); box-shadow: 0 0 0 3px rgba(232,189,98,.12);
        }
        input::placeholder { color: #91a0b6; }
        input[type="radio"] { width: 18px; min-height: auto; accent-color: var(--gold); }
        button, .btn {
            display: inline-flex; justify-content: center; align-items: center; gap: 8px;
            min-height: 46px; border: 0; border-radius: 13px; padding: 11px 19px;
            color: #172033; font-size: 14px; font-weight: 800; cursor: pointer;
            background: linear-gradient(135deg, var(--gold2), #d5a447);
            box-shadow: 0 7px 20px rgba(207,158,65,.17);
            transition: transform .15s, filter .15s;
        }
        button:hover, .btn:hover { transform: translateY(-2px); filter: brightness(1.06); }
        .btn-teal { color: white; background: linear-gradient(135deg, #32c7b2, #168f85); }
        .btn-purple { color: white; background: linear-gradient(135deg, #9a88ec, #6754ba); }
        .btn-red, button.no { color: white; background: linear-gradient(135deg, #e77882, #bd4055); }
        .btn-outline {
            color: var(--gold2); background: transparent;
            border: 1px solid rgba(232,189,98,.45); box-shadow: none;
        }
        .full { width: 100%; margin-top: 18px; }
        .actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 18px; }
        .choice-row { display: flex; flex-wrap: wrap; gap: 12px; }
        .choice {
            flex: 1; min-width: 130px; display: flex; align-items: center; gap: 9px;
            padding: 13px; border-radius: 13px; border: 1px solid rgba(232,189,98,.25);
            background: rgba(0,0,0,.13);
        }
        .member {
            padding: 15px; margin: 12px 0; border-radius: 16px;
            border: 1px solid rgba(232,189,98,.2); background: rgba(3,13,27,.25);
        }
        .member-grid { display: grid; grid-template-columns: 1fr 1fr auto; gap: 10px; align-items: end; }
        .member-grid label { margin-top: 0; }
        .remove-member { min-width: 48px; padding: 10px; }
        .alert {
            margin: 12px auto; padding: 12px 15px; border-radius: 13px;
            background: rgba(40,183,165,.14); border: 1px solid rgba(40,183,165,.4);
            color: #c8fff5; font-size: 13px;
        }
        .alert.error { background: rgba(216,91,105,.13); border-color: rgba(216,91,105,.4); color: #ffd7dc; }
        .stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin: 18px 0; }
        .stat {
            padding: 18px 12px; border-radius: 17px; text-align: center;
            border: 1px solid rgba(232,189,98,.24);
            background: linear-gradient(145deg, rgba(232,189,98,.13), rgba(255,255,255,.035));
        }
        .stat strong { display: block; color: var(--gold2); font-size: 27px; }
        .stat span { color: var(--muted); font-size: 12px; }
        .search-box { display: flex; flex-wrap: wrap; gap: 9px; margin: 20px 0; }
        .search-box form { display: flex; flex: 1; gap: 8px; min-width: 230px; }
        .search-box input { flex: 1; min-width: 100px; }
        .table-wrap { width: 100%; overflow-x: auto; border-radius: 15px; border: 1px solid rgba(232,189,98,.2); }
        table { width: 100%; border-collapse: collapse; min-width: 570px; background: rgba(3,13,27,.25); }
        th, td { padding: 12px 10px; text-align: right; border-bottom: 1px solid rgba(255,255,255,.09); font-size: 12px; vertical-align: top; }
        th { color: var(--gold2); background: rgba(232,189,98,.1); }
        td { color: #e5eaf2; }
        .status-yes { color: #73e1c9; font-weight: 800; }
        .status-no { color: #ff9ba5; font-weight: 800; }
        .empty { text-align: center; color: var(--muted); padding: 25px; }
        .footer { text-align: center; padding: 25px 10px 18px; color: #8999b0; font-size: 11px; }
        .footer b { color: var(--gold); font-weight: 500; }
        .small-link { color: var(--gold2); font-size: 13px; }
        .hidden { display: none !important; }
        @media (max-width: 600px) {
            .topbar { padding: 14px 15px; }
            .top-tag { display: none; }
            .brand-en { font-size: 14px; }
            .member-grid { grid-template-columns: 1fr 1fr; }
            .remove-member { grid-column: 1 / -1; width: 100%; }
            .stats { grid-template-columns: 1fr; }
            .search-box form { flex-direction: column; }
        }
    </style>
</head>
<body>
    <header class="topbar">
        <div class="brand">
            <div class="brand-mark">الغ</div>
            <div>
                <div class="brand-en">ALGHADIR WALK</div>
                <div class="brand-fa">سامانه ثبت‌نام پیاده‌روی خانوادگی الغدیر</div>
            </div>
        </div>
        <div class="top-tag">TOGETHER • WALK • JOY</div>
    </header>

    <main>
        <section class="hero">
            <div class="eyebrow">A DAY TO REMEMBER</div>
            <h1>الغدیر واک</h1>
            <p>با هم، برای یک روز زیبا و خاطره‌انگیز</p>
        </section>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% for category, message in messages %}
                <div class="alert {% if category == 'error' %}error{% endif %}">{{ message }}</div>
            {% endfor %}
        {% endwith %}

        {{ content|safe }}
    </main>

    <footer class="footer">
        سازنده: <b>علیرضا روستا</b>
    </footer>
</body>
</html>
"""


# ---------------------------------------------------------
# 4) HOME / FAMILY LOGIN
# ---------------------------------------------------------

HOME_TEMPLATE = r"""
<div class="card">
    <div class="section-kicker">FAMILY ACCESS</div>
    <h2>ورود خانواده</h2>
    <p class="muted">
        برای ثبت یا ویرایش پاسخ خانواده، شماره تلفن و رمز مشترک خانواده را وارد کنید.
    </p>

    <form method="post">
        <input type="hidden" name="_csrf" value="{{ token }}">

        <label>شماره تلفن همراه</label>
        <input
            type="tel" name="phone" inputmode="numeric"
            placeholder="مثلاً 09123456789"
            pattern="09[0-9]{9}" maxlength="11" required
            value="{{ phone }}"
        >

        <label>رمز مشترک خانواده</label>
        <input type="password" name="password" placeholder="رمز ورود خانواده" required>

        <button class="full" type="submit">ورود و ادامه</button>
    </form>

    <p class="muted" style="margin-top:18px">
        مدیر سامانه هستید؟
        <a class="small-link" href="{{ url_for('admin') }}">ورود مدیر</a>
    </p>
</div>
"""


@app.route("/", methods=["GET", "POST"])
def home():
    if session.get("family_phone"):
        return redirect(url_for("register"))

    if request.method == "POST":
        require_csrf()
        phone = request.form.get("phone", "").strip()
        password = request.form.get("password", "")

        if not re.fullmatch(r"09\d{9}", phone):
            flash("شماره تلفن باید ۱۱ رقم و با 09 شروع شود.", "error")
            return page(HOME_TEMPLATE, phone=phone)

        if not hmac.compare_digest(password, FAMILY_ACCESS_PASSWORD):
            flash("رمز ورود خانواده اشتباه است.", "error")
            return page(HOME_TEMPLATE, phone=phone)

        session.clear()
        session.permanent = True
        session["family_phone"] = phone
        csrf_token()
        return redirect(url_for("register"))

    return page(HOME_TEMPLATE, phone="")


# ---------------------------------------------------------
# 5) FAMILY REGISTRATION / EDIT
# ---------------------------------------------------------

REGISTER_TEMPLATE = r"""
<div class="card">
    <div class="section-kicker">YOUR FAMILY</div>
    <h2>ثبت‌نام یا ویرایش اطلاعات</h2>
    <p class="muted">شماره ثبت‌شده: <b>{{ phone }}</b></p>

    <form method="post" id="familyForm">
        <input type="hidden" name="_csrf" value="{{ token }}">

        <label>آیا خانواده شما در پیاده‌روی شرکت می‌کند؟</label>
        <div class="choice-row">
            <label class="choice">
                <input type="radio" name="attending" value="yes"
                    {% if attending %}checked{% endif %}>
                بله، شرکت می‌کنیم
            </label>
            <label class="choice">
                <input type="radio" name="attending" value="no"
                    {% if not attending %}checked{% endif %}>
                خیر، شرکت نمی‌کنیم
            </label>
        </div>

        <div id="membersArea" class="{% if not attending %}hidden{% endif %}">
            <h3>اعضای خانواده</h3>
            <p class="muted">نام و نام خانوادگی همه افراد شرکت‌کننده را وارد کنید.</p>

            <div id="members">
                {% for member in members %}
                <div class="member">
                    <div class="member-grid">
                        <div>
                            <label>نام</label>
                            <input name="first_name" placeholder="نام" value="{{ member.first_name }}" maxlength="50">
                        </div>
                        <div>
                            <label>نام خانوادگی</label>
                            <input name="last_name" placeholder="نام خانوادگی" value="{{ member.last_name }}" maxlength="60">
                        </div>
                        <button type="button" class="btn-red remove-member" onclick="removeMember(this)">حذف</button>
                    </div>
                </div>
                {% endfor %}
                {% if not members %}
                <div class="member">
                    <div class="member-grid">
                        <div>
                            <label>نام</label>
                            <input name="first_name" placeholder="نام" maxlength="50">
                        </div>
                        <div>
                            <label>نام خانوادگی</label>
                            <input name="last_name" placeholder="نام خانوادگی" maxlength="60">
                        </div>
                        <button type="button" class="btn-red remove-member" onclick="removeMember(this)">حذف</button>
                    </div>
                </div>
                {% endif %}
            </div>

            <button type="button" class="btn-teal" onclick="addMember()">+ افزودن عضو خانواده</button>
        </div>

        <button class="full" type="submit">ذخیره اطلاعات</button>
    </form>

    <div class="actions">
        <form method="post" action="{{ url_for('family_logout') }}">
            <input type="hidden" name="_csrf" value="{{ token }}">
            <button type="submit" class="btn-outline">خروج از حساب خانواده</button>
        </form>
    </div>
</div>

<script>
const membersArea = document.getElementById("membersArea");
document.querySelectorAll('input[name="attending"]').forEach((radio) => {
    radio.addEventListener("change", () => {
        membersArea.classList.toggle("hidden", radio.value !== "yes");
    });
});

function addMember() {
    const container = document.getElementById("members");
    const member = document.createElement("div");
    member.className = "member";
    member.innerHTML = `
        <div class="member-grid">
            <div>
                <label>نام</label>
                <input name="first_name" placeholder="نام" maxlength="50">
            </div>
            <div>
                <label>نام خانوادگی</label>
                <input name="last_name" placeholder="نام خانوادگی" maxlength="60">
            </div>
            <button type="button" class="btn-red remove-member" onclick="removeMember(this)">حذف</button>
        </div>`;
    container.appendChild(member);
}

function removeMember(button) {
    const container = document.getElementById("members");
    const cards = container.querySelectorAll(".member");
    if (cards.length > 1) {
        button.closest(".member").remove();
    } else {
        cards[0].querySelectorAll("input").forEach(input => input.value = "");
    }
}
</script>
"""


@app.route("/register", methods=["GET", "POST"])
@family_required
def register():
    phone = session["family_phone"]

    try:
        existing = (
            db.table("families")
            .select("*")
            .eq("phone", phone)
            .limit(1)
            .execute()
            .data
        )
        family = existing[0] if existing else None
        members = []

        if family:
            members = (
                db.table("family_members")
                .select("*")
                .eq("family_id", family["id"])
                .execute()
                .data
            )

        if request.method == "POST":
            requir
