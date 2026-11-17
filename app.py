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


# --------------------------------------------------
# SETTINGS
# --------------------------------------------------

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
FLASK_SECRET_KEY = os.environ.get("FLASK_SECRET_KEY")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")
FAMILY_ACCESS_PASSWORD = os.environ.get("FAMILY_ACCESS_PASSWORD")

required = {
    "SUPABASE_URL": SUPABASE_URL,
    "SUPABASE_SERVICE_ROLE_KEY": SUPABASE_SERVICE_ROLE_KEY,
    "FLASK_SECRET_KEY": FLASK_SECRET_KEY,
    "ADMIN_PASSWORD": ADMIN_PASSWORD,
    "FAMILY_ACCESS_PASSWORD": FAMILY_ACCESS_PASSWORD,
}
missing = [name for name, value in required.items() if not value]
if missing:
    raise RuntimeError("Missing environment variables: " + ", ".join(missing))

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


# --------------------------------------------------
# SECURITY AND HELPERS
# --------------------------------------------------

def csrf_token():
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_urlsafe(32)
    return session["_csrf"]


def verify_csrf():
    sent = request.form.get("_csrf", "")
    saved = session.get("_csrf", "")
    if not sent or not saved or not hmac.compare_digest(sent, saved):
        abort(400)


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


def render_page(template, **values):
    values.setdefault("token", csrf_token())
    inner = render_template_string(template, **values)
    return render_template_string(BASE, content=inner)


def excel_safe(value):
    value = "" if value is None else str(value)
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


# --------------------------------------------------
# DESIGN
BASE = r"""
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#07162b">
<title>ALGHADIR WALK | الغدیر واک</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">

<style>
:root {
    --navy: #07162b;
    --gold: #e8bd62;
    --gold-light: #f7df9c;
    --muted: #aebbd0;
    --teal: #28b7a5;
    --red: #d85b69;
}
* {
    box-sizing: border-box;
    font-family: "Vazirmatn", Tahoma, sans-serif;
}
body {
    margin: 0;
    min-height: 100vh;
    color: #fffaf0;
    background:
        radial-gradient(ellipse at 12% 8%, rgba(232,189,98,.18), transparent 32%),
        radial-gradient(ellipse at 90% 75%, rgba(40,183,165,.13), transparent 30%),
        linear-gradient(135deg, #061225, #0b1e37 55%, #102744);
}
a {
    color: var(--gold-light);
    text-decoration: none;
}
.topbar {
    width: min(1100px, 94%);
    margin: 24px auto 18px;
    padding: 19px 24px;
    border: 1px solid rgba(232,189,98,.35);
    border-radius: 22px;
    background: rgba(255,255,255,.06);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 15px;
}
.brand {
    display: flex;
    align-items: center;
    gap: 13px;
}
.brand-mark {
    width: 48px;
    height: 48px;
    border-radius: 16px;
    display: grid;
    place-items: center;
    color: var(--navy);
    font-size: 22px;
    font-weight: 900;
    background: linear-gradient(145deg, var(--gold-light), #c99438);
}
.brand-en {
    color: var(--gold-light);
    font-size: 17px;
    font-weight: 900;
    letter-spacing: 2px;
}
.brand-fa {
    color: white;
    font-size: 12px;
}
.top-tag {
    color: var(--gold-light);
    border: 1px solid rgba(232,189,98,.35);
    border-radius: 30px;
    padding: 8px 13px;
    font-size: 12px;
}
main {
    width: min(1000px, 94%);
    margin: 0 auto 45px;
}
.hero {
    text-align: center;
    padding: 24px 10px 19px;
}
.hero .eyebrow {
    color: var(--gold);
    letter-spacing: 3px;
    font-size: 11px;
    font-weight: 800;
}
.hero h1 {
    margin: 8px 0 5px;
    font-size: clamp(27px, 6vw, 43px);
    font-weight: 900;
}
.hero p {
    color: var(--muted);
    margin: 0;
    font-size: 14px;
}
.card {
    max-width: 760px;
    margin: 15px auto;
    padding: clamp(20px, 4vw, 34px);
    border: 1px solid rgba(232,189,98,.23);
    border-radius: 25px;
    background: rgba(255,255,255,.07);
    box-shadow: 0 20px 55px rgba(0,0,0,.25);
}
.card.wide {
    max-width: 100%;
}
.section-kicker {
    color: var(--gold);
    font-size: 11px;
    letter-spacing: 2px;
    font-weight: 800;
}
h2 {
    color: white;
    margin: 7px 0 18px;
    font-size: 23px;
}
h3 {
    color: var(--gold-light);
    margin: 25px 0 12px;
    font-size: 17px;
}
p {
    line-height: 1.9;
}
.muted {
    color: var(--muted);
    font-size: 13px;
    line-height: 1.9;
}
label {
    display: block;
    margin: 15px 0 7px;
    color: #f4e6c8;
    font-size: 13px;
    font-weight: 600;
}
input, textarea, select {
    width: 100%;
    min-height: 49px;
    padding: 11px 14px;
    border: 1px solid rgba(232,189,98,.3);
    border-radius: 13px;
    outline: none;
    color: white;
    background: rgba(3,13,27,.55);
    font-size: 14px;
}
input:focus, textarea:focus, select:focus {
    border-color: var(--gold);
    box-shadow: 0 0 0 3px rgba(232,189,98,.12);
}
input::placeholder {
    color: #91a0b6;
}
input[type="radio"] {
    width: 18px;
    min-height: auto;
    accent-color: var(--gold);
}
button, .btn {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    gap: 8px;
    min-height: 46px;
    border: 0;
    border-radius: 13px;
    padding: 11px 19px;
    color: #172033;
    font-size: 14px;
    font-weight: 800;
    cursor: pointer;
    background: linear-gradient(135deg, var(--gold-light), #d5a447);
}
button:hover, .btn:hover {
    filter: brightness(1.07);
}
.btn-teal {
    color: white;
    background: linear-gradient(135deg, #32c7b2, #168f85);
}
.btn-purple {
    color: white;
    background: linear-gradient(135deg, #9a88ec, #6754ba);
}
.btn-red {
    color: white;
    background: linear-gradient(135deg, #e77882, #bd4055);
}
.btn-outline {
    color: var(--gold-light);
    background: transparent;
    border: 1px solid rgba(232,189,98,.45);
}
.full {
    width: 100%;
    margin-top: 18px;
}
.actions {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 18px;
}
.choice-row {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
}
.choice {
    flex: 1;
    min-width: 130px;
    display: flex;
    align-items: center;
    gap: 9px;
    padding: 13px;
    border-radius: 13px;
    border: 1px solid rgba(232,189,98,.25);
    background: rgba(0,0,0,.13);
}
.member {
    padding: 15px;
    margin: 12px 0;
    border-radius: 16px;
    border: 1px solid rgba(232,189,98,.2);
    background: rgba(3,13,27,.25);
}
.member-grid {
    display: grid;
    grid-template-columns: 1fr 1fr auto;
    gap: 10px;
    align-items: end;
}
.member-grid label {
    margin-top: 0;
}
.remove-member {
    min-width: 48px;
    padding: 10px;
}
.alert {
    margin: 12px auto;
    padding: 12px 15px;
    border-radius: 13px;
    background: rgba(40,183,165,.14);
    border: 1px solid rgba(40,183,165,.4);
    color: #c8fff5;
    font-size: 13px;
}
.alert.error {
    background: rgba(216,91,105,.13);
    border-color: rgba(216,91,105,.4);
    color: #ffd7dc;
}
.stats {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
    margin: 18px 0;
}
.stat {
    padding: 18px 12px;
    border-radius: 17px;
    text-align: center;
    border: 1px solid rgba(232,189,98,.24);
    background: rgba(232,189,98,.08);
}
.stat strong {
    display: block;
    color: var(--gold-light);
    font-size: 27px;
}
.stat span {
    color: var(--muted);
    font-size: 12px;
}
.search-box {
    display: flex;
    flex-wrap: wrap;
    gap: 9px;
    margin: 20px 0;
}
.search-box form {
    display: flex;
    flex: 1;
    gap: 8px;
    min-width: 230px;
}
.search-box input {
    flex: 1;
    min-width: 100px;
}
.table-wrap {
    width: 100%;
    overflow-x: auto;
    border-radius: 15px;
    border: 1px solid rgba(232,189,98,.2);
}
table {
    width: 100%;
    border-collapse: collapse;
    min-width: 570px;
    background: rgba(3,13,27,.25);
}
th, td {
    padding: 12px 10px;
    text-align: right;
    border-bottom: 1px solid rgba(255,255,255,.09);
    font-size: 12px;
    vertical-align: top;
}
th {
    color: var(--gold-light);
    background: rgba(232,189,98,.1);
}
td {
    color: #e5eaf2;
}
.status-yes {
    color: #73e1c9;
    font-weight: 800;
}
.status-no {
    color: #ff9ba5;
    font-weight: 800;
}
.empty {
    text-align: center;
    color: var(--muted);
    padding: 25px;
}
.footer {
    text-align: center;
    padding: 25px 10px 18px;
    color: #8999b0;
    font-size: 11px;
}
.footer b {
    color: var(--gold);
    font-weight: 500;
}
.small-link {
    color: var(--gold-light);
    font-size: 13px;
}
.hidden {
    display: none !important;
}
@media (max-width: 600px) {
    .topbar {
        padding: 14px 15px;
    }
    .top-tag {
        display: none;
    }
    .brand-en {
        font-size: 14px;
    }
    .member-grid {
        grid-template-columns: 1fr 1fr;
    }
    .remove-member {
        grid-column: 1 / -1;
        width: 100%;
    }
    .stats {
        grid-template-columns: 1fr;
    }
    .search-box form {
        flex-direction: column;
    }
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
            <div class="alert {% if category == 'error' %}error{% endif %}">
                {{ message }}
            </div>
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


# --------------------------------------------------
# HOME PAGE
# --------------------------------------------------

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
        <input type="tel" name="phone" inputmode="numeric"
               placeholder="مثلاً 09123456789"
               pattern="09[0-9]{9}" maxlength="11" required
               value="{{ phone }}">
        <label>رمز مشترک خانواده</label>
        <input type="password" name="password"
               placeholder="رمز ورود خانواده" required>
        <button class="full" type="submit">ورود و ادامه</button>
    </form>
    <p class="muted">
        مدیر سامانه هستید؟
        <a class="small-link" href="{{ url_for('admin') }}">ورود مدیر</a>
    </p>
</div>
"""
# --------------------------------------------------
# FAMILY LOGIN
# --------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def home():
    if session.get("family_phone"):
        return redirect(url_for("register"))

    if request.method == "POST":
        verify_csrf()
        phone = request.form.get("phone", "").strip()
        password = request.form.get("password", "")

        if not re.fullmatch(r"09\d{9}", phone):
            flash("شماره تلفن باید ۱۱ رقم و با 09 شروع شود.", "error")
            return render_page(HOME_TEMPLATE, phone=phone)

        if not hmac.compare_digest(password, FAMILY_ACCESS_PASSWORD):
            flash("رمز ورود خانواده اشتباه است.", "error")
            return render_page(HOME_TEMPLATE, phone=phone)

        session.clear()
        session.permanent = True
        session["family_phone"] = phone
        csrf_token()
        return redirect(url_for("register"))

    return render_page(HOME_TEMPLATE, phone="")


# --------------------------------------------------
# FAMILY REGISTRATION PAGE
# --------------------------------------------------

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

        <div id="membersArea" {% if not attending %}class="hidden"{% endif %}>
            <h3>اعضای خانواده</h3>
            <p class="muted">
                نام و نام خانوادگی همه افراد شرکت‌کننده را وارد کنید.
            </p>

            <div id="members">
                {% for member in members %}
                <div class="member">
                    <div class="member-grid">
                        <div>
                            <label>نام</label>
                            <input name="first_name" placeholder="نام"
                                   value="{{ member.first_name }}" maxlength="50">
                        </div>
                        <div>
                            <label>نام خانوادگی</label>
                            <input name="last_name" placeholder="نام خانوادگی"
                                   value="{{ member.last_name }}" maxlength="60">
                        </div>
                        <button type="button" class="btn-red remove-member"
                                onclick="removeMember(this)">حذف</button>
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
                        <button type="button" class="btn-red remove-member"
                                onclick="removeMember(this)">حذف</button>
                    </div>
                </div>
                {% endif %}
            </div>

            <button type="button" class="btn-teal" onclick="addMember()">
                + افزودن عضو خانواده
            </button>
        </div>

        <button class="full" type="submit">ذخیره اطلاعات</button>
    </form>

    <form method="post" action="{{ url_for('family_logout') }}">
        <input type="hidden" name="_csrf" value="{{ token }}">
        <button type="submit" class="btn-outline" style="margin-top:15px">
            خروج از حساب خانواده
        </button>
    </form>
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
            <button type="button" class="btn-red remove-member"
                    onclick="removeMember(this)">حذف</button>
        </div>
    `;

    container.appendChild(member);
}

function removeMember(button) {
    const container = document.getElementById("members");
    const cards = container.querySelectorAll(".member");

    if (cards.length > 1) {
        button.closest(".member").remove();
    } else {
        cards[0].querySelectorAll("input").forEach(input => {
            input.value = "";
        });
    }
}
</script>
"""


@app.route("/register", methods=["GET", "POST"])
@family_required
def register():
    phone = session["family_phone"]

    try:
        result = (
            db.table("families")
            .select("*")
            .eq("phone", phone)
            .limit(1)
            .execute()
        )

        family = result.data[0] if result.data else None
        members = []

        if family:
            result = (
                db.table("family_members")
                .select("*")
                .eq("family_id", family["id"])
                .execute()
            )
            members = result.data or []

        if request.method == "POST":
            verify_csrf()

            attending_value = request.form.get("attending", "")
            if attending_value not in ("yes", "no"):
                flash("لطفاً وضعیت حضور را انتخاب کنید.", "error")
                return redirect(url_for("register"))

            attending = attending_value == "yes"
            new_members = []

            if attending:
                first_names = request.form.getlist("first_name")
                last_names = request.form.getlist("last_name")

                if len(first_names) > 30:
                    flash("حداکثر ۳۰ عضو می‌توانید ثبت کنید.", "error")
                    return redirect(url_for("register"))

                for first, last in zip(first_names, last_names):
                    first = first.strip()
                    last = last.strip()

                    if not first and not last:
                        continue

                    if not first or not last:
                        flash("نام و نام خانوادگی هر عضو را کامل کنید.", "error")
                        return redirect(url_for("register"))

                    new_members.append({
                        "first_name": first[:50],
                        "last_name": last[:60],
                    })

                if not new_members:
                    flash("حداقل یک عضو خانواده را وارد کنید.", "error")
                    return redirect(url_for("register"))
                                if family:
                family_id = family["id"]

                db.table("families").update({
                    "attending": attending
                }).eq("id", family_id).execute()

            else:
                created = db.table("families").insert({
                    "phone": phone,
                    "attending": attending
                }).execute().data

                if not created:
                    raise RuntimeError("Family record was not created")

                family_id = created[0]["id"]

            db.table("family_members").delete().eq(
                "family_id", family_id
            ).execute()

            if new_members:
                rows = [
                    {
                        "family_id": family_id,
                        "first_name": member["first_name"],
                        "last_name": member["last_name"]
                    }
                    for member in new_members
                ]

                db.table("family_members").insert(rows).execute()

            flash("اطلاعات خانواده با موفقیت ذخیره شد.", "success")
            return redirect(url_for("register"))

        return render_page(
            REGISTER_TEMPLATE,
            phone=phone,
            attending=bool(family and family.get("attending")),
            members=members
        )

    except Exception:
        app.logger.exception("Family registration error")
        flash("دریافت یا ذخیره اطلاعات انجام نشد.", "error")

        return render_page(
            REGISTER_TEMPLATE,
            phone=phone,
            attending=False,
            members=[]
        )


# --------------------------------------------------
# FAMILY LOGOUT
# --------------------------------------------------

@app.route("/family-logout", methods=["POST"])
@family_required
def family_logout():
    verify_csrf()
    session.clear()
    return redirect(url_for("home"))


# --------------------------------------------------
# ADMIN LOGIN
# --------------------------------------------------

ADMIN_LOGIN_TEMPLATE = r"""
<div class="card">
    <div class="section-kicker">ADMINISTRATION</div>
    <h2>ورود مدیر سامانه</h2>

    <form method="post">
        <input type="hidden" name="_csrf" value="{{ token }}">

        <label>رمز مدیریت</label>
        <input type="password" name="password"
               placeholder="رمز مدیر" required>

        <button class="full" type="submit">
            ورود به پنل مدیریت
        </button>
    </form>

    <p class="muted">
        <a href="{{ url_for('home') }}">بازگشت به صفحه خانواده‌ها</a>
    </p>
</div>
"""


# --------------------------------------------------
# ADMIN DASHBOARD DESIGN
# --------------------------------------------------

ADMIN_DASHBOARD_TEMPLATE = r"""
<div class="card wide">
    <div class="section-kicker">ADMIN DASHBOARD</div>
    <h2>پنل مدیریت الغدیر واک</h2>

    <div class="stats">
        <div class="stat">
            <strong>{{ total }}</strong>
            <span>کل خانواده‌های ثبت‌شده</span>
        </div>

        <div class="stat">
            <strong>{{ yes }}</strong>
            <span>خانواده‌های شرکت‌کننده</span>
        </div>

        <div class="stat">
            <strong>{{ total_people }}</strong>
            <span>تعداد افراد شرکت‌کننده</span>
        </div>
    </div>

    <div class="stat">
        <strong>{{ no }}</strong>
        <span>خانواده‌های غایب</span>
    </div>

    <div class="actions">
        <a class="btn btn-teal" href="{{ url_for('export') }}">
            دانلود گزارش Excel
        </a>

        <form method="post" action="{{ url_for('admin_logout') }}">
            <input type="hidden" name="_csrf" value="{{ token }}">
            <button type="submit" class="btn-outline">
                خروج از مدیریت
            </button>
        </form>
    </div>

    <h3>جست‌وجوی خانواده</h3>

    <div class="search-box">
        <form method="get" action="{{ url_for('admin') }}">
            <input name="q"
                   placeholder="شماره تلفن یا نام عضو خانواده"
                   value="{{ search }}">

            <button type="submit" class="btn-purple">
                جست‌وجو
            </button>
        </form>

        <a class="btn btn-outline" href="{{ url_for('admin') }}">
            نمایش همه
        </a>
    </div>
"""
    <h3>فهرست خانواده‌ها</h3>

    <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>شماره تلفن</th>
                    <th>وضعیت حضور</th>
                    <th>اعضای خانواده</th>
                    <th>عملیات</th>
                </tr>
            </thead>

            <tbody>
                {% for family in families %}
                <tr>
                    <td>{{ family.phone }}</td>

                    <td>
                        {% if family.attending %}
                            <span class="status-yes">شرکت می‌کند</span>
                        {% else %}
                            <span class="status-no">شرکت نمی‌کند</span>
                        {% endif %}
                    </td>

                    <td>
                        {% set family_members = members_by_family.get(family.id|string, []) %}
                        {% if family_members %}
                            {% for member in family_members %}
                                <div>
                                    {{ member.first_name }}
                                    {{ member.last_name }}
                                </div>
                            {% endfor %}
                        {% else %}
                            <span class="muted">عضوی ثبت نشده</span>
                        {% endif %}
                    </td>

                    <td>
                        <form method="post"
                              action="{{ url_for('delete_family', family_id=family.id) }}"
                              onsubmit="return confirm('آیا از حذف این خانواده مطمئن هستید؟')">

                            <input type="hidden" name="_csrf" value="{{ token }}">

                            <button type="submit" class="btn-red">
                                حذف خانواده
                            </button>
                        </form>
                    </td>
                </tr>

                {% else %}
                <tr>
                    <td colspan="4">هیچ خانواده‌ای پیدا نشد.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
"""


# --------------------------------------------------
# ADMIN DASHBOARD
# --------------------------------------------------

@app.route("/admin", methods=["GET", "POST"])
def admin():

    if not session.get("is_admin"):

        if request.method == "POST":
            verify_csrf()

            password = request.form.get("password", "")

            if hmac.compare_digest(password, ADMIN_PASSWORD):
                session.clear()
                session.permanent = True
                session["is_admin"] = True
                csrf_token()

                return redirect(url_for("admin"))

            flash("رمز مدیریت اشتباه است.", "error")

        return render_page(ADMIN_LOGIN_TEMPLATE)

    try:
        families_result = (
            db.table("families")
            .select("*")
            .order("id", desc=True)
            .execute()
        )

        members_result = (
            db.table("family_members")
            .select("*")
            .execute()
        )

        families = families_result.data or []
        all_members = members_result.data or []

        members_by_family = {}

        for member in all_members:
            family_id = str(member["family_id"])

            if family_id not in members_by_family:
                members_by_family[family_id] = []

            members_by_family[family_id].append(member)

        total = len(families)
        yes = sum(1 for family in families if family.get("attending"))
        no = total - yes

        total_people = sum(
            len(members_by_family.get(str(family["id"]), []))
            for family in families
            if family.get("attending")
        )

        search = request.args.get("q", "").strip()
        visible_families = families

        if search:
            query = search.casefold()
            visible_families = []

            for family in families:
                phone_match = query in str(family.get("phone", "")).casefold()

                member_match = any(
                    query in (
                        str(member.get("first_name", "")) + " " +
                        str(member.get("last_name", ""))
                    ).casefold()
                    for member in members_by_family.get(
                        str(family["id"]), []
                    )
                )

                if phone_match or member_match:
                    visible_families.append(family)

        return render_page(
            ADMIN_DASHBOARD_TEMPLATE,
            families=visible_families,
            members_by_family=members_by_family,
            total=total,
            yes=yes,
            no=no,
            total_people=total_people,
            search=search
        )

    except Exception:
        app.logger.exception("Admin dashboard error")
        flash("دریافت اطلاعات مدیریت انجام نشد. دوباره تلاش کنید.", "error")
        return render_page(ADMIN_LOGIN_TEMPLATE)
        # --------------------------------------------------
# DELETE FAMILY
# --------------------------------------------------

@app.route("/admin/delete/<family_id>", methods=["POST"])
@admin_required
def delete_family(family_id):
    verify_csrf()

    try:
        db.table("family_members").delete().eq(
            "family_id", family_id
        ).execute()

        db.table("families").delete().eq(
            "id", family_id
        ).execute()

        flash("خانواده با موفقیت حذف شد.", "success")

    except Exception:
        app.logger.exception("Delete family error")
        flash("حذف خانواده انجام نشد.", "error")

    return redirect(url_for("admin"))


# --------------------------------------------------
# EXCEL EXPORT
# --------------------------------------------------

@app.route("/admin/export")
@admin_required
def export():
    try:
        families = (
            db.table("families")
            .select("*")
            .order("id", desc=True)
            .execute()
            .data or []
        )

        members = (
            db.table("family_members")
            .select("*")
            .execute()
            .data or []
        )

        members_by_family = {}

        for member in members:
            family_id = str(member["family_id"])

            if family_id not in members_by_family:
                members_by_family[family_id] = []

            members_by_family[family_id].append(member)

        workbook = Workbook()

        sheet = workbook.active
        sheet.title = "Families"

        sheet.append([
            "شماره تلفن",
            "وضعیت حضور",
            "تعداد اعضا",
            "نام اعضای خانواده"
        ])

        for family in families:
            family_members = members_by_family.get(
                str(family["id"]), []
            )

            names = "، ".join(
                member["first_name"] + " " + member["last_name"]
                for member in family_members
            )

            sheet.append([
                excel_safe(family.get("phone")),
                "شرکت می‌کند" if family.get("attending") else "شرکت نمی‌کند",
                len(family_members),
                excel_safe(names)
            ])

        sheet.column_dimensions["A"].width = 20
        sheet.column_dimensions["B"].width = 22
        sheet.column_dimensions["C"].width = 15
        sheet.column_dimensions["D"].width = 55

        members_sheet = workbook.create_sheet("Members")

        members_sheet.append([
            "شماره تلفن خانواده",
            "نام",
            "نام خانوادگی"
        ])

        phone_by_id = {
            str(family["id"]): family.get("phone", "")
            for family in families
        }

        for member in members:
            members_sheet.append([
                excel_safe(
                    phone_by_id.get(str(member["family_id"]), "")
                ),
                excel_safe(member.get("first_name")),
                excel_safe(member.get("last_name"))
            ])

        members_sheet.column_dimensions["A"].width = 22
        members_sheet.column_dimensions["B"].width = 20
        members_sheet.column_dimensions["C"].width = 25

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="alghadir-family-walk.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception:
        app.logger.exception("Excel export error")
        flash("ساخت فایل Excel انجام نشد.", "error")
        return redirect(url_for("admin"))


# --------------------------------------------------
# ADMIN LOGOUT
# --------------------------------------------------

@app.route("/admin/logout", methods=["POST"])
@admin_required
def admin_logout():
    verify_csrf()
    session.clear()
    return redirect(url_for("admin"))


# --------------------------------------------------
# START APPLICATION
# --------------------------------------------------

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000))
    )
        
