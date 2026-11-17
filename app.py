import os
import re
import io
import hmac
import secrets
from functools import wraps

from flask import (
    Flask, request, redirect, session,
    render_template_string, send_file
)
from supabase import create_client
from openpyxl import Workbook


# -------------------- تنظیمات اصلی --------------------

app = Flask(__name__)
app.secret_key = os.environ["FLASK_SECRET_KEY"]

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

db = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SERVICE_ROLE_KEY"]
)

FAMILY_PASSWORD = os.environ["FAMILY_ACCESS_PASSWORD"]
ADMIN_PASSWORD = os.environ["ADMIN_PASSWORD"]

MAX_MEMBERS = 20


# -------------------- طراحی و رنگ‌بندی --------------------

BASE = """
<!doctype html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#101a35">
<title>ALGHADIR WALK | الغدیر واک</title>

<style>
:root {
    --navy: #101a35;
    --navy2: #1b2949;
    --gold: #d6ad55;
    --gold2: #f4d991;
    --cream: #fffaf0;
    --white: #ffffff;
    --muted: #778098;
    --red: #b94352;
    --teal: #258d88;
    --purple: #7860a8;
}

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    direction: rtl;
    font-family: Tahoma, Arial, sans-serif;
    color: var(--navy);
    background:
        radial-gradient(circle at 10% 10%, #e9dfc8 0, transparent 30%),
        radial-gradient(circle at 90% 80%, #d9e4ed 0, transparent 30%),
        #f5f3ee;
    min-height: 100vh;
}

.hero {
    position: relative;
    overflow: hidden;
    padding: 35px 15px 25px;
    text-align: center;
    color: white;
    background:
        radial-gradient(circle at 15% 20%, #33456c 0, transparent 35%),
        radial-gradient(circle at 85% 80%, #4a3b61 0, transparent 35%),
        linear-gradient(135deg, #0d1730, #202d4c);
    border-bottom: 3px solid var(--gold);
}

.hero:before,
.hero:after {
    content: "";
    position: absolute;
    border: 1px solid #d6ad5530;
    border-radius: 50%;
    width: 220px;
    height: 220px;
    pointer-events: none;
}

.hero:before {
    top: -130px;
    right: -80px;
}

.hero:after {
    bottom: -170px;
    left: -80px;
}

.hero-inner {
    position: relative;
    z-index: 1;
    max-width: 850px;
    margin: auto;
}

.eyebrow {
    display: inline-block;
    padding: 7px 16px;
    border: 1px solid #d6ad5580;
    border-radius: 30px;
    color: var(--gold2);
    font-size: 12px;
    margin-bottom: 12px;
    background: #ffffff0a;
}

.brand {
    font-family: Arial, sans-serif;
    font-size: clamp(27px, 6vw, 43px);
    font-weight: 900;
    letter-spacing: 3px;
    margin: 5px 0;
}

.brand span {
    color: var(--gold2);
}

.fa-brand {
    color: var(--gold2);
    font-size: 25px;
    font-weight: bold;
    margin: 8px 0;
}

.tagline {
    color: #d5dced;
    font-size: 13px;
    margin: 8px 0 20px;
}

.nav-menu {
    display: flex;
    justify-content: center;
    flex-wrap: wrap;
    gap: 9px;
}

.nav-menu a {
    text-decoration: none;
    padding: 10px 17px;
    border-radius: 12px;
    color: white;
    font-size: 13px;
    transition: .2s;
}

.nav-menu a:hover {
    transform: translateY(-2px);
    filter: brightness(1.15);
}

.nav-home {
    background: #52617f;
}

.nav-register {
    background: #258d88;
}

.nav-admin {
    background: #7860a8;
}

main {
    max-width: 850px;
    margin: 28px auto;
    padding: 0 14px;
    min-height: 350px;
}

.card {
    background: rgba(255,255,255,.96);
    border: 1px solid #e8dfca;
    border-radius: 20px;
    padding: 25px;
    margin-bottom: 20px;
    box-shadow: 0 12px 35px #101a3510;
}

.card h2 {
    text-align: center;
    color: var(--navy);
    margin-top: 5px;
}

.section-kicker {
    text-align: center;
    color: #a88743;
    font-size: 11px;
    font-weight: bold;
    letter-spacing: 1px;
}

.description {
    color: #687187;
    line-height: 2;
    text-align: center;
    font-size: 14px;
}

input {
    width: 100%;
    padding: 14px;
    margin: 8px 0;
    border: 1px solid #dce0e8;
    border-radius: 11px;
    background: #fafbfe;
    font-size: 15px;
    font-family: Tahoma, Arial, sans-serif;
    outline: none;
}

input:focus {
    border-color: var(--gold);
    box-shadow: 0 0 0 3px #d6ad5520;
}

button, .btn {
    display: block;
    width: 100%;
    padding: 13px;
    border: 0;
    border-radius: 11px;
    background: linear-gradient(135deg, #b68a36, #e5c879);
    color: #17213a;
    font-size: 14px;
    font-weight: bold;
    font-family: Tahoma, Arial, sans-serif;
    text-align: center;
    text-decoration: none;
    margin-top: 12px;
    cursor: pointer;
    transition: .2s;
}

button:hover, .btn:hover {
    transform: translateY(-2px);
    filter: brightness(1.04);
}

.btn-teal {
    background: var(--teal);
    color: white;
}

.btn-purple {
    background: var(--purple);
    color: white;
}

.no, .btn-red {
    background: var(--red);
    color: white;
}

.member {
    background: #f7f8fb;
    border: 1px solid #e5e8ef;
    padding: 14px;
    border-radius: 13px;
    margin: 12px 0;
}

.member-title {
    color: #9a7a3c;
    font-size: 12px;
    font-weight: bold;
}

.stat {
    background: linear-gradient(135deg, #f7f1e3, #fffaf0);
    border: 1px solid #eadbb8;
    padding: 16px;
    border-radius: 13px;
    margin: 9px 0;
    text-align: center;
}

.stats {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
    margin: 18px 0;
}

.stat strong {
    display: block;
    font-size: 25px;
    color: #9b7734;
    margin-bottom: 6px;
}

.stat span {
    font-size: 12px;
    color: #687187;
}

.search-box {
    background: #f7f8fb;
    border: 1px solid #e7e9ef;
    padding: 15px;
    border-radius: 15px;
    margin: 18px 0;
}

.table-wrap {
    overflow-x: auto;
    border-radius: 12px;
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    min-width: 560px;
    background: white;
}

td, th {
    padding: 12px 9px;
    border-bottom: 1px solid #ececf0;
    text-align: right;
    vertical-align: top;
}

th {
    background: var(--navy2);
    color: white;
}

table button {
    min-width: 60px;
    width: auto;
    padding: 8px 12px;
    margin: 0;
    font-size: 12px;
}

.empty {
    text-align: center;
    color: var(--muted);
    padding: 20px;
}

footer {
    text-align: center;
    padding: 22px 10px;
    color: #9299a8;
    font-size: 9px;
    letter-spacing: .2px;
}

footer span {
    color: #9a814f;
}

@media(max-width:600px) {
    .hero {
        padding-top: 27px;
    }

    .card {
        padding: 20px 16px;
        border-radius: 18px;
    }

    .stats {
        grid-template-columns: 1fr;
        gap: 7px;
    }

    .stat {
        padding: 13px 5px;
    }

    .stat strong {
        font-size: 21px;
    }

    .stat span {
        font-size: 10px;
    }

    .nav-menu a {
        padding: 9px 12px;
        font-size: 12px;
    }
}
</style>
</head>

<body>

<header class="hero">
    <div class="hero-inner">
        <div class="eyebrow">✦ یک روز زیبا در کنار خانواده ✦</div>

        <h1 class="brand">ALGHADIR <span>WALK</span></h1>

        <div class="fa-brand">الغدیر واک</div>

        <p class="tagline">
            سامانه ثبت‌نام پیاده‌روی خانوادگی
        </p>

        <nav class="nav-menu">
            <a class="nav-home" href="/">⌂ صفحه اصلی</a>
            <a class="nav-register" href="/register">✦ ثبت‌نام خانواده</a>
            <a class="nav-admin" href="/admin">♛ مدیریت</a>
        </nav>
    </div>
</header>

<main>
    {{ body|safe }}
</main>

<footer>
    <span>سازنده: علیرضا روستا</span>
</footer>

</body>
</html>
"""


# -------------------- توابع کمکی --------------------

def page(template, **data):
    body = render_template_string(template, **data)
    return render_template_string(BASE, body=body)


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


def check_csrf():
    token = request.form.get("csrf", "")
    saved = session.get("csrf", "")
    return bool(saved) and hmac.compare_digest(token, saved)


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin"):
            return redirect("/admin")
        return fn(*args, **kwargs)
    return wrapper


def safe_excel(value):
    value = str(value or "")
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


# -------------------- صفحه اصلی و ورود خانواده --------------------

@app.route("/", methods=["GET", "POST"])
def home():

    if request.method == "POST":

        if not check_csrf():
            return "درخواست نامعتبر است. صفحه را دوباره باز کنید.", 400

        phone = request.form.get("phone", "").strip()
        password = request.form.get("password", "")

        if not re.fullmatch(r"09[0-9]{9}", phone):
            return page("""
            <div class="card">
                <h2>شماره تلفن معتبر نیست</h2>
                <p class="description">
                    شماره را با فرمت 09xxxxxxxxx وارد کنید.
                </p>
                <a class="btn" href="/">بازگشت</a>
            </div>
            """)

        if not hmac.compare_digest(password, FAMILY_PASSWORD):
            return page("""
            <div class="card">
                <h2>رمز ورود اشتباه است</h2>
                <p class="description">
                    رمز را بررسی و دوباره تلاش کنید.
                </p>
                <a class="btn" href="/">تلاش دوباره</a>
            </div>
            """)

        session.clear()
        session["phone"] = phone
        session["csrf"] = secrets.token_urlsafe(32)

        return redirect("/register")

    return page("""
    <div class="card">
        <div class="section-kicker">WELCOME TO ALGHADIR WALK</div>

        <h2>به جمع خانواده‌های الغدیر خوش آمدید</h2>

        <p class="description">
            برای ثبت‌نام یا ویرایش اطلاعات خانواده،
            شماره تلفن و رمز ورود را وارد کنید.
        </p>

        <form method="post">
            <input type="hidden" name="csrf" value="{{ token }}">

            <input
                name="phone"
                placeholder="شماره تلفن همراه"
                inputmode="numeric"
                maxlength="11"
                required
            >

            <input
                type="password"
                name="password"
                placeholder="رمز عمومی ثبت‌نام"
                required
            >

            <button>ورود به سامانه ←</button>
        </form>
    </div>

    <div class="card">
        <div class="section-kicker">ALGHADIR FAMILY EVENT</div>
        <h2>کنار هم، برای یک خاطره زیبا</h2>
        <p class="description">
            حضور شما و خانواده عزیزتان باعث خوشحالی ماست.
        </p>
    </div>
    """, token=csrf_token())


# -------------------- ثبت‌نام و ویرایش خانواده --------------------

@app.route("/register", methods=["GET", "POST"])
def register():

    phone = session.get("phone")

    if not phone:
        return redirect("/")

    try:
        result = db.table("families").select("*").eq(
            "phone", phone
        ).execute()

        family = result.data[0] if result.data else None
        members = []

        if family:
            members = db.table("family_members").select("*").eq(
                "family_id", family["id"]
            ).execute().data

    except Exception:
        return page("""
        <div class="card">
            <h2>خطا در دریافت اطلاعات</h2>
            <p class="description">
                لطفاً دوباره تلاش کنید.
            </p>
            <a class="btn" href="/">بازگشت</a>
        </div>
        """)

    if request.method == "POST":

        if not check_csrf():
            return "درخواست نامعتبر است. صفحه را دوباره باز کنید.", 400

        answer = request.form.get("answer")

        if answer not in ("yes", "no"):
            return redirect("/register")

        new_members = []

        if answer == "yes":

            first_names = request.form.getlist("first_name")
            last_names = request.form.getlist("last_name")

            if not first_names or len(first_names) > MAX_MEMBERS:
                return "تعداد اعضای خانواده نامعتبر است.", 400

            if len(first_names) != len(last_names):
                return "اطلاعات اعضا ناقص است.", 400

            for first, last in zip(first_names, last_names):

                first = first.strip()
                last = last.strip()

                if (
                    not first or not last
                    or len(first) > 50
                    or len(last) > 50
                ):
                    return "نام و نام خانوادگی را درست وارد کنید.", 400

                new_members.append({
                    "first_name": first,
                    "last_name": last
                })

        try:

            if family:

                family_id = family["id"]

                db.table("families").update({
                    "attending": answer == "yes"
                }).eq("id", family_id).execute()

                db.table("family_members").delete().eq(
                    "family_id", family_id
                ).execute()

            else:

                result = db.table("families").insert({
                    "phone": phone,
                    "attending": answer == "yes"
                }).execute()

                family_id = result.data[0]["id"]

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

            session.pop("phone", None)

            return page("""
            <div class="card">
                <h2>اطلاعات شما با موفقیت ثبت شد</h2>
                <p class="description">
                    از همراهی شما سپاسگزاریم.
                </p>
                <a class="btn" href="/">بازگشت به صفحه اصلی</a>
            </div>
            """)

        except Exception:

            return page("""
            <div class="card">
                <h2>ثبت اطلاعات انجام نشد</h2>
                <p class="description">
                    لطفاً دوباره تلاش کنید.
                </p>
                <a class="btn" href="/register">تلاش مجدد</a>
            </div>
            """)

    if family and family["attending"]:
        initial_members = members
    else:
        initial_members = [{"first_name": "", "last_name": ""}]

    return page("""
    <div class="card">

        <div class="section-kicker">FAMILY REGISTRATION</div>

        <h2>{{ heading }}</h2>

        <p class="description">
            شماره تماس: {{ phone }}
        </p>

        <form method="post">
            <input type="hidden" name="csrf" value="{{ token }}">
            <input type="hidden" name="answer" value="yes">

            <h3>اعضای خانواده</h3>

            <div id="members">

                {% for m in initial_members %}
                <div class="member">
                    <div class="member-title">عضو خانواده</div>

                    <input
                        name="first_name"
                        placeholder="نام"
                        value="{{ m.first_name }}"
                        maxlength="50"
                        required
                    >

                    <input
                        name="last_name"
                        placeholder="نام خانوادگی"
                        value="{{ m.last_name }}"
                        maxlength="50"
                        required
                    >
                </div>
                {% endfor %}

            </div>

            <button
                type="button"
                class="btn-teal"
                onclick="addMember()"
            >
                + افزودن عضو خانواده
            </button>

            <button type="submit">
                بله، شرکت می‌کنیم / ذخیره تغییرات
            </button>
        </form>

        <form method="post">
            <input type="hidden" name="csrf" value="{{ token }}">
            <input type="hidden" name="answer" value="no">

            <button class="no" type="submit">
                خیر، امکان حضور نداریم
            </button>
        </form>

        <a class="btn btn-purple" href="/leave">
            خروج
        </a>

    </div>

    <script>
    function addMember() {

        const box = document.getElementById("members");

        if (box.children.length >= {{ max_members }}) {
            alert("حداکثر تعداد اعضا {{ max_members }} نفر است.");
            return;
        }

        const div = document.createElement("div");
        div.className = "member";

        const title = document.createElement("div");
        title.className = "member-title";
        title.textContent = "عضو خانواده";

        const first = document.createElement("input");
        first.name = "first_name";
        first.placeholder = "نام";
        first.maxLength = 50;
        first.required = true;

        const last = document.createElement("input");
        last.name = "last_name";
        last.placeholder = "نام خانوادگی";
        last.maxLength = 50;
        last.required = true;

        div.append(title, first, last);
        box.appendChild(div);
    }
    </script>
    """,
    heading="ویرایش اطلاعات خانواده" if family else "ثبت‌نام خانواده",
    phone=phone,
    initial_members=initial_members,
    token=csrf_token(),
    max_members=MAX_MEMBERS
    )


# -------------------- خروج خانواده --------------------

@app.route("/leave")
def leave():
    session.clear()
    return redirect("/")


# -------------------- پنل مدیریت --------------------

@app.route("/admin", methods=["GET", "POST"])
def admin():

    if request.method == "POST":

        if not check_csrf():
            return "درخواست نامعتبر است.", 400

        password = request.form.get("password", "")

        if hmac.compare_digest(password, ADMIN_PASSWORD):

            session.clear()
            session["admin"] = True
            session["csrf"] = secrets.token_urlsafe(32)

            return redirect("/admin")

        return page("""
        <div class="card">
            <h2>رمز مدیریت اشتباه است</h2>
            <a class="btn" href="/admin">تلاش دوباره</a>
        </div>
        """)

    if not session.get("admin"):

        return page("""
        <div class="card">
            <div class="section-kicker">ADMINISTRATION</div>
            <h2>ورود مدیریت</h2>

            <form method="post">
                <input type="hidden" name="csrf" value="{{ token }}">

                <input
                    type="password"
                    name="password"
                    placeholder="رمز مدیریت"
                    required
                >

                <button>ورود به پنل مدیریت</button>
            </form>
        </div>
        """, token=csrf_token())

    try:

        all_families = db.table("families").select("*").order(
            "created_at", desc=True
        ).execute().data

        all_members = db.table("family_members").select(
            "*"
        ).execute().data

        # آمار کلی، بدون تأثیر جست‌وجو
        total = len(all_families)
        yes = sum(1 for f in all_families if f["attending"])
        no = total - yes

        attending_ids = {
            f["id"] for f in all_families if f["attending"]
        }

        total_people = sum(
            1 for m in all_members
            if m["family_id"] in attending_ids
        )

        # آماده‌سازی اطلاعات برای نمایش
        members_by_family = {}

        for member in all_members:
            family_id = member["fam
