import os
import io
import re
import hmac
import secrets
import logging
from datetime import timedelta
from functools import wraps

from flask import (
    Flask, request, redirect, url_for, session,
    render_template_string, flash, abort, send_file
)
from supabase import create_client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# -------------------- SETTINGS --------------------

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")
FAMILY_ACCESS_PASSWORD = os.getenv("FAMILY_ACCESS_PASSWORD")

required_settings = {
    "SUPABASE_URL": SUPABASE_URL,
    "SUPABASE_SERVICE_ROLE_KEY": SUPABASE_SERVICE_ROLE_KEY,
    "FLASK_SECRET_KEY": FLASK_SECRET_KEY,
    "ADMIN_PASSWORD": ADMIN_PASSWORD,
    "FAMILY_ACCESS_PASSWORD": FAMILY_ACCESS_PASSWORD,
}

missing_settings = [
    key for key, value in required_settings.items() if not value
]

if missing_settings:
    raise RuntimeError(
        "Missing Render environment variables: "
        + ", ".join(missing_settings)
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

logging.basicConfig(level=logging.INFO)


# -------------------- SECURITY HELPERS --------------------

def get_csrf_token():
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_urlsafe(32)

    return session["_csrf_token"]


def check_csrf():
    sent = request.form.get("_csrf_token", "")
    saved = session.get("_csrf_token", "")

    if not sent or not saved or not hmac.compare_digest(sent, saved):
        abort(400)


def family_login_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if not session.get("family_phone"):
            return redirect(url_for("home"))

        return function(*args, **kwargs)

    return wrapper


def admin_login_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin"))

        return function(*args, **kwargs)

    return wrapper


def show_page(template, **values):
    values["csrf_token"] = get_csrf_token()

    inner_html = render_template_string(template, **values)

    return render_template_string(
        BASE_TEMPLATE,
        content=inner_html
    )


def safe_excel_text(value):
    value = "" if value is None else str(value)

    if value.startswith(("=", "+", "-", "@")):
        return "'" + value

    return value


# -------------------- MAIN DESIGN --------------------

BASE_TEMPLATE = r'''
<!doctype html>
<html lang="fa" dir="rtl">

<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>الغدیر واک | ثبت‌نام خانواده‌ها</title>

<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>

<link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<style>

:root {
    --navy: #14213d;
    --blue: #2563eb;
    --gold: #f5b942;
    --teal: #0f9d8a;
    --red: #dc4545;
    --ink: #1d2940;
    --muted: #718096;
}

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    font-family: Vazirmatn, Tahoma, sans-serif;
    color: var(--ink);
    background: linear-gradient(
        135deg,
        #eaf1ff,
        #fff8e9 55%,
        #e6faf5
    );
    min-height: 100vh;
}

header {
    background: linear-gradient(
        110deg,
        #101b35,
        #203d70
    );
    color: white;
    padding: 25px 16px 34px;
    text-align: center;
    border-bottom: 5px solid var(--gold);
}

.brand {
    font-size: clamp(25px, 6vw, 39px);
    font-weight: 800;
    letter-spacing: 2px;
    color: #ffd36b;
}

.brand-fa {
    font-size: 22px;
    font-weight: 700;
    margin-top: 5px;
}

.tagline {
    font-size: 13px;
    color: #dce7ff;
    margin-top: 8px;
}

main {
    width: min(940px, calc(100% - 24px));
    margin: 26px auto 45px;
}

.card {
    background: white;
    border: 1px solid #e5eaf3;
    border-radius: 22px;
    padding: 24px;
    box-shadow: 0 12px 35px #15284b12;
    margin-bottom: 18px;
}

.card h2 {
    margin: 4px 0 12px;
    color: var(--navy);
    font-size: 23px;
}

.card h3 {
    color: var(--navy);
    margin-top: 24px;
}

.kicker {
    color: #bd831b;
    font-weight: 800;
    font-size: 12px;
    letter-spacing: 1.5px;
}

.muted {
    color: var(--muted);
    font-size: 14px;
    line-height: 1.9;
}

label {
    display: block;
    font-weight: 600;
    margin: 12px 0 7px;
}

input {
    font: inherit;
    width: 100%;
    border: 1px solid #d8dfeb;
    border-radius: 12px;
    padding: 12px 13px;
    background: #fbfcff;
    outline: none;
}

input:focus {
    border-color: var(--blue);
    box-shadow: 0 0 0 3px #2563eb18;
}

button, .btn {
    font: inherit;
    border: 0;
    border-radius: 12px;
    padding: 11px 17px;
    cursor: pointer;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    transition: .15s;
}

button:hover, .btn:hover {
    filter: brightness(.96);
    transform: translateY(-1px);
}

.primary {
    background: linear-gradient(110deg, #2458d7, #3678f6);
    color: white;
}

.gold {
    background: linear-gradient(110deg, #e9a928, #ffd36b);
    color: #392600;
}

.teal {
    background: #e0f7f1;
    color: #087c6d;
    border: 1px solid #bceadd;
}

.outline {
    background: white;
    border: 1px solid #cbd5e1;
    color: #34445f;
}

.danger {
    background: #fff0f0;
    color: #bd3030;
    border: 1px solid #ffd0d0;
}

.full {
    width: 100%;
    margin-top: 18px;
}

.choice-row {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin: 14px 0;
}

.choice {
    display: flex;
    align-items: center;
    gap: 8px;
    background: #f4f7fc;
    border: 1px solid #e0e7f2;
    border-radius: 12px;
    padding: 12px 14px;
}

.choice input {
    width: auto;
}

.member {
    border: 1px solid #e4eaf3;
    border-radius: 15px;
    padding: 13px;
    margin: 12px 0;
    background: #fbfcff;
}

.member-grid {
    display: grid;
    grid-template-columns: 1fr 1fr auto;
    gap: 10px;
    align-items: end;
}

.flash {
    padding: 12px 15px;
    border-radius: 12px;
    margin-bottom: 14px;
    font-size: 14px;
}

.flash.success {
    background: #e5f8ef;
    color: #147345;
}

.flash.error {
    background: #fff0f0;
    color: #b83232;
}

.hidden {
    display: none !important;
}

.actions {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    margin: 18px 0;
}

.stats {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin: 20px 0;
}

.stat {
    background: linear-gradient(145deg, #f0f5ff, #fff9e9);
    border: 1px solid #e3eaf6;
    border-radius: 16px;
    padding: 16px;
    text-align: center;
}

.stat strong {
    display: block;
    font-size: 28px;
    color: #1d4c9b;
}

.stat span {
    font-size: 12px;
    color: #65738a;
}

.search-form {
    display: flex;
    gap: 9px;
    margin: 12px 0;
}

.search-form input {
    flex: 1;
}

.table-wrap {
    overflow-x: auto;
    margin-top: 14px;
}

table {
    width: 100%;
    border-collapse: collapse;
    min-width: 650px;
}

th, td {
    padding: 12px 10px;
    text-align: right;
    border-bottom: 1px solid #e8edf4;
    vertical-align: top;
    font-size: 13px;
}

th {
    background: #f0f4fb;
    color: #263b60;
}

.status-yes {
    color: #087c57;
    font-weight: 700;
}

.status-no {
    color: #bd4545;
    font-weight: 700;
}

footer {
    text-align: center;
    color: #7b8798;
    font-size: 12px;
    padding: 0 12px 24px;
}

a {
    color: #2458d7;
}

@media(max-width:650px) {

    .card {
        padding: 18px;
    }

    .member-grid {
        grid-template-columns: 1fr;
    }

    .member-grid button {
        width: 100%;
    }

    .stats {
        grid-template-columns: 1fr 1fr;
    }

    .search-form {
        flex-direction: column;
    }

    .brand {
        font-size: 27px;
    }

}

</style>
</head>

<body>

<header>

    <div class="brand">ALGHADIR WALK</div>

    <div class="brand-fa">الغدیر واک</div>

    <div class="tagline">
        سامانه ثبت‌نام پیاده‌روی خانوادگی
    </div>

</header>

<main>

    {% with messages = get_flashed_messages(with_categories=true) %}

        {% for category, message in messages %}

            <div class="flash {{ category }}">
                {{ message }}
            </div>

        {% endfor %}

    {% endwith %}

    {{ content|safe }}

</main>

<footer>
    سامانه الغدیر واک · سازنده: علیرضا روستا
</footer>

</body>
</html>
'''


# -------------------- FAMILY LOGIN PAGE --------------------

HOME_TEMPLATE = r'''
<div class="card">

    <div class="kicker">FAMILY REGISTRATION</div>

    <h2>ورود خانواده</h2>

    <p class="muted">
        شماره موبایل و رمز مشترک اعلام‌شده برای خانواده‌ها را وارد کنید.
    </p>

    <form method="post">

        <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">

        <label>شماره موبایل</label>

        <input
            name="phone"
            inputmode="numeric"
            autocomplete="tel"
            maxlength="11"
            placeholder="مثلاً 09123456789"
            value="{{ phone }}"
            required
        >

        <label>رمز ورود خانواده</label>

        <input
            type="password"
            name="password"
            autocomplete="current-password"
            placeholder="رمز ورود"
            required
        >

        <button class="primary full" type="submit">
            ورود و ثبت اطلاعات
        </button>

    </form>

    <p class="muted" style="text-align:center;margin-top:20px">
        <a href="{{ url_for('admin') }}">ورود مدیر سامانه</a>
    </p>

</div>
'''


# -------------------- FAMILY REGISTRATION PAGE --------------------

REGISTER_TEMPLATE = r'''
<div class="card">

    <div class="kicker">YOUR FAMILY</div>

    <h2>ثبت‌نام یا ویرایش اطلاعات</h2>

    <p class="muted">
        شماره موبایل: <b>{{ phone }}</b>
    </p>

    <form method="post">

        <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">

        <label>آیا خانواده شما در پیاده‌روی شرکت می‌کند؟</label>

        <div class="choice-row">

            <label class="choice">
                <input
                    type="radio"
                    name="attending"
                    value="yes"
                    {% if attending %}checked{% endif %}
                >
                بله، شرکت می‌کنیم
            </label>

            <label class="choice">
                <input
                    type="radio"
                    name="attending"
                    value="no"
                    {% if not attending %}checked{% endif %}
                >
                خیر، شرکت نمی‌کنیم
            </label>

        </div>

        <section id="members-area" {% if not attending %}class="hidden"{% endif %}>

            <h3>اعضای شرکت‌کننده</h3>

            <p class="muted">
                نام و نام خانوادگی هر فرد را جداگانه وارد کنید.
            </p>

            <div id="members-list">

                {% for member in members %}

                <div class="member">

                    <div class="member-grid">

                        <div>
                            <label>نام</label>
                            <input
                                name="first_name"
                                maxlength="50"
                                value="{{ member.first_name }}"
                                placeholder="نام"
                            >
                        </div>

                        <div>
                            <label>نام خانوادگی</label>
                            <input
                                name="last_name"
                                maxlength="60"
                                value="{{ member.last_name }}"
                                placeholder="نام خانوادگی"
                            >
                        </div>

                        <button
                            type="button"
                            class="danger"
                            onclick="removeMember(this)"
                        >
                            حذف
                        </button>

                    </div>

                </div>

                {% endfor %}

                {% if not members %}

                <div class="member">

                    <div class="member-grid">

                        <div>
                            <label>نام</label>
                            <input
                                name="first_name"
                                maxlength="50"
                                placeholder="نام"
                            >
                        </div>

                        <div>
                            <label>نام خانوادگی</label>
                            <input
                                name="last_name"
                                maxlength="60"
                                placeholder="نام خانوادگی"
                            >
                        </div>

                        <button
                            type="button"
                            class="danger"
                            onclick="removeMember(this)"
                        >
                            حذف
                        </button>

                    </div>

                </div>

                {% endif %}

            </div>

            <button
                type="button"
                class="teal"
                onclick="addMember()"
            >
                + افزودن عضو خانواده
            </button>

        </section>

        <button class="primary full" type="submit">
            ذخیره اطلاعات
        </button>

    </form>

    <form method="post" action="{{ url_for('family_logout') }}">

        <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">

        <button class="outline full" type="submit">
            خروج از حساب خانواده
        </button>

    </form>

</div>

<script>

const membersArea = document.getElementById("members-area");

document.querySelectorAll('input[name="attending"]').forEach((radio) => {

    radio.addEventListener("change", () => {

        membersArea.classList.toggle(
            "hidden",
            radio.value !== "yes"
        );

    });

});


function addMember() {

    const row = document.createElement("div");

    row.className = "member";

    row.innerHTML = `
        <div class="member-grid">

            <div>
                <label>نام</label>
                <input name="first_name" maxlength="50" placeholder="نام">
            </div>

            <div>
                <label>نام خانوادگی</label>
                <input name="last_name" maxlength="60" placeholder="نام خانوادگی">
            </div>

            <button
                type="button"
                class="danger"
                onclick="removeMember(this)"
            >
                حذف
            </button>

        </div>
    `;

    document.getElementById("members-list").appendChild(row);

}


function removeMember(button) {

    const list = document.getElementById("members-list");

    const rows = list.querySelectorAll(".member");

    if (rows.length > 1) {

        button.closest(".member").remove();

    } else {

        rows[0].querySelectorAll("input").forEach(
            (input) => input.value = ""
        );

    }

}

</script>
'''


# -------------------- ADMIN LOGIN PAGE --------------------

ADMIN_LOGIN_TEMPLATE = r'''
<div class="card">

    <div class="kicker">ADMIN ACCESS</div>

    <h2>ورود مدیر سامانه</h2>

    <form method="post">

        <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">

        <label>رمز مدیریت</label>

        <input
            type="password"
            name="password"
            placeholder="رمز مدیر"
            required
        >

        <button class="primary full" type="submit">
            ورود به پنل مدیریت
        </button>

    </form>

    <p class="muted">
        <a href="{{ url_for('home') }}">بازگشت به صفحه خانواده‌ها</a>
    </p>

</div>
'''


# -------------------- ADMIN DASHBOARD PAGE --------------------

ADMIN_TEMPLATE = r'''
<div class="card">

    <div class="kicker">ADMIN DASHBOARD</div>

    <h2>پنل مدیریت الغدیر واک</h2>

    <div class="stats">

        <div class="stat">
            <strong>{{ total }}</strong>
            <span>کل خانواده‌ها</span>
        </div>

        <div class="stat">
            <strong>{{ yes }}</strong>
            <span>خانواده‌های حاضر</span>
        </div>

        <div class="stat">
            <strong>{{ total_people }}</strong>
            <span>افراد شرکت‌کننده</span>
        </div>

        <div class="stat">
            <strong>{{ no }}</strong>
            <span>خانواده‌های غایب</span>
        </div>

    </div>

    <div class="actions">

        <a class="btn teal" href="{{ url_for('export_excel') }}">
            دانلود گزارش Excel
        </a>

        <form method="post" action="{{ url_for('admin_logout') }}">

            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">

            <button class="outline" type="submit">
                خروج از مدیریت
            </button>

        </form>

    </div>

    <h3>جست‌وجوی خانواده</h3>

    <form class="search-form" method="get" action="{{ url_for('admin') }}">

        <input
            name="q"
            value="{{ search }}"
            placeholder="جست‌وجو با شماره یا نام اعضا"
        >

        <button class="primary" type="submit">
            جست‌وجو
        </button>

        <a class="btn outline" href="{{ url_for('admin') }}">
            نمایش همه
        </a>

    </form>

    <h3>فهرست خانواده‌ها</h3>

    <div class="table-wrap">

        <table>

            <thead>

                <tr>
                    <th>شماره موبایل</th>
                    <th>وضعیت</th>
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

                            <span class="status-yes">
                                شرکت می‌کند
                            </span>

                        {% else %}

                            <span class="status-no">
                                شرکت نمی‌کند
                            </span>

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

                            <span class="muted">
                                عضوی ثبت نشده
                            </span>

                        {% endif %}

                    </td>

                    <td>

                        <form
                            method="post"
                            action="{{ url_
