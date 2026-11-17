
import os
import re
import io
import csv
import hmac
import secrets
from datetime import datetime
from functools import wraps

from flask import (
    Flask, request, redirect, session,
    render_template_string, send_file, url_for
)
from supabase import create_client
from openpyxl import Workbook

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

STYLE = """
<style>
*{box-sizing:border-box}
body{margin:0;background:#f0f5f1;color:#203b2e;
font-family:Tahoma,Arial;direction:rtl}
header{background:#123d2b;color:white;padding:25px 12px;text-align:center}
header h1{font-size:25px;margin:8px}
header p{color:#d4eadb}
main{max-width:750px;margin:25px auto;padding:12px}
.card{background:white;border-radius:16px;padding:22px;
box-shadow:0 5px 22px #00000010;margin-bottom:18px}
h2{text-align:center}
input{width:100%;padding:13px;border:1px solid #ccd9ce;
border-radius:9px;margin:7px 0;font-size:15px}
button,.btn{display:block;width:100%;padding:13px;
border:0;border-radius:9px;background:#176b43;color:white;
font-size:15px;text-align:center;text-decoration:none;
margin-top:10px;cursor:pointer}
.no{background:#a83c3c}
.member{background:#f0f5f1;padding:12px;border-radius:10px;margin:10px 0}
.stat{background:#e7f4eb;padding:14px;border-radius:10px;margin:8px 0}
table{width:100%;border-collapse:collapse;font-size:13px}
td,th{padding:9px;border-bottom:1px solid #ddd;text-align:right}
.table-wrap{overflow-x:auto}
footer{text-align:center;color:#718078;padding:25px;font-size:12px}
.small{font-size:12px;color:#66756b}
</style>
"""

BASE = """
<!doctype html>
<html lang="fa">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ALGHADIR FAMILY WALK</title>
</head>
<body>
<header>
<h1>ALGHADIR FAMILY WALK</h1>
<p>سامانه پیاده‌روی خانوادگی الغدیر</p>
</header>
<main>{{ body|safe }}</main>
<footer>سازنده: علیرضا روستا</footer>
</body>
</html>
"""

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
            <a class="btn" href="/">بازگشت</a>
            </div>
            """)

        if not hmac.compare_digest(password, FAMILY_PASSWORD):
            return page("""
            <div class="card">
            <h2>رمز ورود اشتباه است</h2>
            <a class="btn" href="/">تلاش دوباره</a>
            </div>
            """)

        session.clear()
        session["phone"] = phone
        session["csrf"] = secrets.token_urlsafe(32)
        return redirect("/register")

    return page("""
    <div class="card">
    <h2>به الغدیر فمیلی واک خوش آمدید 🌿</h2>
    <p>شماره تلفن و رمز عمومی ثبت‌نام را وارد کنید.</p>
    <form method="post">
    <input type="hidden" name="csrf" value="{{ token }}">
    <input name="phone" placeholder="09123456789"
    inputmode="numeric" maxlength="11" required>
    <input type="password" name="password"
    placeholder="رمز عمومی ثبت‌نام" required>
    <button>ورود به سامانه</button>
    </form>
    </div>
    """, token=csrf_token())

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

                if not first or not last or len(first) > 50 or len(last) > 50:
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
                        "first_name": m["first_name"],
                        "last_name": m["last_name"]
                    }
                    for m in new_members
                ]
                db.table("family_members").insert(rows).execute()

            session.pop("phone", None)

            return page("""
            <div class="card">
            <h2>✅ اطلاعات شما ثبت شد</h2>
            <p style="text-align:center">از همراهی شما سپاسگزاریم.</p>
            <a class="btn" href="/">بازگشت</a>
            </div>
            """)

        except Exception:
            return page("""
            <div class="card">
            <h2>ثبت اطلاعات انجام نشد</h2>
            <p>لطفاً دوباره تلاش کنید.</p>
            <a class="btn" href="/register">تلاش مجدد</a>
            </div>
            """)

    if family and family["attending"]:
        initial_members = members
    else:
        initial_members = [{"first_name": "", "last_name": ""}]

    return page("""
    <div class="card">
    <h2>{{ heading }}</h2>
    <p>شماره: {{ phone }}</p>

    <form method="post">
    <input type="hidden" name="csrf" value="{{ token }}">
    <input type="hidden" name="answer" value="yes">

    <div id="members">
    {% for m in initial_members %}
    <div class="member">
    <input name="first_name" placeholder="نام"
    value="{{ m.first_name }}" required>
    <input name="last_name" placeholder="نام خانوادگی"
    value="{{ m.last_name }}" required>
    </div>
    {% endfor %}
    </div>

    <button type="button" onclick="addMember()">+ افزودن عضو خانواده</button>
    <button type="submit">بله، شرکت می‌کنیم / ذخیره تغییرات</button>
    </form>

    <form method="post">
    <input type="hidden" name="csrf" value="{{ token }}">
    <input type="hidden" name="answer" value="no">
    <button class="no">خیر، امکان حضور نداریم</button>
    </form>

    <a class="btn" href="/leave">خروج</a>
    </div>

    <script>
    function addMember(){
      const box=document.getElementById('members');
      if(box.children.length >= {{ max_members }}){
        alert('حداکثر تعداد اعضا {{ max_members }} نفر است.');
        return;
      }
      const div=document.createElement('div');
      div.className='member';
      const first=document.createElement('input');
      first.name='first_name';
      first.placeholder='نام';
      first.required=true;
      const last=document.createElement('input');
      last.name='last_name';
      last.placeholder='نام خانوادگی';
      last.required=true;
      div.append(first,last);
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

@app.route("/leave")
def leave():
    session.clear()
    return redirect("/")

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
        <h2>ورود مدیریت</h2>
        <form method="post">
        <input type="hidden" name="csrf" value="{{ token }}">
        <input type="password" name="password"
        placeholder="رمز مدیریت" required>
        <button>ورود</button>
        </form>
        </div>
        """, token=csrf_token())

    try:
        families = db.table("families").select("*").order(
            "created_at", desc=True
        ).execute().data

        members = db.table("family_members").select("*").execute().data

        search = request.args.get("q", "").strip()
        if search:
            families = [
                f for f in families
                if search in f["phone"]
            ]

        yes = sum(1 for f in families if f["attending"])
        no = len(families) - yes

        return page("""
        <div class="card">
        <h2>پنل مدیریت الغدیر</h2>

        <div class="stat">کل خانواده‌ها: {{ total }}</div>
        <div class="stat">خانواده‌های شرکت‌کننده: {{ yes }}</div>
        <div class="stat">خانواده‌های غایب: {{ no }}</div>

        <a class="btn" href="/export">دانلود فایل Excel</a>

        <form method="get">
        <input name="q" placeholder="جست‌وجوی شماره تلفن"
        value="{{ search }}">
        <button>جست‌وجو</button>
        </form>

        <h3>فهرست ثبت‌نام‌ها</h3>
        <div class="table-wrap">
        <table>
        <tr><th>تلفن</th><th>پاسخ</th><th>اعضا</th><th>حذف</th></tr>
        {% for f in families %}
        <tr>
        <td>{{ f.phone }}</td>
        <td>{{ "بله" if f.attending else "خیر" }}</td>
        <td>
        {% for m in members %}
          {% if m.family_id == f.id %}
            {{ m.first_name }} {{ m.last_name }}<br>
          {% endif %}
        {% endfor %}
        </td>
        <td>
        <form method="post" action="/delete/{{ f.id }}"
        onsubmit="return confirm('این ثبت‌نام حذف شود؟')">
        <input type="hidden" name="csrf" value="{{ token }}">
        <button class="no">حذف</button>
        </form>
        </td>
        </tr>
        {% endfor %}
        </table>
        </div>

        <form method="post" action="/admin-logout">
        <input type="hidden" name="csrf" value="{{ token }}">
        <button class="no">خروج مدیر</button>
        </form>
        </div>
        """,
        families=families,
        members=members,
        total=len(families),
        yes=yes,
        no=no,
        search=search,
        token=csrf_token()
        )

    except Exception:
        return page("<div class='card'>خطا در دریافت اطلاعات</div>")

@app.route("/delete/<int:family_id>", methods=["POST"])
@admin_required
def delete_family(family_id):
    if not check_csrf():
        return "درخواست نامعتبر است.", 400

    try:
        db.table("families").delete().eq("id", family_id).execute()
        return redirect("/admin")
    except Exception:
        return "حذف انجام نشد.", 500

@app.route("/export")
@admin_required
def export():
    families = db.table("families").select("*").execute().data
    members = db.table("family_members").select("*").execute().data

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ثبت نام خانواده ها"

    sheet.append(["شماره تلفن", "پاسخ", "اعضای خانواده"])

    for f in families:
        names = [
            m["first_name"] + " " + m["last_name"]
            for m in members if m["family_id"] == f["id"]
        ]
        sheet.append([
            f["phone"],
            "بله" if f["attending"] else "خیر",
            "، ".join(names)
        ])

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="alghadir-family-walk.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin-logout", methods=["POST"])
def admin_logout():
    if not check_csrf():
        return "درخواست نامعتبر است.", 400
    session.clear()
    return redirect("/admin")

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000))
    )
        
