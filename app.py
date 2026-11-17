
import os
import re
from flask import Flask, request, redirect, session, render_template_string, flash
from supabase import create_client

app = Flask(__name__)
app.secret_key = os.environ["FLASK_SECRET_KEY"]

db = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SERVICE_ROLE_KEY"]
)

STYLE = """
<style>
*{box-sizing:border-box}
body{margin:0;background:#f0f5f1;color:#203b2e;
font-family:Tahoma,Arial;direction:rtl}
header{background:#123d2b;color:white;padding:28px 15px;
text-align:center}
header h1{font-size:30px;margin:8px}
header p{color:#d4eadb}
main{max-width:520px;margin:30px auto;padding:15px}
.card{background:white;border-radius:18px;padding:23px;
box-shadow:0 5px 22px #00000010;margin-bottom:18px}
h2{text-align:center}
input{width:100%;padding:14px;border:1px solid #ccd9ce;
border-radius:10px;margin:9px 0;font-size:16px}
button,.btn{display:block;width:100%;padding:14px;
border:0;border-radius:10px;background:#176b43;
color:white;font-size:16px;text-align:center;
text-decoration:none;margin-top:12px;cursor:pointer}
.no{background:#a83c3c}
.member{background:#f0f5f1;padding:12px;border-radius:12px;
margin:12px 0}
footer{text-align:center;color:#718078;padding:25px;font-size:12px}
.stat{background:#e7f4eb;padding:15px;border-radius:12px;
margin:8px 0}
table{width:100%;border-collapse:collapse;font-size:13px}
td,th{padding:10px;border-bottom:1px solid #ddd;text-align:right}
</style>
"""

def page(content, **kwargs):
    return render_template_string(
        STYLE + """
        <header>
        <h1>ALGHADIR FAMILY WALK</h1>
        <p>سامانه پیاده‌روی خانوادگی الغدیر</p>
        </header>
        <main>{{ content|safe }}</main>
        <footer>سازنده: علیرضا روستا</footer>
        """,
        content=render_template_string(content, **kwargs)
    )

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        phone = request.form.get("phone", "").strip()

        if not re.fullmatch(r"09[0-9]{9}", phone):
            return page("""
            <div class="card">
            <h2>شماره تلفن معتبر نیست</h2>
            <a class="btn" href="/">بازگشت</a>
            </div>
            """)

        session["phone"] = phone
        return redirect("/register")

    return page("""
    <div class="card">
    <h2>به الغدیر فمیلی واک خوش آمدید 🌿</h2>
    <p>برای ورود، شماره تلفن خود را وارد کنید.</p>
    <form method="post">
    <input name="phone" placeholder="09123456789"
    inputmode="numeric" maxlength="11" required>
    <button>ورود به سامانه</button>
    </form>
    </div>
    """)

@app.route("/register", methods=["GET", "POST"])
def register():
    phone = session.get("phone")
    if not phone:
        return redirect("/")

    if request.method == "POST":
        answer = request.form.get("answer")

        if answer not in ("yes", "no"):
            return redirect("/register")

        try:
            existing = db.table("families").select("id").eq(
                "phone", phone
            ).execute()

            if existing.data:
                return page("""
                <div class="card">
                <h2>قبلاً پاسخ شما ثبت شده است.</h2>
                <p>برای اصلاح اطلاعات با مدیریت تماس بگیرید.</p>
                <a class="btn" href="/">بازگشت</a>
                </div>
                """)

            result = db.table("families").insert({
                "phone": phone,
                "attending": answer == "yes"
            }).execute()

            family_id = result.data[0]["id"]

            if answer == "yes":
                first_names = request.form.getlist("first_name")
                last_names = request.form.getlist("last_name")

                members = []
                for first, last in zip(first_names, last_names):
                    first, last = first.strip(), last.strip()
                    if first and last:
                        members.append({
                            "family_id": family_id,
                            "first_name": first,
                            "last_name": last
                        })

                if not members:
                    raise ValueError("نام حداقل یک عضو لازم است.")

                db.table("family_members").insert(members).execute()

            session.pop("phone", None)

            return page("""
            <div class="card">
            <h2>✅ پاسخ شما ثبت شد</h2>
            <p style="text-align:center">
            از همراهی شما سپاسگزاریم.
            </p>
            <a class="btn" href="/">بازگشت</a>
            </div>
            """)

        except Exception:
            return page("""
            <div class="card">
            <h2>ثبت پاسخ انجام نشد</h2>
            <p>لطفاً دوباره تلاش کنید یا با مدیر تماس بگیرید.</p>
            <a class="btn" href="/register">تلاش مجدد</a>
            </div>
            """)

    return page("""
    <div class="card">
    <h2>آیا در پیاده‌روی شرکت می‌کنید؟</h2>
    <form method="post">
    <input type="hidden" name="answer" value="yes">
    <div id="members">
    <div class="member">
    <input name="first_name" placeholder="نام" required>
    <input name="last_name" placeholder="نام خانوادگی" required>
    </div>
    </div>
    <button type="button" onclick="addMember()">
    + افزودن عضو خانواده
    </button>
    <button type="submit">بله، شرکت می‌کنیم</button>
    </form>
    <form method="post">
    <input type="hidden" name="answer" value="no">
    <button class="no">خیر، امکان حضور نداریم</button>
    </form>
    </div>
    <script>
    function addMember(){
      let div=document.createElement('div');
      div.className='member';
      div.innerHTML=
      '<input name="first_name" placeholder="نام" required>'+
      '<input name="last_name" placeholder="نام خانوادگی" required>';
      document.getElementById('members').appendChild(div);
    }
    </script>
    """)

@app.route("/admin", methods=["GET", "POST"])
def admin():
    if request.method == "POST":
        if request.form.get("password") == os.environ["ADMIN_PASSWORD"]:
            session["admin"] = True
            return redirect("/admin")
        return page("""
        <div class="card"><h2>رمز اشتباه است</h2>
        <a class="btn" href="/admin">تلاش مجدد</a></div>
        """)

    if not session.get("admin"):
        return page("""
        <div class="card">
        <h2>ورود مدیریت</h2>
        <form method="post">
        <input type="password" name="password"
        placeholder="رمز مدیریت" required>
        <button>ورود</button>
        </form>
        </div>
        """)

    try:
        families = db.table("families").select("*").order(
            "created_at", desc=True
        ).execute().data

        members = db.table("family_members").select("*").execute().data

        yes = sum(1 for f in families if f["attending"])
        no = len(families) - yes

        rows = ""
        for f in families:
            names = [
                m["first_name"] + " " + m["last_name"]
                for m in members if m["family_id"] == f["id"]
            ]
            names_text = "، ".join(names) if names else "-"
            status = "بله" if f["attending"] else "خیر"
            rows += (
                "<tr><td>" + f["phone"] +
                "</td><td>" + status +
                "</td><td>" + names_text + "</td></tr>"
            )

        return page("""
        <div class="card">
        <h2>پنل مدیریت الغدیر</h2>
        <div class="stat">کل خانواده‌ها: {{total}}</div>
        <div class="stat">شرکت‌کنندگان: {{yes}}</div>
        <div class="stat">غایبان: {{no}}</div>
        <h3>فهرست ثبت‌نام‌ها</h3>
        <table>
        <tr><th>تلفن</th><th>پاسخ</th><th>اعضا</th></tr>
        {{rows|safe}}
        </table>
        <a class="btn no" href="/logout">خروج مدیر</a>
        </div>
        """, total=len(families), yes=yes, no=no, rows=rows)

    except Exception:
        return page("<div class='card'>خطا در دریافت اطلاعات</div>")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/admin")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
  
