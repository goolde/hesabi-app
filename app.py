#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════╗
║         حسابي SaaS — نسخة الإنتاج               ║
║   Python + Flask + SQLite/PostgreSQL             ║
╚══════════════════════════════════════════════════╝

التثبيت:
  pip install -r requirements.txt

التشغيل المحلي:
  py app.py

النشر على Railway/Render:
  gunicorn app:app
"""

import sqlite3, json, os, hashlib, secrets, io, re
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, request, jsonify, session, send_file, g

# Excel & PDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER

# ============================================================
# إعداد التطبيق
# ============================================================
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))

# مسار قاعدة البيانات — يمكن تغييره لـ PostgreSQL في الإنتاج
DB_PATH = os.environ.get('DATABASE_URL', 'hesabi.db')

# إعدادات SMS (Unifonic أو Taqnyat)
SMS_PROVIDER = os.environ.get('SMS_PROVIDER', 'demo')  # demo / unifonic / taqnyat
SMS_API_KEY  = os.environ.get('SMS_API_KEY', '')
SMS_SENDER   = os.environ.get('SMS_SENDER', 'Hesabi')

# معلومات المطوّر
DEVELOPER_PHONE   = '0567867414'
DEVELOPER_NAME    = 'وجيه علي'
DEVELOPER_WHATSAPP = 'https://wa.me/966567867414'

# باقات الاشتراك
PLANS = {
    'free':         {'name':'مجانية',    'price':0,  'contacts':10,  'transactions':50,  'export':False},
    'professional': {'name':'احترافية',  'price':29, 'contacts':0,   'transactions':0,   'export':True},
    'business':     {'name':'أعمال',     'price':79, 'contacts':0,   'transactions':0,   'export':True},
}

# ============================================================
# قاعدة البيانات
# ============================================================
def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db: db.close()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT    NOT NULL,
            phone         TEXT    UNIQUE NOT NULL,
            email         TEXT    DEFAULT '',
            business      TEXT    DEFAULT '',
            password      TEXT    NOT NULL,
            plan          TEXT    DEFAULT 'free',
            plan_expires  TEXT    DEFAULT '',
            is_active     INTEGER DEFAULT 1,
            otp_code      TEXT    DEFAULT '',
            otp_expires   TEXT    DEFAULT '',
            created       TEXT    DEFAULT (datetime('now')),
            last_login    TEXT    DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS contacts (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id       INTEGER NOT NULL,
            name          TEXT    NOT NULL,
            phone         TEXT    DEFAULT '',
            email         TEXT    DEFAULT '',
            type          TEXT    DEFAULT 'عميل',
            notes         TEXT    DEFAULT '',
            lat           REAL,
            lng           REAL,
            location_text TEXT    DEFAULT '',
            created       TEXT    DEFAULT (datetime('now')),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS transactions (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id       INTEGER NOT NULL,
            contact_id    INTEGER NOT NULL,
            contact_name  TEXT    DEFAULT '',
            amount        REAL    NOT NULL,
            currency      TEXT    DEFAULT 'ر.س',
            type          TEXT    NOT NULL,
            date          TEXT    DEFAULT '',
            due_date      TEXT    DEFAULT '',
            notes         TEXT    DEFAULT '',
            status        TEXT    DEFAULT 'pending',
            reminder      TEXT    DEFAULT 'none',
            created       TEXT    DEFAULT (datetime('now')),
            FOREIGN KEY (user_id)    REFERENCES users(id)    ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS notifications (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id  INTEGER NOT NULL,
            title    TEXT    DEFAULT '',
            body     TEXT    DEFAULT '',
            icon     TEXT    DEFAULT '📌',
            is_read  INTEGER DEFAULT 0,
            created  TEXT    DEFAULT (datetime('now')),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS settings (
            user_id  INTEGER PRIMARY KEY,
            data     TEXT    DEFAULT '{}',
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS payments (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id     INTEGER NOT NULL,
            plan        TEXT,
            amount      REAL,
            currency    TEXT    DEFAULT 'SAR',
            status      TEXT    DEFAULT 'pending',
            ref_id      TEXT    DEFAULT '',
            created     TEXT    DEFAULT (datetime('now')),
            FOREIGN KEY (user_id) REFERENCES users(id)
        );

        CREATE INDEX IF NOT EXISTS idx_contacts_user    ON contacts(user_id);
        CREATE INDEX IF NOT EXISTS idx_transactions_user ON transactions(user_id);
        CREATE INDEX IF NOT EXISTS idx_transactions_contact ON transactions(contact_id);
        CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id);
    """)
    db.commit()
    db.close()
    print("✅ قاعدة البيانات جاهزة")

# ============================================================
# دوال مساعدة
# ============================================================
def hash_pwd(p): return hashlib.sha256(p.encode()).hexdigest()
def gen_otp():   return str(secrets.randbelow(9000) + 1000)
def now_str():   return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def login_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if 'user_id' not in session:
            return jsonify({'error': 'غير مسجّل', 'redirect': '/'}), 401
        return f(*a, **kw)
    return dec

def check_plan_limit(user_id, resource):
    """التحقق من حدود الباقة"""
    db   = get_db()
    user = db.execute("SELECT plan FROM users WHERE id=?", (user_id,)).fetchone()
    if not user: return False
    plan = PLANS.get(user['plan'], PLANS['free'])
    if resource == 'contacts':
        limit = plan['contacts']
        if limit == 0: return True
        count = db.execute("SELECT COUNT(*) FROM contacts WHERE user_id=?", (user_id,)).fetchone()[0]
        return count < limit
    if resource == 'transactions':
        limit = plan['transactions']
        if limit == 0: return True
        count = db.execute("SELECT COUNT(*) FROM transactions WHERE user_id=?", (user_id,)).fetchone()[0]
        return count < limit
    if resource == 'export':
        return plan['export']
    return True

def add_notif(uid, title, body, icon='📌'):
    db = get_db()
    db.execute("INSERT INTO notifications (user_id,title,body,icon) VALUES (?,?,?,?)", (uid,title,body,icon))
    db.commit()

def get_settings(uid):
    db  = get_db()
    row = db.execute("SELECT data FROM settings WHERE user_id=?", (uid,)).fetchone()
    if row:
        try: return json.loads(row['data'])
        except: return {}
    return {}

def save_settings(uid, data):
    db = get_db()
    db.execute("INSERT INTO settings (user_id,data) VALUES (?,?) ON CONFLICT(user_id) DO UPDATE SET data=excluded.data",
               (uid, json.dumps(data, ensure_ascii=False)))
    db.commit()

def send_sms(phone, message):
    """إرسال SMS — يدعم Unifonic وTaqnyat"""
    if SMS_PROVIDER == 'demo':
        print(f"📱 [SMS Demo] To: {phone} | Msg: {message}")
        return True
    if SMS_PROVIDER == 'unifonic':
        import urllib.request, urllib.parse
        data = urllib.parse.urlencode({
            'AppSid': SMS_API_KEY, 'SenderID': SMS_SENDER,
            'Body': message, 'Recipient': phone, 'responseType': 'JSON'
        }).encode()
        try:
            req = urllib.request.urlopen('https://el.cloud.unifonic.com/rest/SMS/messages', data, timeout=10)
            return True
        except: return False
    if SMS_PROVIDER == 'taqnyat':
        import urllib.request
        payload = json.dumps({'recipients':[phone],'body':message,'sender':SMS_SENDER}).encode()
        req = urllib.request.Request('https://api.taqnyat.sa/v1/messages',
            data=payload, headers={'Authorization':f'Bearer {SMS_API_KEY}','Content-Type':'application/json'})
        try:
            urllib.request.urlopen(req, timeout=10)
            return True
        except: return False
    return False

def get_balance(uid, cid):
    db = get_db()
    rows = db.execute("""
        SELECT type, SUM(amount) as total FROM transactions
        WHERE user_id=? AND contact_id=? AND status!='settled' GROUP BY type
    """, (uid, cid)).fetchall()
    fm = om = 0.0
    for r in rows:
        if r['type']=='incoming': fm = r['total'] or 0
        else: om = r['total'] or 0
    return {'for_me':fm,'on_me':om,'net':fm-om}

# ============================================================
# قراءة ملف HTML
# ============================================================
def load_html():
    html_path = os.path.join(os.path.dirname(__file__), 'templates', 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        return f.read()

# ============================================================
# Routes — الصفحات
# ============================================================
@app.route('/')
def index():
    return load_html()

# ============================================================
# Routes — المصادقة
# ============================================================
@app.route('/api/register', methods=['POST'])
def register():
    d = request.get_json()
    name  = d.get('name','').strip()
    phone = d.get('phone','').strip()
    pwd   = d.get('password','')
    if not all([name, phone, pwd]):
        return jsonify({'error':'يرجى تعبئة جميع الحقول'})
    if len(pwd) < 6:
        return jsonify({'error':'كلمة المرور قصيرة (6 أحرف على الأقل)'})
    if not re.match(r'^05\d{8}$', phone):
        return jsonify({'error':'رقم الجوال غير صحيح (يجب أن يبدأ بـ 05)'})
    db = get_db()
    if db.execute("SELECT id FROM users WHERE phone=?", (phone,)).fetchone():
        return jsonify({'error':'رقم الجوال مسجّل مسبقاً'})
    otp     = gen_otp()
    expires = (datetime.now() + timedelta(minutes=5)).strftime('%Y-%m-%d %H:%M:%S')
    db.execute("INSERT INTO users (name,phone,password,otp_code,otp_expires) VALUES (?,?,?,?,?)",
               (name, phone, hash_pwd(pwd), otp, expires))
    db.commit()
    msg = f'حسابي: رمز التحقق الخاص بك هو {otp}. صالح لـ 5 دقائق.'
    send_sms(phone, msg)
    return jsonify({'success':True, 'otp_demo':otp, 'message':'تم إرسال رمز التحقق'})

@app.route('/api/verify-otp', methods=['POST'])
def verify_otp():
    d     = request.get_json()
    phone = d.get('phone','').strip()
    otp   = d.get('otp','').strip()
    db    = get_db()
    user  = db.execute("SELECT * FROM users WHERE phone=?", (phone,)).fetchone()
    if not user:
        return jsonify({'error':'المستخدم غير موجود'})
    if user['otp_code'] != otp:
        return jsonify({'error':'رمز التحقق غير صحيح'})
    if datetime.now() > datetime.strptime(user['otp_expires'], '%Y-%m-%d %H:%M:%S'):
        return jsonify({'error':'انتهت صلاحية الرمز — اطلب رمزاً جديداً'})
    db.execute("UPDATE users SET otp_code='',last_login=? WHERE id=?", (now_str(), user['id']))
    db.commit()
    session['user_id']   = user['id']
    session['user_name'] = user['name']
    add_notif(user['id'], 'مرحباً بك! 👋', f'أهلاً {user["name"]}، ابدأ بإضافة جهاتك ومعاملاتك')
    return jsonify({'success':True, 'user':{'id':user['id'],'name':user['name'],'plan':user['plan']}})

@app.route('/api/login', methods=['POST'])
def login():
    d     = request.get_json()
    phone = d.get('phone','').strip()
    pwd   = d.get('password','')
    db    = get_db()
    user  = db.execute("SELECT * FROM users WHERE phone=? AND password=?",
                       (phone, hash_pwd(pwd))).fetchone()
    if not user:
        return jsonify({'error':'الجوال أو كلمة المرور غير صحيحة'})
    if not user['is_active']:
        return jsonify({'error':'الحساب موقوف — تواصل مع الدعم'})
    db.execute("UPDATE users SET last_login=? WHERE id=?", (now_str(), user['id']))
    db.commit()
    session['user_id']   = user['id']
    session['user_name'] = user['name']
    return jsonify({'success':True, 'user':{'id':user['id'],'name':user['name'],'plan':user['plan']}})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success':True})

@app.route('/api/me')
def me():
    if 'user_id' not in session: return jsonify({'user':None})
    db   = get_db()
    user = db.execute("SELECT id,name,phone,email,business,plan FROM users WHERE id=?",
                      (session['user_id'],)).fetchone()
    if not user: return jsonify({'user':None})
    return jsonify({'user':dict(user)})

@app.route('/api/resend-otp', methods=['POST'])
def resend_otp():
    d     = request.get_json()
    phone = d.get('phone','').strip()
    db    = get_db()
    user  = db.execute("SELECT id,name FROM users WHERE phone=?", (phone,)).fetchone()
    if not user: return jsonify({'error':'المستخدم غير موجود'})
    otp     = gen_otp()
    expires = (datetime.now() + timedelta(minutes=5)).strftime('%Y-%m-%d %H:%M:%S')
    db.execute("UPDATE users SET otp_code=?,otp_expires=? WHERE id=?", (otp,expires,user['id']))
    db.commit()
    send_sms(phone, f'حسابي: رمز التحقق هو {otp}. صالح لـ 5 دقائق.')
    return jsonify({'success':True, 'otp_demo':otp})

# ============================================================
# Routes — الملف الشخصي
# ============================================================
@app.route('/api/profile', methods=['POST'])
@login_required
def profile():
    d    = request.get_json()
    uid  = session['user_id']
    name = d.get('name','').strip()
    if not name: return jsonify({'error':'الاسم مطلوب'})
    db = get_db()
    db.execute("UPDATE users SET name=?,email=?,business=? WHERE id=?",
               (name, d.get('email',''), d.get('business',''), uid))
    db.commit()
    session['user_name'] = name
    return jsonify({'success':True})

@app.route('/api/change-password', methods=['POST'])
@login_required
def change_password():
    d   = request.get_json()
    uid = session['user_id']
    db  = get_db()
    if not db.execute("SELECT id FROM users WHERE id=? AND password=?",
                      (uid, hash_pwd(d.get('current_password','')))).fetchone():
        return jsonify({'error':'كلمة المرور الحالية غير صحيحة'})
    new_pwd = d.get('new_password','')
    if len(new_pwd) < 6: return jsonify({'error':'كلمة المرور قصيرة جداً'})
    db.execute("UPDATE users SET password=? WHERE id=?", (hash_pwd(new_pwd), uid))
    db.commit()
    return jsonify({'success':True})

# ============================================================
# Routes — الجهات
# ============================================================
@app.route('/api/contacts', methods=['GET','POST'])
@login_required
def contacts():
    uid = session['user_id']
    db  = get_db()
    if request.method == 'GET':
        rows = db.execute("SELECT * FROM contacts WHERE user_id=? ORDER BY name", (uid,)).fetchall()
        return jsonify({'data':[dict(r) for r in rows]})
    d    = request.get_json()
    name = d.get('name','').strip()
    if not name: return jsonify({'error':'الاسم مطلوب'})
    if not check_plan_limit(uid,'contacts'):
        return jsonify({'error':'وصلت للحد الأقصى في باقتك — يرجى الترقية','upgrade':True})
    db.execute("INSERT INTO contacts (user_id,name,phone,email,type,notes,lat,lng,location_text) VALUES (?,?,?,?,?,?,?,?,?)",
               (uid,name,d.get('phone',''),d.get('email',''),d.get('type','عميل'),d.get('notes',''),d.get('lat'),d.get('lng'),d.get('location_text','')))
    db.commit()
    add_notif(uid,'جهة جديدة 👤',f'تمت إضافة {name} كـ{d.get("type","عميل")}')
    return jsonify({'success':True})

@app.route('/api/contacts/<int:cid>', methods=['POST'])
@login_required
def contact_op(cid):
    uid = session['user_id']
    db  = get_db()
    d   = request.get_json()
    if d.get('_method') == 'DELETE':
        db.execute("DELETE FROM contacts WHERE id=? AND user_id=?", (cid,uid))
        db.commit()
        return jsonify({'success':True})
    db.execute("UPDATE contacts SET name=?,phone=?,email=?,type=?,notes=?,lat=?,lng=?,location_text=? WHERE id=? AND user_id=?",
               (d.get('name'),d.get('phone',''),d.get('email',''),d.get('type','عميل'),d.get('notes',''),
                d.get('lat'),d.get('lng'),d.get('location_text',''),cid,uid))
    db.commit()
    return jsonify({'success':True})

@app.route('/api/contacts/<int:cid>/balance')
@login_required
def contact_balance(cid):
    return jsonify(get_balance(session['user_id'], cid))

# ============================================================
# Routes — المعاملات
# ============================================================
@app.route('/api/transactions', methods=['GET','POST'])
@login_required
def transactions():
    uid = session['user_id']
    db  = get_db()
    if request.method == 'GET':
        rows = db.execute("""
            SELECT t.*, c.name as contact_name FROM transactions t
            LEFT JOIN contacts c ON c.id=t.contact_id
            WHERE t.user_id=? ORDER BY t.created DESC""", (uid,)).fetchall()
        return jsonify({'data':[dict(r) for r in rows]})
    d   = request.get_json()
    cid = d.get('contact_id')
    amt = float(d.get('amount',0))
    if not cid or amt <= 0: return jsonify({'error':'بيانات غير مكتملة'})
    if not check_plan_limit(uid,'transactions'):
        return jsonify({'error':'وصلت للحد الأقصى في باقتك — يرجى الترقية','upgrade':True})
    c = db.execute("SELECT name FROM contacts WHERE id=?", (cid,)).fetchone()
    cname = c['name'] if c else '—'
    db.execute("""INSERT INTO transactions
        (user_id,contact_id,contact_name,amount,currency,type,date,due_date,notes,status,reminder)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (uid,cid,cname,amt,d.get('currency','ر.س'),d.get('type','incoming'),
         d.get('date',str(date.today())),d.get('due_date',''),d.get('notes',''),
         d.get('status','pending'),d.get('reminder','none')))
    db.commit()
    lbl = 'لي من' if d.get('type')=='incoming' else 'علي من'
    add_notif(uid,'معاملة جديدة 📋',f'{lbl} {cname}: {amt} {d.get("currency","ر.س")}')
    return jsonify({'success':True})

@app.route('/api/transactions/<int:tid>', methods=['POST'])
@login_required
def tx_op(tid):
    uid = session['user_id']
    db  = get_db()
    d   = request.get_json()
    if d.get('_method') == 'DELETE':
        db.execute("DELETE FROM transactions WHERE id=? AND user_id=?", (tid,uid))
        db.commit()
        return jsonify({'success':True})
    db.execute("""UPDATE transactions
        SET contact_id=?,amount=?,currency=?,type=?,date=?,due_date=?,notes=?,status=?
        WHERE id=? AND user_id=?""",
        (d.get('contact_id'),d.get('amount'),d.get('currency','ر.س'),d.get('type'),
         d.get('date'),d.get('due_date',''),d.get('notes',''),d.get('status'),tid,uid))
    db.commit()
    return jsonify({'success':True})

# ============================================================
# Routes — الكشوف
# ============================================================
@app.route('/api/statements', methods=['POST'])
@login_required
def statements():
    uid = session['user_id']
    db  = get_db()
    d   = request.get_json()
    q   = "SELECT t.*,c.name as contact_name,c.type as contact_type FROM transactions t LEFT JOIN contacts c ON c.id=t.contact_id WHERE t.user_id=?"
    p   = [uid]
    if d.get('from'):    q += " AND t.date>=?"; p.append(d['from'])
    if d.get('to'):      q += " AND t.date<=?"; p.append(d['to'])
    if d.get('contact_id'): q += " AND t.contact_id=?"; p.append(d['contact_id'])
    if d.get('type'):    q += " AND t.type=?"; p.append(d['type'])
    if d.get('status'):  q += " AND t.status=?"; p.append(d['status'])
    q += " ORDER BY t.date DESC,t.created DESC"
    rows = db.execute(q, p).fetchall()
    return jsonify({'data':[dict(r) for r in rows]})

# ============================================================
# Routes — تصدير PDF
# ============================================================
@app.route('/api/export/pdf')
@login_required
def export_pdf():
    uid = session['user_id']
    if not check_plan_limit(uid,'export'):
        return jsonify({'error':'التصدير متاح في الباقة الاحترافية فقط'}), 403
    db   = get_db()
    user = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    s    = get_settings(uid)
    frm  = request.args.get('from','')
    to   = request.args.get('to','')
    cid  = request.args.get('contact_id','')
    q    = "SELECT t.*,c.name as contact_name FROM transactions t LEFT JOIN contacts c ON c.id=t.contact_id WHERE t.user_id=?"
    p    = [uid]
    if frm: q += " AND t.date>=?"; p.append(frm)
    if to:  q += " AND t.date<=?"; p.append(to)
    if cid: q += " AND t.contact_id=?"; p.append(cid)
    rows = db.execute(q+" ORDER BY t.date DESC", p).fetchall()
    gold   = colors.HexColor('#f6c90e')
    dark   = colors.HexColor('#1a2340')
    green_ = colors.HexColor('#48bb78')
    red_   = colors.HexColor('#fc8181')
    buf    = io.BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle('t', fontName='Helvetica-Bold', fontSize=18, alignment=TA_CENTER, textColor=dark, spaceAfter=8)
    sub_s   = ParagraphStyle('s', fontName='Helvetica', fontSize=10, alignment=TA_CENTER, textColor=colors.grey, spaceAfter=16)
    story   = []
    biz = user['business'] or user['name']
    story.append(Paragraph(f'Hesabi | {biz}', title_s))
    story.append(Paragraph(f'Account Statement — {datetime.now().strftime("%Y-%m-%d")}', sub_s))
    story.append(HRFlowable(width="100%", thickness=2, color=gold, spaceAfter=12))
    fm = sum(float(r['amount']) for r in rows if r['type']=='incoming' and r['status']!='settled')
    om = sum(float(r['amount']) for r in rows if r['type']=='outgoing' and r['status']!='settled')
    nt = fm - om
    cur = s.get('currency','ر.س')
    summ = Table([
        ['For Me (لي)', f'{fm:,.2f} {cur}'],
        ['On Me (علي)', f'{om:,.2f} {cur}'],
        ['Net (الصافي)', f'{nt:+,.2f} {cur}'],
        ['Total Rows', str(len(rows))],
    ], colWidths=[10*cm,7*cm])
    summ.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),'Helvetica'),
        ('FONTSIZE',(0,0),(-1,-1),11),
        ('FONTNAME',(1,0),(1,-1),'Helvetica-Bold'),
        ('TEXTCOLOR',(1,0),(1,0),green_),
        ('TEXTCOLOR',(1,1),(1,1),red_),
        ('TEXTCOLOR',(1,2),(1,2),green_ if nt>=0 else red_),
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[colors.white,colors.HexColor('#f5f5f5')]),
        ('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#dddddd')),
        ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
    ]))
    story.append(summ)
    story.append(Spacer(1,16))
    tdata = [['Date','Contact','Notes','For Me','On Me','Status']]
    for r in rows:
        isIn = r['type']=='incoming'
        tdata.append([r['date'] or '—', r['contact_name'] or '—', (r['notes'] or '—')[:28],
            f"{r['amount']} {r['currency']}" if isIn else '—',
            f"{r['amount']} {r['currency']}" if not isIn else '—',
            {'pending':'Pending','partial':'Partial','settled':'Settled'}.get(r['status'],'')])
    mt = Table(tdata, colWidths=[2.5*cm,3.5*cm,4*cm,3*cm,3*cm,2*cm], repeatRows=1)
    mt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),dark),('TEXTCOLOR',(0,0),(-1,0),gold),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),10),
        ('FONTNAME',(0,1),(-1,-1),'Helvetica'),('FONTSIZE',(0,1),(-1,-1),9),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f9f9f9')]),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#e0e0e0')),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    story.append(mt)
    story.append(Spacer(1,16))
    story.append(HRFlowable(width="100%",thickness=1,color=colors.grey))
    foot = ParagraphStyle('f',fontName='Helvetica',fontSize=9,alignment=TA_CENTER,textColor=colors.grey)
    story.append(Paragraph(f'Hesabi App | {user["name"]} | {datetime.now().strftime("%Y-%m-%d")}', foot))
    doc.build(story)
    buf.seek(0)
    return send_file(buf, mimetype='application/pdf', as_attachment=True,
                     download_name=f'hesabi_{date.today()}.pdf')

# ============================================================
# Routes — تصدير Excel
# ============================================================
@app.route('/api/export/excel')
@login_required
def export_excel():
    uid = session['user_id']
    if not check_plan_limit(uid,'export'):
        return jsonify({'error':'التصدير متاح في الباقة الاحترافية فقط'}), 403
    db   = get_db()
    user = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    s    = get_settings(uid)
    frm  = request.args.get('from','')
    to   = request.args.get('to','')
    cid  = request.args.get('contact_id','')
    q    = "SELECT t.*,c.name as contact_name FROM transactions t LEFT JOIN contacts c ON c.id=t.contact_id WHERE t.user_id=?"
    p    = [uid]
    if frm: q += " AND t.date>=?"; p.append(frm)
    if to:  q += " AND t.date<=?"; p.append(to)
    if cid: q += " AND t.contact_id=?"; p.append(cid)
    rows = db.execute(q+" ORDER BY t.date DESC", p).fetchall()
    cur  = s.get('currency','ر.س')
    wb   = Workbook()
    ws   = wb.active
    ws.title = "Statement"
    ws.sheet_view.rightToLeft = True
    gold_f  = PatternFill("solid", fgColor="F6C90E")
    dark_f  = PatternFill("solid", fgColor="1A2340")
    hdr_fnt = Font(bold=True, size=11)
    ctr     = Alignment(horizontal="center", vertical="center")
    rt      = Alignment(horizontal="right", vertical="center")
    thin    = Border(left=Side(style='thin',color='DDDDDD'),right=Side(style='thin',color='DDDDDD'),
                     top=Side(style='thin',color='DDDDDD'),bottom=Side(style='thin',color='DDDDDD'))
    ws.merge_cells('A1:H1')
    ws['A1'] = f'حسابي — {user["business"] or user["name"]}'
    ws['A1'].font = Font(bold=True,size=16,color="F6C90E"); ws['A1'].fill = dark_f; ws['A1'].alignment = ctr
    ws.row_dimensions[1].height = 40
    ws.merge_cells('A2:H2')
    ws['A2'] = f'تاريخ التصدير: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
    ws['A2'].font = Font(size=10,color="AAAAAA"); ws['A2'].fill = PatternFill("solid",fgColor="0F1628"); ws['A2'].alignment = ctr
    fm = sum(float(r['amount']) for r in rows if r['type']=='incoming' and r['status']!='settled')
    om = sum(float(r['amount']) for r in rows if r['type']=='outgoing' and r['status']!='settled')
    ws.row_dimensions[4].height = 30
    for ci,(lbl,val,color) in enumerate([('إجمالي لي',f'{fm:,.2f} {cur}','48BB78'),
                                          ('إجمالي علي',f'{om:,.2f} {cur}','FC8181'),
                                          ('الصافي',f'{fm-om:+,.2f} {cur}','48BB78' if fm>=om else 'FC8181'),
                                          ('المعاملات',str(len(rows)),'F6C90E')],1):
        ws.merge_cells(f'{get_column_letter(ci*2-1)}4:{get_column_letter(ci*2)}4')
        c = ws[f'{get_column_letter(ci*2-1)}4']
        c.value = f'{lbl}: {val}'
        c.font = Font(bold=True,size=11,color=color)
        c.fill = PatternFill("solid",fgColor="151D35"); c.alignment = ctr
    hdrs = ['التاريخ','الجهة','النوع','المبلغ','العملة','البيان','الحالة','الاستحقاق']
    for ci,h in enumerate(hdrs,1):
        c = ws.cell(row=6,column=ci,value=h)
        c.font = Font(bold=True,size=10); c.fill = gold_f; c.alignment = ctr; c.border = thin
    ws.row_dimensions[6].height = 28
    sl = {'pending':'معلقة','partial':'جزئية','settled':'مسوّاة'}
    tl = {'incoming':'لي 📥','outgoing':'علي 📤'}
    for i,r in enumerate(rows,7):
        isIn = r['type']=='incoming'
        fill = PatternFill("solid",fgColor="E8F5E9" if isIn else "FFEBEE")
        vals = [r['date'] or '',r['contact_name'] or '',tl.get(r['type'],r['type']),
                float(r['amount']),r['currency'],r['notes'] or '',sl.get(r['status'],r['status']),r['due_date'] or '']
        for ci,v in enumerate(vals,1):
            c = ws.cell(row=i,column=ci,value=v)
            c.fill = fill if ci<=4 else PatternFill()
            c.alignment = rt; c.border = thin
            if ci==4: c.font = Font(bold=True,color="1B5E20" if isIn else "B71C1C"); c.number_format='#,##0.00'
        ws.row_dimensions[i].height = 22
    for ci,w in enumerate([14,20,12,14,10,28,12,14],1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'hesabi_{date.today()}.xlsx')

# ============================================================
# Routes — الإشعارات
# ============================================================
@app.route('/api/notifications')
@login_required
def notifs():
    db   = get_db()
    rows = db.execute("SELECT * FROM notifications WHERE user_id=? ORDER BY created DESC LIMIT 50",
                      (session['user_id'],)).fetchall()
    return jsonify({'data':[dict(r) for r in rows]})

@app.route('/api/notifications/read-all', methods=['POST'])
@login_required
def read_notifs():
    db = get_db()
    db.execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (session['user_id'],))
    db.commit()
    return jsonify({'success':True})

# ============================================================
# Routes — الإعدادات
# ============================================================
@app.route('/api/settings', methods=['GET','POST'])
@login_required
def settings():
    uid = session['user_id']
    if request.method == 'GET':
        return jsonify({'data': get_settings(uid)})
    data = get_settings(uid)
    data.update(request.get_json())
    save_settings(uid, data)
    return jsonify({'success':True})

# ============================================================
# Routes — البحث
# ============================================================
@app.route('/api/search')
@login_required
def search():
    q   = request.args.get('q','').strip()
    uid = session['user_id']
    if len(q) < 2: return jsonify({'contacts':[],'transactions':[]})
    db = get_db()
    contacts = db.execute("SELECT * FROM contacts WHERE user_id=? AND (name LIKE ? OR phone LIKE ?) LIMIT 6",
                          (uid,f'%{q}%',f'%{q}%')).fetchall()
    txs = db.execute("""SELECT t.*,c.name as contact_name FROM transactions t
        LEFT JOIN contacts c ON c.id=t.contact_id
        WHERE t.user_id=? AND (c.name LIKE ? OR t.notes LIKE ?)
        ORDER BY t.created DESC LIMIT 6""", (uid,f'%{q}%',f'%{q}%')).fetchall()
    return jsonify({'contacts':[dict(r) for r in contacts],'transactions':[dict(r) for r in txs]})

# ============================================================
# Routes — الاشتراكات والدفع (Moyasar)
# ============================================================

MOYASAR_API_KEY    = os.environ.get('MOYASAR_API_KEY', '')
MOYASAR_SECRET_KEY = os.environ.get('MOYASAR_SECRET_KEY', '')
APP_URL            = os.environ.get('APP_URL', 'http://localhost:5000')

@app.route('/api/plans')
def plans():
    return jsonify({'plans': PLANS})

@app.route('/api/my-plan')
@login_required
def my_plan():
    db   = get_db()
    uid  = session['user_id']
    user = db.execute("SELECT plan,plan_expires FROM users WHERE id=?", (uid,)).fetchone()
    plan_info = PLANS.get(user['plan'], PLANS['free'])
    contacts_count = db.execute("SELECT COUNT(*) FROM contacts WHERE user_id=?", (uid,)).fetchone()[0]
    txs_count = db.execute("SELECT COUNT(*) FROM transactions WHERE user_id=?", (uid,)).fetchone()[0]
    # التحقق من انتهاء الاشتراك
    expired = False
    if user['plan_expires'] and user['plan_expires'] != '':
        try:
            exp_date = datetime.strptime(user['plan_expires'], '%Y-%m-%d')
            if datetime.now() > exp_date:
                expired = True
                db.execute("UPDATE users SET plan='free',plan_expires='' WHERE id=?", (uid,))
                db.commit()
        except: pass
    return jsonify({
        'plan':        user['plan'],
        'plan_name':   plan_info['name'],
        'plan_expires':user['plan_expires'],
        'expired':     expired,
        'limits':      plan_info,
        'usage': {'contacts': contacts_count, 'transactions': txs_count}
    })

@app.route('/api/subscribe', methods=['POST'])
@login_required
def subscribe():
    """
    بدء عملية الدفع عبر Moyasar
    يرجع رابط الدفع لإعادة توجيه المستخدم
    """
    d    = request.get_json()
    plan = d.get('plan', '')
    uid  = session['user_id']

    if plan not in PLANS:
        return jsonify({'error': 'باقة غير صحيحة'})

    if PLANS[plan]['price'] == 0:
        db = get_db()
        db.execute("UPDATE users SET plan='free',plan_expires='' WHERE id=?", (uid,))
        db.commit()
        return jsonify({'success': True, 'message': 'تم التحويل للباقة المجانية'})

    # إذا لم يتم إعداد Moyasar بعد
    if not MOYASAR_API_KEY:
        return jsonify({
            'success':     False,
            'coming_soon': True,
            'message':     'بوابة الدفع قيد الإعداد — أضف MOYASAR_API_KEY في إعدادات السيرفر'
        })

    # إنشاء فاتورة Moyasar
    try:
        import urllib.request, base64
        price_halalas = int(PLANS[plan]['price'] * 100)  # تحويل للهللات
        db       = get_db()
        user     = db.execute("SELECT name,phone,email FROM users WHERE id=?", (uid,)).fetchone()

        payload = json.dumps({
            'amount':       price_halalas,
            'currency':     'SAR',
            'description':  f'حسابي — اشتراك {PLANS[plan]["name"]}',
            'callback_url': f'{APP_URL}/api/payment/callback',
            'source': {
                'type':          'creditcard',
                'name':          user['name'],
                'company':       'hesabi',
                'number':        '',
                'cvc':           '',
                'month':         '',
                'year':          ''
            },
            'metadata': {
                'user_id': str(uid),
                'plan':    plan
            }
        }).encode('utf-8')

        # إنشاء رابط الدفع عبر Moyasar
        auth = base64.b64encode(f'{MOYASAR_SECRET_KEY}:'.encode()).decode()
        req  = urllib.request.Request(
            'https://api.moyasar.com/v1/payments',
            data=payload,
            headers={
                'Content-Type':  'application/json',
                'Authorization': f'Basic {auth}'
            }
        )
        res    = urllib.request.urlopen(req, timeout=15)
        result = json.loads(res.read().decode())

        # حفظ معرّف الدفع في قاعدة البيانات
        db.execute("INSERT INTO payments (user_id,plan,amount,ref_id,status) VALUES (?,?,?,?,?)",
                   (uid, plan, PLANS[plan]['price'], result.get('id',''), 'pending'))
        db.commit()

        return jsonify({
            'success':     True,
            'payment_url': result.get('source', {}).get('transaction_url', ''),
            'payment_id':  result.get('id', '')
        })

    except Exception as e:
        return jsonify({'error': f'خطأ في بوابة الدفع: {str(e)}'})


@app.route('/api/payment/callback', methods=['GET', 'POST'])
def payment_callback():
    """
    Moyasar يرسل هنا بعد إتمام الدفع
    نتحقق من الدفع ونحدّث الباقة
    """
    payment_id = request.args.get('id') or (request.get_json() or {}).get('id', '')
    status     = request.args.get('status', '')

    if not payment_id:
        return '<h2 style="font-family:Cairo;direction:rtl;text-align:center;color:red">❌ رقم الدفع غير صحيح</h2>'

    # التحقق من الدفع مع Moyasar
    try:
        import urllib.request, base64
        auth = base64.b64encode(f'{MOYASAR_SECRET_KEY}:'.encode()).decode()
        req  = urllib.request.Request(
            f'https://api.moyasar.com/v1/payments/{payment_id}',
            headers={'Authorization': f'Basic {auth}'}
        )
        res    = urllib.request.urlopen(req, timeout=15)
        result = json.loads(res.read().decode())

        if result.get('status') == 'paid':
            meta    = result.get('metadata', {})
            uid     = int(meta.get('user_id', 0))
            plan    = meta.get('plan', '')

            if uid and plan in PLANS:
                db      = get_db()
                expires = (datetime.now() + timedelta(days=30)).strftime('%Y-%m-%d')
                db.execute("UPDATE users SET plan=?,plan_expires=? WHERE id=?", (plan, expires, uid))
                db.execute("UPDATE payments SET status='paid' WHERE ref_id=?", (payment_id,))
                db.commit()
                add_notif(uid, 'تم تفعيل اشتراكك! 🎉',
                          f'تم تفعيل باقة {PLANS[plan]["name"]} بنجاح. صالحة حتى {expires}')
                return f'''<html><head><meta charset="utf-8">
                <meta http-equiv="refresh" content="3;url=/">
                <style>body{{font-family:Cairo,sans-serif;direction:rtl;text-align:center;background:#0a0f1e;color:#e2e8f0;padding:60px;}}</style></head>
                <body><div style="font-size:60px;margin-bottom:20px;">🎉</div>
                <h2 style="color:#f6c90e;">تم الدفع بنجاح!</h2>
                <p>تم تفعيل باقة {PLANS[plan]["name"]} — جاري التحويل...</p></body></html>'''

    except Exception as e:
        pass

    return f'''<html><head><meta charset="utf-8">
    <meta http-equiv="refresh" content="4;url=/">
    <style>body{{font-family:Cairo,sans-serif;direction:rtl;text-align:center;background:#0a0f1e;color:#e2e8f0;padding:60px;}}</style></head>
    <body><div style="font-size:60px;margin-bottom:20px;">❌</div>
    <h2 style="color:#fc8181;">فشل الدفع</h2>
    <p>يرجى المحاولة مرة أخرى أو التواصل مع الدعم</p></body></html>'''


@app.route('/api/payment/history')
@login_required
def payment_history():
    """سجل المدفوعات"""
    db   = get_db()
    rows = db.execute(
        "SELECT * FROM payments WHERE user_id=? ORDER BY created DESC",
        (session['user_id'],)
    ).fetchall()
    return jsonify({'data': [dict(r) for r in rows]})


@app.route('/plans')
def plans_page():
    """صفحة الباقات المستقلة"""
    uid       = session.get('user_id')
    cur_plan  = 'free'
    user_name = ''
    if uid:
        db   = get_db()
        user = db.execute("SELECT plan,name FROM users WHERE id=?", (uid,)).fetchone()
        if user:
            cur_plan  = user['plan']
            user_name = user['name']

    plans_html = ''
    for pid, p in PLANS.items():
        is_current = pid == cur_plan
        badge = '<div style="position:absolute;top:-12px;right:16px;background:#3182ce;color:#fff;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;">⭐ مميز</div>' if pid=='professional' else ''
        if pid == 'business':
            badge = '<div style="position:absolute;top:-12px;right:16px;background:#f6c90e;color:#000;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;">💎 أعمال</div>'
        feats = {
            'free':         ['50 معاملة شهرياً','10 جهات اتصال','تقارير أساسية','بدون تصدير'],
            'professional': ['معاملات غير محدودة','جهات غير محدودة','تقارير متقدمة','تصدير PDF وExcel','نسخ احتياطي','دعم أولوية'],
            'business':     ['كل مميزات الاحترافية','حتى 5 مستخدمين','فواتير وإيصالات','مزامنة فورية','مدير حساب مخصص'],
        }
        feats_html = ''.join([f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:13px;"><span style="color:#48bb78;font-weight:700;">✓</span>{f}</div>' for f in feats.get(pid,[])])
        if is_current:
            btn = f'<div style="background:#1a2340;color:#a0aec0;border-radius:10px;padding:12px;text-align:center;font-weight:700;font-size:14px;">✅ باقتك الحالية</div>'
        else:
            price_label = "مجاني — ابدأ الآن" if p["price"] == 0 else f'اشترك — {p["price"]} ر.س/شهر'
            btn = f'<button onclick="choosePlan(\'{pid}\')" style="width:100%;padding:13px;background:linear-gradient(135deg,#f6c90e,#d4a017);color:#000;border:none;border-radius:10px;font-size:15px;font-weight:700;font-family:Cairo,sans-serif;cursor:pointer;">{price_label}</button>'

        plans_html += f'''
        <div style="border:{("2px solid #f6c90e" if is_current else "1px solid rgba(99,179,237,0.2)")};border-radius:16px;padding:24px;background:#111827;position:relative;margin-bottom:20px;">
            {badge}
            <div style="font-size:20px;font-weight:800;margin-bottom:8px;">{p["name"]}</div>
            <div style="font-size:32px;font-weight:900;color:#f6c90e;margin-bottom:16px;">{p["price"]}<span style="font-size:14px;color:#a0aec0;font-weight:400;"> ر.س / شهر</span></div>
            {feats_html}
            <div style="margin-top:16px;">{btn}</div>
        </div>'''

    return f'''<!DOCTYPE html><html lang="ar" dir="rtl">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>حسابي — الباقات والأسعار</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
<style>*{{margin:0;padding:0;box-sizing:border-box;}}body{{font-family:Cairo,sans-serif;background:#0a0f1e;color:#e2e8f0;min-height:100vh;}}</style>
</head><body>
<div style="background:#111827;border-bottom:1px solid rgba(246,201,14,0.3);padding:16px 20px;display:flex;align-items:center;justify-content:space-between;">
    <div style="display:flex;align-items:center;gap:10px;">
        <div style="width:38px;height:38px;background:linear-gradient(135deg,#f6c90e,#d4a017);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;">💰</div>
        <span style="font-size:18px;font-weight:900;color:#f6c90e;">حسابي</span>
    </div>
    <a href="/" style="color:#a0aec0;text-decoration:none;font-size:13px;">← العودة للتطبيق</a>
</div>
<div style="max-width:480px;margin:0 auto;padding:30px 20px;">
    <div style="text-align:center;margin-bottom:32px;">
        <div style="font-size:40px;margin-bottom:12px;">💎</div>
        <h1 style="font-size:28px;font-weight:900;color:#f6c90e;margin-bottom:8px;">الباقات والأسعار</h1>
        <p style="color:#a0aec0;font-size:14px;">اختر الباقة المناسبة لك</p>
        {f'<p style="color:#48bb78;font-size:13px;margin-top:8px;">مرحباً {user_name} 👋</p>' if user_name else ''}
    </div>
    {plans_html}
    <div style="text-align:center;margin-top:24px;padding:20px;background:#111827;border-radius:16px;border:1px solid rgba(99,179,237,0.15);">
        <div style="font-size:24px;margin-bottom:10px;">👨‍💻</div>
        <div style="font-size:15px;font-weight:700;margin-bottom:4px;">هل تحتاج مساعدة؟</div>
        <div style="font-size:13px;color:#a0aec0;margin-bottom:14px;">تواصل مع المطوّر مباشرة</div>
        <a href="https://wa.me/966567867414" target="_blank"
           style="display:inline-block;padding:10px 24px;background:rgba(72,187,120,0.15);color:#48bb78;border:1px solid rgba(72,187,120,0.3);border-radius:10px;text-decoration:none;font-weight:700;font-size:14px;">💬 واتساب</a>
    </div>
</div>
<script>
function choosePlan(plan) {{
    if (!{str(bool(uid)).lower()}) {{
        window.location.href = '/';
        return;
    }}
    if (plan === 'free') {{
        if (confirm('هل تريد التحويل للباقة المجانية؟')) {{
            fetch('/api/subscribe', {{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{plan:'free'}})}})
            .then(r=>r.json()).then(d=>{{ if(d.success) {{ alert('تم!'); location.reload(); }} else alert(d.message||d.error); }});
        }}
        return;
    }}
    const prices = {{professional:29,business:79}};
    if (confirm(`الاشتراك في الباقة\nالسعر: ${{prices[plan]}} ر.س / شهر\n\nهل تريد المتابعة للدفع؟`)) {{
        fetch('/api/subscribe', {{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{plan:plan}})}})
        .then(r=>r.json())
        .then(d=>{{
            if (d.payment_url) {{
                window.location.href = d.payment_url;
            }} else if (d.coming_soon) {{
                alert('بوابة الدفع قيد الإعداد\\nسيتم تفعيلها قريباً\\n\\nللتفعيل اليدوي تواصل: 0567867414');
            }} else {{
                alert(d.error || d.message || 'خطأ');
            }}
        }});
    }}
}}
</script>
</body></html>'''



# ============================================================
# تشغيل التطبيق
# ============================================================
if __name__ == '__main__':
    print("=" * 55)
    print("  💰 حسابي SaaS — نسخة الإنتاج")
    print("=" * 55)
    init_db()
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'true').lower() == 'true'
    print(f"🌐 http://localhost:{port}")
    print(f"📁 قاعدة البيانات: {DB_PATH}")
    print(f"📱 SMS Provider: {SMS_PROVIDER}")
    print("🛑 Ctrl+C للإيقاف")
    print("=" * 55)
    app.run(debug=debug, host='0.0.0.0', port=port)
