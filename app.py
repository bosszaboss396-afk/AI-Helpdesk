from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
import pymysql
import pymysql.cursors
from datetime import datetime
import openpyxl
import os
from io import BytesIO
from dotenv import load_dotenv

# นำเข้า Google GenAI SDK
from google import genai
from google.genai import types

from ai_service import analyze_issue_with_ai
from line_service import send_line_notification
from excel_service import export_tickets_to_excel

app = Flask(__name__)
app.secret_key = 'court_phetchabun_secret_key'

# ----------------------------------------------------
# ⚙️ การตั้งค่า Database MySQL (phpMyAdmin)
# ----------------------------------------------------
DB_HOST = 'localhost'
DB_USER = 'root'
DB_PASS = ''
DB_NAME = 'ai_helpdesk_db'

# ----------------------------------------------------
# 🤖 การตั้งค่า Google Gemini AI API
# ----------------------------------------------------
load_dotenv(override=True)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

print(f"🔥 Current Gemini API Key in use: {GEMINI_API_KEY[:10] if GEMINI_API_KEY else 'None'}...")

# เรียกใช้ Client ของ Google GenAI
client = genai.Client(api_key=GEMINI_API_KEY)


def get_db():
    return pymysql.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        cursorclass=pymysql.cursors.DictCursor
    )


# ฟังก์ชันสร้างตารางและข้อมูลเริ่มต้นอัตโนมัติ (MySQL Version)
def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # 1. ตาราง Master Data "แผนก"
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS departments (
            id INTEGER PRIMARY KEY AUTO_INCREMENT,
            dept_name VARCHAR(255) UNIQUE NOT NULL
        )
    ''')
    
    # 2. ตาราง Master Data "ห้อง/สถานที่"
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS rooms (
            id INTEGER PRIMARY KEY AUTO_INCREMENT,
            room_name VARCHAR(255) UNIQUE NOT NULL
        )
    ''')
    
    # 3. ตาราง Master Data "ชนิด/หมวดหมู่อุปกรณ์"
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTO_INCREMENT,
            category_name VARCHAR(255) UNIQUE NOT NULL
        )
    ''')
    
    # 4. ตารางผู้ใช้งาน
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id VARCHAR(50) PRIMARY KEY,
            password VARCHAR(255) NOT NULL,
            fullname VARCHAR(255) NOT NULL,
            department VARCHAR(255),
            position VARCHAR(255),
            room VARCHAR(255),
            role VARCHAR(50) DEFAULT 'user'
        )
    ''')
    
    # 5. ตารางครุภัณฑ์
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS assets (
            asset_id VARCHAR(100) PRIMARY KEY,
            asset_name VARCHAR(255) NOT NULL,
            category VARCHAR(255),
            warranty VARCHAR(100),
            serial_number VARCHAR(100),
            room VARCHAR(255),
            status VARCHAR(50) DEFAULT 'ปกติ',
            assigned_user_id VARCHAR(50)
        )
    ''')
    
    # 6. ตารางแจ้งซ่อม (Tickets) - เปลี่ยนชนิดภาพเป็น LONGTEXT
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tickets (
            ticket_id VARCHAR(50) PRIMARY KEY,
            created_at DATETIME,
            user_id VARCHAR(50),
            asset_id VARCHAR(100),
            issue_description TEXT,
            image_data LONGTEXT,
            status VARCHAR(50) DEFAULT 'pending',
            assigned_admin VARCHAR(255),
            repair_note TEXT
        )
    ''')
    
    # 7. ตารางประวัติการซ่อม (Repair Logs)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS repair_logs (
            log_id INTEGER PRIMARY KEY AUTO_INCREMENT,
            ticket_id VARCHAR(50),
            asset_id VARCHAR(100),
            repaired_at DATETIME,
            issue_summary TEXT,
            action_taken TEXT,
            ai_resolved INTEGER
        )
    ''')

    # 8. ตารางคู่มือแก้ปัญหาไอทีภายใน (Knowledge Base)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS knowledge_base (
            id INTEGER PRIMARY KEY AUTO_INCREMENT,
            keyword VARCHAR(255) NOT NULL,
            issue_title VARCHAR(255) NOT NULL,
            solution_steps TEXT NOT NULL
        )
    ''')

    # 9. ตารางเก็บค่าตั้งค่าระบบ (App Settings)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS app_settings (
            `key` VARCHAR(100) PRIMARY KEY,
            `value` TEXT
        )
    ''')
    cursor.execute("INSERT IGNORE INTO app_settings (`key`, `value`) VALUES ('google_form_url', 'https://forms.gle/your-default-form')")

    # เพิ่มข้อมูลตั้งต้น (Default Data)
    cursor.execute("SELECT COUNT(*) AS count FROM departments")
    if cursor.fetchone()['count'] == 0:
        cursor.executemany("INSERT INTO departments (dept_name) VALUES (%s)", [
            ('ผู้อำนวยการ',), ('ผู้พิพากษา',), ('ศูนย์เทคโนโลยีสารสนเทศ',),
            ('ส่วนช่วยอำนวยการ',), ('ส่วนคลัง',), ('ส่วนบริหารจัดการคดี',),
            ('ส่วนช่วยพิจารณาคดี',), ('ส่วนเจ้าพนักงานคดี',), ('ส่วนบริการประชาชนและประชาสัมพันธ์',),
            ('ส่วนไกล่เกลี่ยและประนอมข้อพิพาท',), ('ส่วนเจ้าพนักงานตำรวจศาล',), ('อื่นๆ',)
        ])
        
    cursor.execute("SELECT COUNT(*) AS count FROM rooms")
    if cursor.fetchone()['count'] == 0:
        cursor.executemany("INSERT INTO rooms (room_name) VALUES (%s)", [
            ('ห้องธุรการขนาดใหญ่',), ('ห้องพักผู้พิพากษา',), ('ศูนย์นัดความ',),
            ('ห้องเก็บสำนวน',), ('ห้องเจ้าพนักงานคดี',), ('ห้องไกล่เกลี่ย',),
            ('ห้องศูนย์ server',), ('ห้อง CCTV',), ('ห้องพิจารณาคดี 1',),
            ('ห้องพิจารณาคดี 2',), ('ห้องพิจารณาคดี 3',), ('ห้องพิจารณาคดี 4',),
            ('ห้องพิจารณาคดี 5',), ('ห้องพิจารณาคดี 6',), ('อื่นๆ',)
        ])

    cursor.execute("SELECT COUNT(*) AS count FROM categories")
    if cursor.fetchone()['count'] == 0:
        cursor.executemany("INSERT INTO categories (category_name) VALUES (%s)", [
            ('คอมพิวเตอร์ pc',), ('อุปกรณ์พกพา',), ('อุปกรณ์ต่อพ่วง',),
            ('เครื่องพิมพ์และสแกนเอกสาร',), ('อุปกรณ์สำรองไฟ',), ('อุปกรณ์เครือข่าย',),
            ('อุปกรณ์รักษาความปลอดภัย',), ('อุปกรณ์สื่อสาร',), ('อุปกรณ์เฉพาะทาง',), ('อื่นๆ',)
        ])

    cursor.execute("SELECT COUNT(*) AS count FROM knowledge_base")
    if cursor.fetchone()['count'] == 0:
        cursor.executemany("INSERT INTO knowledge_base (keyword, issue_title, solution_steps) VALUES (%s, %s, %s)", [
            ('word', 'ใช้งาน Word ค้าง เม้าส์ขยับไม่ได้', '1. กดปุ่ม Power ที่เครื่องค้างไว้ประมาณ 5-10 วินาที เพื่อรีสตาร์ทเครื่อง\n2. ตรวจสอบและเคลียร์ไฟล์ Temp ในระบบ\n3. หากยังไม่หาย ให้แจ้งช่างเพื่อตรวจสอบโปรแกรม Office'),
            ('printer', 'เครื่องพิมพ์ไม่ออก / ปริ้นไม่ได้', '1. ตรวจสอบการเชื่อมต่อสาย USB หรือสายแลนที่ตัวเครื่องพิมพ์\n2. ไปที่ Services.msc แล้วทำการ Restart บริการ Print Spooler\n3. ลบงานค้างในคิวการพิมพ์ (Print Queue) แล้วลองพิมพ์ใหม่อีกครั้ง'),
            ('blue screen', 'จอฟ้า (Blue Screen / BSOD)', '1. จดรหัส Stop Code ที่ปรากฏบนหน้าจอฟ้า\n2. ถอดอุปกรณ์ต่อพ่วงที่ไม่จำเป็นออก (เช่น Flash Drive, External HDD)\n3. ทำการรีสตาร์ทเครื่อง หากยังเกิดซ้ำให้แจ้งช่างตรวจเช็คฮาร์ดดิสก์'),
            ('internet', 'อินเทอร์เน็ตหลุด / เข้าเว็บศาลไม่ได้', '1. ตรวจสอบสาย LAN ด้านหลังคอมพิวเตอร์ว่าไฟสถานะกะพริบปกติหรือไม่\n2. ทำการ Disconnect และ Connect สัญญาณเครือข่ายใหม่\n3. ตรวจสอบการตั้งค่า IP Address ภายในวงแลนของศาล'),
            ('scanner', 'เครื่องสแกนเอกสารไม่ทำงาน / สแกนไม่ได้', '1. ตรวจสอบสาย USB ที่เชื่อมต่อระหว่างเครื่องสแกนกับคอมพิวเตอร์\n2. ตรวจสอบสถานะไฟเลี้ยงที่ตัวเครื่องสแกนเอกสาร\n3. เปิดโปรแกรม Scanner Utility หรือ Driver ขึ้นมาใหม่เพื่อรีเฟรชการเชื่อมต่อ'),
            ('password', 'ลืมรหัสผ่านเข้าระบบ / เข้าใช้งานไม่ได้', '1. ตรวจสอบการพิมพ์ภาษาไทย/ภาษาอังกฤษ (Caps Lock) ให้ถูกต้อง\n2. ติดต่อผู้ดูแลระบบ (Admin) เพื่อรีเซ็ตโครงสร้างรหัสผ่านชั่วคราวผ่านหน้าจัดการผู้ใช้งาน')
        ])

    conn.commit()
    conn.close()


# เรียกใช้ init_db ตอนเริ่มแอป
try:
    init_db()
    print("✅ MySQL Database Initialized Successfully!")
except Exception as e:
    print(f"⚠️ MySQL Error: {e}")


@app.route('/')
def index():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT `value` FROM app_settings WHERE `key` = 'google_form_url'")
    setting = cursor.fetchone()
    conn.close()
    google_form_url = setting['value'] if setting else ''
    return render_template('login.html', google_form_url=google_form_url)


@app.route('/login', methods=['POST'])
def login():
    user_id = request.form['user_id']
    password = request.form['password']

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = %s AND password = %s", (user_id, password))
    user = cursor.fetchone()
    conn.close()

    if user:
        session['user_id'] = user['user_id']
        session['fullname'] = user['fullname']
        session['department'] = user['department']
        session['role'] = user['role']
        
        if user['role'] == 'admin':
            return redirect(url_for('admin_dashboard'))
        return redirect(url_for('chat'))
    else:
        return "<script>alert('รหัสบุคลากรหรือรหัสผ่านไม่ถูกต้อง'); window.location='/';</script>"


@app.route('/quick-check', methods=['POST'])
def quick_check():
    ticket_id = request.form['ticket_id'].strip()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT t.*, a.asset_name, u.fullname 
        FROM tickets t
        JOIN assets a ON t.asset_id = a.asset_id
        JOIN users u ON t.user_id = u.user_id
        WHERE t.ticket_id = %s
    """, (ticket_id,))
    ticket = cursor.fetchone()
    conn.close()

    if ticket:
        status_map = {
            'pending': '⏳ รอรับเคส',
            'in_progress': '🛠️ กำลังดำเนินการ',
            'completed': '✅ ซ่อมสำเร็จ',
            'failed': '❌ ไม่สามารถซ่อมได้'
        }
        status_text = status_map.get(ticket['status'], ticket['status'])
        note = ticket['repair_note'] if ticket['repair_note'] else '-'
        msg = f"สถานะเคส {ticket['ticket_id']}: {status_text}\\nอุปกรณ์: {ticket['asset_name']}\\nวิธีแก้ไข: {note}"
        return f"<script>alert('{msg}'); window.location='/';</script>"
        
    return "<script>alert('ไม่พบข้อมูล Ticket นี้ในระบบ'); window.location='/';</script>"


@app.route('/chat')
def chat():
    if 'user_id' not in session:
        return redirect(url_for('index'))

    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT room FROM users WHERE user_id = %s", (user_id,))
    user_info = cursor.fetchone()
    user_room = user_info['room'] if user_info and user_info['room'] else ''

    cursor.execute("""
        SELECT * FROM assets 
        WHERE assigned_user_id = %s 
            OR ((assigned_user_id IS NULL OR assigned_user_id = '') AND room = %s)
    """, (user_id, user_room))
    assets = cursor.fetchall()
    conn.close()

    user = {'fullname': session['fullname'], 'department': session['department']}
    return render_template('user_chat.html', user=user, assets=assets)


@app.route('/api/ai-analyze', methods=['POST'])
def api_ai_analyze():
    data = request.json
    analysis = analyze_issue_with_ai(data['asset_id'], data['issue'])
    return jsonify({'analysis': analysis})


@app.route('/api/create-ticket', methods=['POST'])
def create_ticket():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) AS count FROM tickets")
    count = cursor.fetchone()['count']
    ticket_id = f"TK-{1001 + count}"
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    image_data = data.get('image')

    cursor.execute("""
        INSERT INTO tickets (ticket_id, created_at, user_id, asset_id, issue_description, image_data, status)
        VALUES (%s, %s, %s, %s, %s, %s, 'pending')
    """, (ticket_id, created_at, session['user_id'], data['asset_id'], data['issue'], image_data))
    
    cursor.execute("""
        SELECT a.asset_name, a.room, u.fullname, u.department
        FROM assets a, users u
        WHERE a.asset_id = %s AND u.user_id = %s
    """, (data['asset_id'], session['user_id']))
    ticket_info = cursor.fetchone()

    conn.commit()
    conn.close()

    if ticket_info:
        send_line_notification(
            ticket_id=ticket_id,
            user_name=ticket_info['fullname'],
            department=ticket_info['department'],
            asset_name=ticket_info['asset_name'],
            room=ticket_info['room'],
            issue=data['issue']
        )

    return jsonify({'ticket_id': ticket_id})


@app.route('/admin')
def admin_dashboard():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('index'))

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    search_text = request.args.get('search_text', '').strip()

    conn = get_db()
    cursor = conn.cursor()
    
    base_query = """
        SELECT t.*, a.asset_name, a.category, a.room, u.fullname, u.department
        FROM tickets t
        JOIN assets a ON t.asset_id = a.asset_id
        JOIN users u ON t.user_id = u.user_id
        WHERE 1=1
    """
    params = []
    
    if start_date and end_date:
        base_query += " AND DATE(t.created_at) BETWEEN %s AND %s"
        params.extend([start_date, end_date])
        
    if search_text:
        base_query += " AND (t.asset_id LIKE %s OR a.asset_name LIKE %s OR t.ticket_id LIKE %s)"
        like_term = f"%{search_text}%"
        params.extend([like_term, like_term, like_term])
        
    query = base_query + " ORDER BY t.created_at DESC"
    cursor.execute(query, params)
    tickets = cursor.fetchall()

    pending_count = sum(1 for t in tickets if t['status'] == 'pending')
    in_progress_count = sum(1 for t in tickets if t['status'] == 'in_progress')
    completed_count = sum(1 for t in tickets if t['status'] == 'completed')

    conn.close()
    
    return render_template('admin.html', 
                           tickets=tickets, 
                           pending=pending_count, 
                           in_progress=in_progress_count, 
                           completed=completed_count,
                           start_date=start_date,
                           end_date=end_date,
                           search_text=search_text)


@app.route('/admin/print-report')
def admin_print_report():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    search_text = request.args.get('search_text', '').strip()

    conn = get_db()
    cursor = conn.cursor()
    
    base_query = """
        SELECT t.*, a.asset_name, a.category, a.room, u.fullname, u.department
        FROM tickets t
        JOIN assets a ON t.asset_id = a.asset_id
        JOIN users u ON t.user_id = u.user_id
        WHERE 1=1
    """
    params = []
    
    if start_date and end_date:
        base_query += " AND DATE(t.created_at) BETWEEN %s AND %s"
        params.extend([start_date, end_date])
        
    if search_text:
        base_query += " AND (t.asset_id LIKE %s OR a.asset_name LIKE %s OR t.ticket_id LIKE %s)"
        like_term = f"%{search_text}%"
        params.extend([like_term, like_term, like_term])
        
    query = base_query + " ORDER BY t.created_at DESC"
    cursor.execute(query, params)
    tickets = cursor.fetchall()
    conn.close()

    return render_template('print_report.html', 
                           tickets=tickets, 
                           start_date=start_date, 
                           end_date=end_date,
                           search_text=search_text)


@app.route('/api/update-ticket', methods=['POST'])
def update_ticket():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    ticket_id = data['ticket_id']
    status = data['status']
    assigned_admin = data['assigned_admin']
    repair_note = data['repair_note']

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tickets 
        SET status = %s, assigned_admin = %s, repair_note = %s
        WHERE ticket_id = %s
    """, (status, assigned_admin, repair_note, ticket_id))

    if status == 'completed':
        cursor.execute("SELECT asset_id, issue_description FROM tickets WHERE ticket_id = %s", (ticket_id,))
        ticket = cursor.fetchone()
        repaired_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute("SELECT log_id FROM repair_logs WHERE ticket_id = %s", (ticket_id,))
        existing_log = cursor.fetchone()
        if not existing_log:
            cursor.execute("""
                INSERT INTO repair_logs (ticket_id, asset_id, repaired_at, issue_summary, action_taken, ai_resolved)
                VALUES (%s, %s, %s, %s, %s, 0)
            """, (ticket_id, ticket['asset_id'], repaired_at, ticket['issue_description'], repair_note))

    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/admin/export-excel')
def export_excel():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('index'))

    excel_file = export_tickets_to_excel()
    return send_file(
        excel_file,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Helpdesk_Report_Phetchabun.xlsx"
    )


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))


@app.route('/api/asset-history/<path:asset_id>')
def get_asset_history(asset_id):
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ticket_id, created_at, issue_description, repair_note, assigned_admin, status
        FROM tickets
        WHERE asset_id = %s
        ORDER BY created_at DESC
    """, (asset_id,))
    history = cursor.fetchall()
    conn.close()

    total_reported = len(history)
    total_completed = sum(1 for log in history if log['status'] == 'completed')

    return jsonify({
        'history': history,
        'total_reported': total_reported,
        'total_completed': total_completed
    })


@app.route('/admin/users')
def admin_users():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users ORDER BY user_id")
    users = cursor.fetchall()
    
    cursor.execute("SELECT dept_name FROM departments ORDER BY CASE WHEN dept_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    departments = cursor.fetchall()
    cursor.execute("SELECT room_name FROM rooms ORDER BY CASE WHEN room_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    rooms = cursor.fetchall()
    conn.close()
    
    return render_template('admin_users.html', 
                           users=users,
                           departments=[d['dept_name'] for d in departments],
                           rooms=[r['room_name'] for r in rooms])


@app.route('/api/users/add', methods=['POST'])
def add_user():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO users (user_id, password, fullname, department, position, room, role) 
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (data['user_id'], data['password'], data['fullname'], data['department'], data['position'], data['room'], data['role']))
        conn.commit()
        success = True
        msg = "เพิ่มผู้ใช้งานสำเร็จ"
    except pymysql.err.IntegrityError:
        success = False
        msg = "รหัสบุคลากรนี้มีอยู่ในระบบแล้ว"
    finally:
        conn.close()
    
    return jsonify({'success': success, 'message': msg})


@app.route('/api/users/delete', methods=['POST'])
def delete_user():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM users WHERE user_id = %s", (data['user_id'],))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'ลบผู้ใช้งานเรียบร้อยแล้ว'})


@app.route('/api/users/edit', methods=['POST'])
def edit_user():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    try:
        if data.get('password') and data['password'].strip() != '':
            cursor.execute("""
                UPDATE users 
                SET fullname = %s, department = %s, position = %s, room = %s, role = %s, password = %s
                WHERE user_id = %s
            """, (data['fullname'], data['department'], data['position'], data['room'], data['role'], data['password'], data['user_id']))
        else:
            cursor.execute("""
                UPDATE users 
                SET fullname = %s, department = %s, position = %s, room = %s, role = %s
                WHERE user_id = %s
            """, (data['fullname'], data['department'], data['position'], data['room'], data['role'], data['user_id']))
            
        conn.commit()
        success = True
        msg = "แก้ไขข้อมูลสำเร็จ"
    except Exception as e:
        success = False
        msg = f"เกิดข้อผิดพลาด: {str(e)}"
    finally:
        conn.close()
        
    return jsonify({'success': success, 'message': msg})


@app.route('/api/users/import', methods=['POST'])
def import_users():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'ไม่ได้แนบไฟล์'})
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'กรุณาเลือกไฟล์ Excel'})
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            wb = openpyxl.load_workbook(file)
            sheet = wb.active
            conn = get_db()
            cursor = conn.cursor()
            count = 0
            
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1] and row[2]: 
                    user_id = str(row[0]).strip()
                    password = str(row[1]).strip()
                    fullname = str(row[2]).strip()
                    position = str(row[3]).strip() if len(row) > 3 and row[3] else '-'
                    department = str(row[4]).strip() if len(row) > 4 and row[4] else 'ไม่ระบุแผนก'
                    room = str(row[5]).strip() if len(row) > 5 and row[5] else 'ไม่ระบุห้อง'
                    
                    role_val = str(row[6]).strip().lower() if len(row) > 6 and row[6] else 'user'
                    role = 'admin' if role_val == 'admin' else 'user'
                    
                    cursor.execute("""
                        INSERT INTO users (user_id, password, fullname, position, department, room, role) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        ON DUPLICATE KEY UPDATE
                            password=VALUES(password),
                            fullname=VALUES(fullname),
                            position=VALUES(position),
                            department=VALUES(department),
                            room=VALUES(room),
                            role=VALUES(role)
                    """, (user_id, password, fullname, position, department, room, role))
                    count += 1
            
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': f'นำเข้า/อัปเดตข้อมูลผู้ใช้งานสำเร็จจำนวน {count} รายการ'})
            
        except Exception as e:
            return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'})
            
    return jsonify({'success': False, 'message': 'รองรับเฉพาะไฟล์ .xlsx หรือ .xls เท่านั้น'})


@app.route('/admin/assets')
def admin_assets():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM assets ORDER BY asset_id")
    assets = cursor.fetchall()
    
    cursor.execute("SELECT category_name FROM categories ORDER BY CASE WHEN category_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    categories = cursor.fetchall()
    cursor.execute("SELECT room_name FROM rooms ORDER BY CASE WHEN room_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    rooms = cursor.fetchall()
    cursor.execute("SELECT user_id, fullname, position FROM users ORDER BY user_id")
    users = cursor.fetchall()
    conn.close()
    
    return render_template('admin_assets.html', 
                           assets=assets,
                           categories=[c['category_name'] for c in categories],
                           rooms=[r['room_name'] for r in rooms],
                           users=users)


@app.route('/api/assets/add', methods=['POST'])
def add_asset():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO assets (asset_id, asset_name, category, warranty, serial_number, room, status, assigned_user_id) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            data['asset_id'], 
            data['asset_name'], 
            data['category'], 
            data['warranty'], 
            data['serial_number'], 
            data['room'], 
            data.get('status', 'ปกติ'),
            data.get('assigned_user_id', '')
        ))
        conn.commit()
        success = True
        msg = "เพิ่มข้อมูลครุภัณฑ์สำเร็จ"
    except pymysql.err.IntegrityError:
        success = False
        msg = "หมายเลขครุภัณฑ์นี้มีอยู่ในระบบแล้ว"
    finally:
        conn.close()
    
    return jsonify({'success': success, 'message': msg})


@app.route('/api/assets/delete', methods=['POST'])
def delete_asset():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    asset_id = data.get('asset_id')
    
    if not asset_id:
        return jsonify({'success': False, 'message': 'ไม่พบรหัสครุภัณฑ์ที่ต้องการลบ'})

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM assets WHERE asset_id = %s", (asset_id,))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'ลบข้อมูลครุภัณฑ์เรียบร้อยแล้ว'})


@app.route('/api/assets/edit', methods=['POST'])
def edit_asset():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE assets 
            SET asset_name = %s, category = %s, warranty = %s, serial_number = %s, room = %s, status = %s, assigned_user_id = %s
            WHERE asset_id = %s
        """, (
            data['asset_name'], 
            data['category'], 
            data['warranty'], 
            data['serial_number'], 
            data['room'], 
            data['status'],
            data.get('assigned_user_id', ''),
            data['asset_id']
        ))
        conn.commit()
        success = True
        msg = "แก้ไขข้อมูลครุภัณฑ์สำเร็จ"
    except Exception as e:
        success = False
        msg = f"เกิดข้อผิดพลาด: {str(e)}"
    finally:
        conn.close()
    
    return jsonify({'success': success, 'message': msg})

@app.route('/api/assets/import', methods=['POST'])
def import_assets():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'ไม่ได้แนบไฟล์'})
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'กรุณาเลือกไฟล์ Excel'})
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            wb = openpyxl.load_workbook(file)
            sheet = wb.active
            conn = get_db()
            cursor = conn.cursor()
            count = 0
            
            # เริ่มอ่านข้อมูลจากแถวที่ 2 (ข้ามแถวหัวข้อ)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]: # อย่างน้อยต้องมีรหัสครุภัณฑ์และชื่อครุภัณฑ์
                    asset_id = str(row[0]).strip()
                    asset_name = str(row[1]).strip()
                    category = str(row[2]).strip() if len(row) > 2 and row[2] else 'อื่นๆ'
                    warranty = str(row[3]).strip() if len(row) > 3 and row[3] else '-'
                    serial_number = str(row[4]).strip() if len(row) > 4 and row[4] else '-'
                    room = str(row[5]).strip() if len(row) > 5 and row[5] else 'ไม่ระบุห้อง'
                    status = str(row[6]).strip() if len(row) > 6 and row[6] else 'ปกติ'
                    
                    # รหัสผู้ถือครอง (ถ้าไม่มี ให้เป็นค่าว่าง)
                    assigned_user_id = str(row[7]).strip() if len(row) > 7 and row[7] else ''
                    if assigned_user_id in ['None', '-']:
                        assigned_user_id = ''
                    
                    cursor.execute("""
                        INSERT INTO assets (asset_id, asset_name, category, warranty, serial_number, room, status, assigned_user_id) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        ON DUPLICATE KEY UPDATE
                            asset_name=VALUES(asset_name),
                            category=VALUES(category),
                            warranty=VALUES(warranty),
                            serial_number=VALUES(serial_number),
                            room=VALUES(room),
                            status=VALUES(status),
                            assigned_user_id=VALUES(assigned_user_id)
                    """, (asset_id, asset_name, category, warranty, serial_number, room, status, assigned_user_id))
                    count += 1
            
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': f'นำเข้า/อัปเดตข้อมูลครุภัณฑ์สำเร็จจำนวน {count} รายการ'})
            
        except Exception as e:
            return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'})
            
    return jsonify({'success': False, 'message': 'รองรับเฉพาะไฟล์ .xlsx หรือ .xls เท่านั้น'})

@app.route('/admin/settings')
def admin_settings():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM departments ORDER BY CASE WHEN dept_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    departments = cursor.fetchall()
    cursor.execute("SELECT * FROM rooms ORDER BY CASE WHEN room_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    rooms = cursor.fetchall()
    cursor.execute("SELECT * FROM categories ORDER BY CASE WHEN category_name = 'อื่นๆ' THEN 1 ELSE 0 END, id ASC")
    categories = cursor.fetchall()
    
    cursor.execute("SELECT `value` FROM app_settings WHERE `key` = 'google_form_url'")
    setting = cursor.fetchone()
    google_form_url = setting['value'] if setting else ''
    
    conn.close()
    
    return render_template('admin_settings.html', 
                           departments=departments, 
                           rooms=rooms,
                           categories=categories,
                           google_form_url=google_form_url)


@app.route('/api/settings/add', methods=['POST'])
def add_setting():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    item_type = data.get('type')
    name = data.get('name', '').strip()
    
    if not name:
        return jsonify({'success': False, 'message': 'กรุณากรอกชื่อที่ต้องการเพิ่ม'})

    conn = get_db()
    cursor = conn.cursor()
    try:
        if item_type == 'department':
            cursor.execute("INSERT INTO departments (dept_name) VALUES (%s)", (name,))
        elif item_type == 'room':
            cursor.execute("INSERT INTO rooms (room_name) VALUES (%s)", (name,))
        elif item_type == 'category':
            cursor.execute("INSERT INTO categories (category_name) VALUES (%s)", (name,))
        conn.commit()
        success = True
        msg = "เพิ่มข้อมูลสำเร็จ"
    except pymysql.err.IntegrityError:
        success = False
        msg = "ชื่อนี้มีอยู่ในระบบแล้ว"
    finally:
        conn.close()
        
    return jsonify({'success': success, 'message': msg})


@app.route('/api/settings/delete', methods=['POST'])
def delete_setting():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    item_type = data.get('type')
    item_id = data.get('id')
    
    conn = get_db()
    cursor = conn.cursor()
    if item_type == 'department':
        cursor.execute("DELETE FROM departments WHERE id = %s", (item_id,))
    elif item_type == 'room':
        cursor.execute("DELETE FROM rooms WHERE id = %s", (item_id,))
    elif item_type == 'category':
        cursor.execute("DELETE FROM categories WHERE id = %s", (item_id,))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'ลบข้อมูลเรียบร้อยแล้ว'})


@app.route('/api/settings/edit', methods=['POST'])
def edit_setting():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    item_type = data.get('type')
    item_id = data.get('id')
    new_name = data.get('name', '').strip()
    
    if not new_name:
        return jsonify({'success': False, 'message': 'กรุณากรอกชื่อที่ต้องการแก้ไข'})

    conn = get_db()
    cursor = conn.cursor()
    try:
        if item_type == 'department':
            cursor.execute("UPDATE departments SET dept_name = %s WHERE id = %s", (new_name, item_id))
        elif item_type == 'room':
            cursor.execute("UPDATE rooms SET room_name = %s WHERE id = %s", (new_name, item_id))
        elif item_type == 'category':
            cursor.execute("UPDATE categories SET category_name = %s WHERE id = %s", (new_name, item_id))
        conn.commit()
        success = True
        msg = "แก้ไขข้อมูลเรียบร้อยแล้ว"
    except pymysql.err.IntegrityError:
        success = False
        msg = "ชื่อนี้มีอยู่ในระบบแล้ว"
    except Exception as e:
        success = False
        msg = f"เกิดข้อผิดพลาด: {str(e)}"
    finally:
        conn.close()
        
    return jsonify({'success': success, 'message': msg})


@app.route('/api/users/bulk-delete', methods=['POST'])
def bulk_delete_users():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    user_ids = data.get('user_ids', [])
    if not user_ids:
        return jsonify({'success': False, 'message': 'ไม่ได้เลือกรายการที่ต้องการลบ'})

    conn = get_db()
    cursor = conn.cursor()
    placeholders = ','.join(['%s'] * len(user_ids))
    cursor.execute(f"DELETE FROM users WHERE user_id IN ({placeholders})", user_ids)
    conn.commit()
    deleted_count = cursor.rowcount
    conn.close()

    return jsonify({'success': True, 'message': f'ลบผู้ใช้งานสำเร็จจำนวน {deleted_count} รายการ'})


@app.route('/api/assets/bulk-delete', methods=['POST'])
def bulk_delete_assets():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    asset_ids = data.get('asset_ids', [])
    if not asset_ids:
        return jsonify({'success': False, 'message': 'ไม่ได้เลือกรายการที่ต้องการลบ'})

    conn = get_db()
    cursor = conn.cursor()
    placeholders = ','.join(['%s'] * len(asset_ids))
    cursor.execute(f"DELETE FROM assets WHERE asset_id IN ({placeholders})", asset_ids)
    conn.commit()
    deleted_count = cursor.rowcount
    conn.close()

    return jsonify({'success': True, 'message': f'ลบครุภัณฑ์สำเร็จจำนวน {deleted_count} รายการ'})


@app.route('/admin/knowledge-base')
def admin_kb():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM knowledge_base ORDER BY id DESC")
    kb_list = cursor.fetchall()
    conn.close()
    
    return render_template('admin_kb.html', kb_list=kb_list)


@app.route('/api/kb/add', methods=['POST'])
def add_kb():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    keyword = data.get('keyword', '').strip().lower()
    issue_title = data.get('issue_title', '').strip()
    solution_steps = data.get('solution_steps', '').strip()
    
    if not keyword or not issue_title or not solution_steps:
        return jsonify({'success': False, 'message': 'กรุณากรอกข้อมูลให้ครบถ้วน'})

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO knowledge_base (keyword, issue_title, solution_steps) 
            VALUES (%s, %s, %s)
        """, (keyword, issue_title, solution_steps))
        conn.commit()
        success = True
        msg = "เพิ่มคู่มือการแก้ไขสำเร็จ"
    except Exception as e:
        success = False
        msg = f"เกิดข้อผิดพลาด: {str(e)}"
    finally:
        conn.close()
        
    return jsonify({'success': success, 'message': msg})


@app.route('/api/kb/delete', methods=['POST'])
def delete_kb():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    kb_id = data.get('id')
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM knowledge_base WHERE id = %s", (kb_id,))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'ลบคู่มือเรียบร้อยแล้ว'})


@app.route('/api/chat/send', methods=['POST'])
def chat_with_ai():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    asset_id = data.get('asset_id')
    user_message = data.get('message')
    image_base64 = data.get('image')
    history = data.get('history', [])

    if not user_message and not image_base64:
        return jsonify({'success': False, 'message': 'กรุณากรอกข้อความหรือแนบรูปภาพ'}), 400

    asset_context = ""
    if asset_id:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM assets WHERE asset_id = %s", (asset_id,))
        asset = cursor.fetchone()
        conn.close()
        if asset:
            asset_context = f"""
[ข้อมูลครุภัณฑ์ที่ผู้ใช้เลือก]
- รหัสครุภัณฑ์: {asset['asset_id']}
- ชื่ออุปกรณ์/สเปก: {asset['asset_name']}
- ชนิดอุปกรณ์: [{asset['category']}]
- สถานที่ตั้ง/ห้อง: {asset['room']}
"""

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM knowledge_base")
    all_kb = cursor.fetchall()
    conn.close()
    
    kb_match = None
    for kb in all_kb:
        if user_message and (kb['keyword'] in user_message.lower() or kb['issue_title'] in user_message):
            kb_match = kb
            break

    internal_solution = kb_match['solution_steps'] if kb_match else "1. ตรวจสอบการเชื่อมต่อสายสัญญาณและปลั๊กไฟเบื้องต้น\n2. ทำการรีสตาร์ทอุปกรณ์ใหม่ 1 ครั้งเพื่อเคลียร์ระบบ\n3. หากยังไม่สามารถแก้ไขได้ กรุณากดปุ่มเปิด Ticket แจ้งช่าง"

    prompt_header = f"""คุณคือ "AI IT Support" ของระบบ IT Helpdesk ศาลจังหวัดเพชรบูรณ์
หน้าที่ของคุณคือช่วยวิเคราะห์ปัญหา IT และแนะนำแนวทางแก้ไขตามคู่มือภายในองค์กรอย่างสุภาพ นอบน้อม และอ่านง่าย

{asset_context}

[คู่มือแนวทางแก้ไขจากระบบภายใน]
{internal_solution}

[คำแนะนำในการตอบ]
1. ให้นำขั้นตอนจากคู่มือภายในด้านบนมาเรียบเรียงอธิบายเป็นข้อๆ ให้ชัดเจนตรงกับอาการที่ผู้ใช้แจ้ง (หากมีรูปภาพประกอบ ให้ช่วยวิเคราะห์อาการจากรูปภาพร่วมด้วย)
2. ถ้าผู้ใช้พิมพ์โต้ตอบเพิ่มเติม ให้แนะนำขั้นตอนการตรวจสอบเชิงลึกถัดไป
3. หากลองแก้ตามขั้นตอนพื้นฐานแล้วยังไม่หาย จึงค่อยแนะนำให้กดปุ่ม "🔴 แก้ไขไม่ได้" เพื่อเปิด Ticket แจ้งช่าง IT ศาลฯ"""

    full_prompt = prompt_header + "\n\n[ประวัติการสนทนาที่ผ่านมา]\n"
    for msg in history:
        role_label = "ผู้ใช้" if msg['role'] == 'user' else "AI"
        full_prompt += f"{role_label}: {msg['content']}\n"
    
    full_prompt += f"ผู้ใช้: {user_message}\nAI Support:"

    try:
        contents = [full_prompt]
        
        if image_base64:
            import base64
            header, encoded = image_base64.split(",", 1) if "," in image_base64 else ("", image_base64)
            mime_type = header.split(";")[0].replace("data:", "") if "data:" in header else "image/jpeg"
            image_bytes = base64.b64decode(encoded)
            
            contents.append(
                types.Part.from_bytes(
                    data=image_bytes,
                    mime_type=mime_type,
                )
            )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                temperature=0.3,
            )
        )
        
        reply_text = response.text
        return jsonify({'success': True, 'reply': reply_text})
        
    except Exception as e:
        error_message = str(e)
        print(f"Gemini API Error Detail: {error_message}")
        
        fallback_reply = f"💡 **คำแนะนำเบื้องต้นจากระบบ (คู่มือภายใน):**\n{internal_solution}\n\nหากยังไม่สามารถแก้ไขได้ กรุณากดปุ่ม **[🔴 แก้ไขไม่ได้ (แจ้งช่าง)]** ด้านล่างเพื่อส่งเรื่องให้ช่าง IT ดำเนินการต่อครับ"
        return jsonify({'success': True, 'reply': fallback_reply}), 200


@app.route('/api/kb/edit', methods=['POST'])
def api_kb_edit():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    kb_id = data.get('id')
    keyword = data.get('keyword', '').strip().lower()
    issue_title = data.get('issue_title', '').strip()
    solution_steps = data.get('solution_steps', '').strip()

    if not all([kb_id, keyword, issue_title, solution_steps]):
        return jsonify({'success': False, 'message': 'กรุณากรอกข้อมูลให้ครบถ้วน'})

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE knowledge_base 
        SET keyword = %s, issue_title = %s, solution_steps = %s
        WHERE id = %s
    """, (keyword, issue_title, solution_steps, kb_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'แก้ไขข้อมูลคู่มือสำเร็จ'})


@app.route('/api/kb/bulk-delete', methods=['POST'])
def api_kb_bulk_delete():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    ids = data.get('ids', [])
    
    if not ids:
        return jsonify({'success': False, 'message': 'ไม่ได้เลือกรายการที่ต้องการลบ'})

    conn = get_db()
    cursor = conn.cursor()
    placeholders = ','.join(['%s'] * len(ids))
    cursor.execute(f"DELETE FROM knowledge_base WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ลบคู่มือสำเร็จ {len(ids)} รายการ'})


@app.route('/api/kb/import', methods=['POST'])
def api_kb_import():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'ไม่พบไฟล์ที่อัปโหลด'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'ยังไม่ได้เลือกไฟล์'})

    try:
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        
        conn = get_db()
        cursor = conn.cursor()
        count = 0
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: 
                continue
            
            keyword = str(row[0]).strip().lower()
            issue_title = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            solution_steps = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            
            if keyword and issue_title and solution_steps:
                cursor.execute("""
                    INSERT INTO knowledge_base (keyword, issue_title, solution_steps)
                    VALUES (%s, %s, %s)
                """, (keyword, issue_title, solution_steps))
                count += 1

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'นำเข้าข้อมูลสำเร็จ {count} รายการ'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาดในการอ่านไฟล์: {str(e)}'})


@app.route('/api/settings/update-form-url', methods=['POST'])
def update_form_url():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    url = data.get('url', '').strip()
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("REPLACE INTO app_settings (`key`, `value`) VALUES ('google_form_url', %s)", (url,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'บันทึกลิงก์ Google Form สำเร็จแล้ว'})


@app.route('/admin/export-kb-excel')
def export_kb_excel():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, keyword, issue_title, solution_steps FROM knowledge_base")
    kb_data = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Knowledge Base"

    ws.append(['ID', 'Keyword', 'หัวข้ออาการเสีย', 'ขั้นตอนการแก้ไข'])

    for kb in kb_data:
        ws.append([kb['id'], kb['keyword'], kb['issue_title'], kb['solution_steps']])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Knowledge_Base_Export.xlsx'
    )


@app.route('/admin/export-users-excel')
def export_users_excel():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT user_id, fullname, position, department, room, role FROM users ORDER BY user_id")
    users_data = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Users List"

    ws.append(['รหัสบุคลากร', 'ชื่อ-นามสกุล', 'ตำแหน่ง', 'แผนก', 'ห้อง/สถานที่', 'สิทธิ์การใช้งาน'])

    for u in users_data:
        ws.append([u['user_id'], u['fullname'], u['position'], u['department'], u['room'], u['role']])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Users_Management_Export.xlsx'
    )


@app.route('/admin/export-assets-excel')
def export_assets_excel():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT a.asset_id, a.asset_name, a.category, a.warranty, a.serial_number, a.room, a.status, a.assigned_user_id, u.fullname 
        FROM assets a 
        LEFT JOIN users u ON a.assigned_user_id = u.user_id
        ORDER BY a.asset_id
    """)
    assets_data = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Assets List"

    ws.append(['เลขครุภัณฑ์', 'ชื่ออุปกรณ์ / สเปก', 'ชนิดอุปกรณ์', 'ประกัน', 'Serial Number', 'ห้อง / สถานที่', 'สถานะ', 'รหัสผู้ถือครอง', 'ชื่อผู้ถือครอง'])

    for ast in assets_data:
        ws.append([
            ast['asset_id'], ast['asset_name'], ast['category'], ast['warranty'], 
            ast['serial_number'], ast['room'], ast['status'], 
            ast['assigned_user_id'] or '-', ast['fullname'] or 'อุปกรณ์ส่วนกลาง'
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Assets_Management_Export.xlsx'
    )


@app.route('/admin/export-settings-excel')
def export_settings_excel():
    if 'user_id' not in session or session.get('role') != 'admin':
        return redirect(url_for('login'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT dept_name FROM departments ORDER BY id ASC")
    depts = cursor.fetchall()
    cursor.execute("SELECT room_name FROM rooms ORDER BY id ASC")
    rooms = cursor.fetchall()
    cursor.execute("SELECT category_name FROM categories ORDER BY id ASC")
    cats = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    
    ws1 = wb.active
    ws1.title = "รายชื่อแผนก"
    ws1.append(['ลำดับ', 'ชื่อแผนก'])
    for idx, d in enumerate(depts, 1):
        ws1.append([idx, d['dept_name']])

    ws2 = wb.create_sheet(title="ห้องและสถานที่")
    ws2.append(['ลำดับ', 'ชื่อห้อง / สถานที่'])
    for idx, r in enumerate(rooms, 1):
        ws2.append([idx, r['room_name']])

    ws3 = wb.create_sheet(title="ชนิดอุปกรณ์")
    ws3.append(['ลำดับ', 'ชนิด / หมวดหมู่อุปกรณ์'])
    for idx, c in enumerate(cats, 1):
        ws3.append([idx, c['category_name']])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Master_Data_Export.xlsx'
    )


@app.route('/api/delete-ticket', methods=['POST'])
def delete_ticket():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    ticket_id = data.get('ticket_id')

    if not ticket_id:
        return jsonify({'success': False, 'message': 'ไม่พบรหัส Ticket ที่ต้องการลบ'})

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM tickets WHERE ticket_id = %s", (ticket_id,))
    cursor.execute("DELETE FROM repair_logs WHERE ticket_id = %s", (ticket_id,))
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'message': 'ลบรายการแจ้งซ่อมสำเร็จ'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)