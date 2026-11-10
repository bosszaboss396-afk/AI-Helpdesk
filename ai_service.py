import pymysql
import google.generativeai as genai
import os
import base64
from dotenv import load_dotenv

# ----------------------------------------------------
# ⚙️ การตั้งค่า Database MySQL (ให้ตรงกับ app.py)
# ----------------------------------------------------
DB_HOST = 'localhost'
DB_USER = 'root'
DB_PASS = ''
DB_NAME = 'ai_helpdesk_db'

# ตั้งค่า API Key ของ Gemini (โหลดค่าจากไฟล์ .env หรือระบุตรงนี้)
load_dotenv(override=True)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

def get_db_connection():
    """สร้างการเชื่อมต่อกับ MySQL"""
    return pymysql.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        cursorclass=pymysql.cursors.DictCursor
    )

def get_asset_and_history(asset_id):
    """ดึงข้อมูลครุภัณฑ์และประวัติการซ่อมย้อนหลังเพื่อส่งให้ AI วิเคราะห์"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # ดึงข้อมูลครุภัณฑ์ (เปลี่ยน ? เป็น %s)
    cursor.execute("SELECT asset_name, category, room FROM assets WHERE asset_id = %s", (asset_id,))
    asset = cursor.fetchone()

    # ดึงประวัติการซ่อมย้อนหลังล่าสุด 3 เคส (เปลี่ยน ? เป็น %s)
    cursor.execute("""
        SELECT issue_summary, action_taken 
        FROM repair_logs 
        WHERE asset_id = %s 
        ORDER BY log_id DESC LIMIT 3
    """, (asset_id,))
    logs = cursor.fetchall()
    conn.close()

    return asset, logs

def analyze_issue_with_ai(asset_id, user_message, image_base64=None):
    """ส่ง Prompt + Context (ข้อมูลครุภัณฑ์ + ประวัติซ่อม + รูปภาพ) ให้ Gemini API"""
    asset, logs = get_asset_and_history(asset_id)
    
    # 📌 เปลี่ยนการอ้างอิงเป็นชื่อคอลัมน์ (เช่น asset['asset_name']) เพราะเราใช้ DictCursor
    asset_info = f"อุปกรณ์: {asset['asset_name']} ({asset['category']}) ติดตั้งที่: {asset['room']}" if asset else "อุปกรณ์ส่วนกลาง"
    
    history_text = ""
    if logs:
        # 📌 เปลี่ยน log[0], log[1] เป็นชื่อคอลัมน์
        history_text = "ประวัติการซ่อมในอดีต:\n" + "\n".join([f"- อาการ: {log['issue_summary']} | การแก้ไข: {log['action_taken']}" for log in logs])
    else:
        history_text = "ประวัติการซ่อมในอดีต: ไม่มีประวัติเสียซ้ำ"

    # สร้าง System Prompt สโคปบริบทศาลจังหวัดเพชรบูรณ์
    system_prompt = f"""
    คุณคือ AI เจ้าหน้าที่วิเคราะห์และแก้ไขปัญหาไอทีประจำศาลจังหวัดเพชรบูรณ์
    ข้อมูลอุปกรณ์ที่แจ้งเสีย: {asset_info}
    {history_text}

    คำแนะนำในการตอบ:
    1. ตอบเป็นภาษาไทย สุภาพ สั้นกระชับ เข้าใจง่าย
    2. หากมีรูปภาพประกอบ ให้นำลักษณะอาการจากรูปภาพมาวิเคราะห์ร่วมด้วย
    3. เสนอวิธีแก้ไขปัญหาเบื้องต้น 1-3 ข้อ เพื่อให้เจ้าหน้าที่ลองทำตาม
    4. หากเป็นปัญหาร้ายแรงเกี่ยวกับสายสัญญาณ ไฟฟ้า หรือระบบศาล ให้แนะนำให้แจ้งช่าง/แอดมินทันที
    """

    try:
        # 📌 อัปเดตโมเดลเป็นเวอร์ชันตามความเหมาะสม
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # จัดเตรียม Contents สำหรับส่งให้โมเดล (รองรับทั้งข้อความและรูปภาพ)
        contents = [system_prompt]
        
        # ถ้ามีรูปภาพแนบมาแบบ Base64 ให้แปลงเป็น Part ของรูปภาพ
        if image_base64:
            header, encoded = image_base64.split(",", 1) if "," in image_base64 else ("", image_base64)
            mime_type = header.split(";")[0].replace("data:", "") if "data:" in header else "image/jpeg"
            image_bytes = base64.b64decode(encoded)
            
            contents.append({
                'mime_type': mime_type,
                'data': image_bytes
            })
            
        contents.append(f"อาการเสียที่พบ: {user_message if user_message else 'ผู้ใช้งานแนบรูปภาพอาการเสียมาโดยไม่ได้พิมพ์ข้อความเพิ่มเติม'}")

        response = model.generate_content(contents)
        return response.text
        
    except Exception as e:
        # 📌 เพิ่มบล็อกดักจับ Error เพื่อป้องกันไม่ให้ระบบพัง (Server Crash)
        print(f"❌ Gemini API Error: {e}")
        return f"ขออภัยครับ ขณะนี้ระบบ AI ไม่สามารถวิเคราะห์อาการได้ (พบปัญหาขัดข้อง) กรุณากดปุ่มแจ้งช่าง (เปิด Ticket) ด้านล่างเพื่อให้ทีม IT ดำเนินการต่อครับ"