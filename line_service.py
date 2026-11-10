import os
import requests
from dotenv import load_dotenv

load_dotenv(override=True)
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
LINE_GROUP_ID = os.getenv("LINE_GROUP_ID")  # ID กลุ่มไลน์ของทีมช่าง IT ศาลฯ

def send_line_notification(ticket_id, user_name, department, asset_name, room, issue):
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_GROUP_ID:
        print("⚠️ LINE Token หรือ Group ID ยังไม่ได้ตั้งค่าในไฟล์ .env")
        return False

    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    
    # จัดรูปแบบข้อความแจ้งเตือนให้สวยงามและอ่านง่าย
    message_text = (
        f"🚨 มีการแจ้งซ่อมใหม่ (AI Smart IT Helpdesk)\n"
        f"━━━━━━━━━━━━━━━\n"
        f"📌 เลขเคส: {ticket_id}\n"
        f"👤 ผู้แจ้ง: {user_name}\n"
        f"🏢 แผนก: {department}\n"
        f"💻 อุปกรณ์: {asset_name}\n"
        f"📍 สถานที่: {room}\n"
        f"🛠️ อาการเสีย: {issue}\n"
        f"━━━━━━━━━━━━━━━\n"
        f"👉 กรุณาเข้าสู่ระบบ Admin เพื่อรับเคส"
    )

    payload = {
        "to": LINE_GROUP_ID,
        "messages": [
            {
                "type": "text",
                "text": message_text
            }
        ]
    }

    try:
        response = requests.post(url, headers=headers, json=payload)
        if response.status_code == 200:
            print("✅ ส่งการแจ้งเตือน LINE สำเร็จ")
            return True
        else:
            print(f"❌ ส่ง LINE ไม่สำเร็จ: {response.text}")
            return False
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ LINE API: {str(e)}")
        return False