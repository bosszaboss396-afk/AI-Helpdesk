import pymysql
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO

# ----------------------------------------------------
# ⚙️ การตั้งค่า Database MySQL
# ----------------------------------------------------
DB_HOST = 'localhost'
DB_USER = 'root'
DB_PASS = ''
DB_NAME = 'ai_helpdesk_db'

def get_db_connection():
    return pymysql.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME
        # ใช้ Cursor ปกติ (Tuple) เพื่อให้ลอจิกการวนลูป Excel เดิมทำงานได้
    )

def export_tickets_to_excel():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT t.ticket_id, t.created_at, u.fullname, u.department, 
               a.asset_name, a.room, t.issue_description, t.status, 
               t.assigned_admin, t.repair_note
        FROM tickets t
        JOIN assets a ON t.asset_id = a.asset_id
        JOIN users u ON t.user_id = u.user_id
        ORDER BY t.created_at DESC
    """)
    tickets = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "รายงานการแจ้งซ่อม"

    headers = [
        "เลข Ticket", "วันเวลาแจ้ง", "ผู้แจ้งซ่อม", "ส่วนงาน/แผนก",
        "ครุภัณฑ์", "สถานที่", "อาการเสีย", "สถานะ", "ช่างผู้ดูแล", "วิธีแก้ไข"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="8B4513", end_color="8B4513", fill_type="solid")
    header_font = Font(name="Sarabun", size=11, bold=True, color="FFFFFF")

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    status_map = {
        'pending': 'รอรับเคส',
        'in_progress': 'กำลังดำเนินการ',
        'completed': 'ซ่อมสำเร็จ',
        'failed': 'ซ่อมไม่สำเร็จ'
    }

    for row_data in tickets:
        row_list = list(row_data)
        row_list[7] = status_map.get(row_list[7], row_list[7])
        ws.append(row_list)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output