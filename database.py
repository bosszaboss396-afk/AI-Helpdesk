import pymysql

DB_HOST = 'localhost'
DB_USER = 'root'
DB_PASS = ''
DB_NAME = 'ai_helpdesk_db'

def init_db():
    """ฟังก์ชันสร้างฐานข้อมูลและสร้างตารางหลัก (MySQL Version)"""
    conn = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASS, database=DB_NAME)
    cursor = conn.cursor()

    # 1. ตาราง users
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        user_id VARCHAR(50) PRIMARY KEY,
        password VARCHAR(255) NOT NULL,
        fullname VARCHAR(255) NOT NULL,
        department VARCHAR(255) NOT NULL,
        position VARCHAR(255) NOT NULL,
        room VARCHAR(255),
        role VARCHAR(50) NOT NULL DEFAULT 'user'
    );
    ''')

    # 2. ตาราง assets
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS assets (
        asset_id VARCHAR(100) PRIMARY KEY,
        asset_name VARCHAR(255) NOT NULL,
        category VARCHAR(255) NOT NULL,
        room VARCHAR(255) NOT NULL,
        assigned_user_id VARCHAR(50),
        warranty VARCHAR(100),
        serial_number VARCHAR(100),
        status VARCHAR(50) NOT NULL DEFAULT 'ปกติ'
    );
    ''')

    # 3. ตาราง tickets
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS tickets (
        ticket_id VARCHAR(50) PRIMARY KEY,
        created_at DATETIME NOT NULL,
        user_id VARCHAR(50) NOT NULL,
        asset_id VARCHAR(100) NOT NULL,
        issue_description TEXT NOT NULL,
        image_data LONGTEXT,
        status VARCHAR(50) NOT NULL DEFAULT 'pending',
        assigned_admin VARCHAR(255),
        repair_note TEXT
    );
    ''')

    # 4. ตาราง repair_logs
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS repair_logs (
        log_id INTEGER PRIMARY KEY AUTO_INCREMENT,
        ticket_id VARCHAR(50) NOT NULL,
        asset_id VARCHAR(100) NOT NULL,
        repaired_at DATETIME NOT NULL,
        issue_summary TEXT NOT NULL,
        action_taken TEXT NOT NULL,
        ai_resolved INTEGER NOT NULL DEFAULT 0
    );
    ''')

    # 5. ตาราง knowledge_base
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS knowledge_base (
        id INTEGER PRIMARY KEY AUTO_INCREMENT,
        keyword VARCHAR(255) NOT NULL,
        issue_title VARCHAR(255) NOT NULL,
        solution_steps TEXT NOT NULL
    )
    ''')

    # 6. ตาราง app_settings
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS app_settings (
        `key` VARCHAR(100) PRIMARY KEY,
        `value` TEXT
    )
    ''')

    # 7. ตาราง departments
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTO_INCREMENT,
        dept_name VARCHAR(255) UNIQUE NOT NULL
    )
    ''')

    # 8. ตาราง rooms
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS rooms (
        id INTEGER PRIMARY KEY AUTO_INCREMENT,
        room_name VARCHAR(255) UNIQUE NOT NULL
    )
    ''')

    # 9. ตาราง categories
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTO_INCREMENT,
        category_name VARCHAR(255) UNIQUE NOT NULL
    )
    ''')

    conn.commit()
    conn.close()
    print("✅ สร้างตารางฐานข้อมูล MySQL เรียบร้อยแล้ว")

if __name__ == "__main__":
    init_db()