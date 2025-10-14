import sqlite3
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from datetime import datetime

# --- 1. การตั้งค่า ---
# DB_FILE = "Drug_Supply.db"  <-- ปิดบรรทัดนี้ไป

# --- ให้เปลี่ยนเป็น Path เต็มแบบนี้แทน ---

# ตัวอย่างสำหรับ Windows (สำคัญ: ให้ใส่ r นำหน้า string เพื่อป้องกันปัญหาเรื่อง backslash \)
DB_FILE = r"D:\งานเต้ย\12.Pharmacy\1.สารน้ำ\Google sheet\Drug_Supply.db"
# เปลี่ยนเป็น Path เต็มที่ชี้ไปหาไฟล์ credentials.json
GSPREAD_CREDENTIALS = r"D:\งานเต้ย\12.Pharmacy\1.สารน้ำ\Google sheet\credentials.json"
SHEET_ID = "19EPvMA2LHyYgG2ljOvYNlseCRxH0nZa4P89tpTazKUA"
WORKSHEET_NAME = "Sheet1"

# --- 2. ฟังก์ชันหลักในการทำงาน ---
def main():
    print("🚀 เริ่มกระบวนการอัปเดตข้อมูล...")
    try:
        # --- 3. ดึงข้อมูลจาก SQLite ---
        print(f"กำลังเชื่อมต่อฐานข้อมูล: {DB_FILE}")
        with sqlite3.connect(DB_FILE) as conn:
            sql_query = "SELECT * FROM view_pr_plan"
            df = pd.read_sql_query(sql_query, conn)
            print(f"ดึงข้อมูลจาก 'view_pr_plan' สำเร็จ จำนวน {len(df)} แถว")

        # --- กรองข้อมูล ---
        if 'Status' in df.columns:
            print(f"ข้อมูลก่อนกรอง: {len(df)} แถว")
            df_filtered = df[df['Status'].astype(str).str.upper() == 'TRUE'].copy()
            print(f"ข้อมูลหลังกรอง (Status=TRUE): {len(df_filtered)} แถว")
        else:
            print("⚠️ ไม่พบคอลัมน์ 'Status' ในข้อมูล, ข้ามขั้นตอนการกรอง")
            df_filtered = df.copy()

        # --- เชื่อมต่อ Google Sheets ---
        gc = gspread.service_account(filename=GSPREAD_CREDENTIALS)
        spreadsheet = gc.open_by_key(SHEET_ID)
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)

        # --- 4. ล้างชีตทั้งหมด (ทั้งค่าและรูปแบบ) ---
        print("กำลังล้างข้อมูลและรูปแบบเก่าทั้งหมดในชีต...")
        # สร้างคำสั่งสำหรับล้าง format ทั้งหมด
        clear_formatting_request = {
            "updateCells": {
                "range": {"sheetId": worksheet.id},
                "fields": "userEnteredFormat"
            }
        }
        spreadsheet.batch_update({"requests": [clear_formatting_request]})
        # ล้างค่าข้อมูลทั้งหมด
        worksheet.clear()
        print("✅ ล้างชีตเรียบร้อย")

        # จัดการกรณีไม่มีข้อมูลหลังกรอง
        if df_filtered.empty:
            print("⚠️ ไม่พบข้อมูลที่มี Status = TRUE, หยุดการทำงาน")
            worksheet.update('A1', 'ไม่พบข้อมูลที่มี Status = TRUE ณ วันที่ ' + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            print("✅ แจ้งสถานะในชีตเรียบร้อย")
            return

        # --- 5. เขียนข้อมูลใหม่ลงในชีต ---
        run_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df_filtered.loc[:, 'Update'] = run_datetime
        
        print(f"กำลังเขียนข้อมูลใหม่ลงในชีต '{WORKSHEET_NAME}'...")
        set_with_dataframe(worksheet, df_filtered, resize=True)
        print("✅ เขียนข้อมูลลง Google Sheets สำเร็จ!")

        # --- 6. จัดรูปแบบชีตใหม่ทั้งหมด (ตั้งค่าพื้นฐาน + จัดรูปแบบพิเศษ) ---
        print("กำลังจัดรูปแบบชีต...")
        requests = []
        
        # รูปแบบที่ 1: ตั้งค่าพื้นฐาน (ตัวอักษรสีดำ) สำหรับข้อมูลทั้งหมด
        default_text_format = {
            "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}
        }
        requests.append({
            "repeatCell": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 1}, # เริ่มจากแถว 2 ไปจนสุด
                "cell": {"userEnteredFormat": default_text_format},
                "fields": "userEnteredFormat(textFormat)"
            }
        })
        
        # รูปแบบที่ 2: รูปแบบพิเศษสำหรับ Header ที่ต้องการ
        special_header_format = {
            "backgroundColor": { "red": 59/255, "green": 113/255, "blue": 104/255 },
            "textFormat": { "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True },
            "horizontalAlignment": "CENTER"
        }
        headers = list(df_filtered.columns)
        cols_to_format = ['Plan_PR', 'Value']
        
        for col_name in cols_to_format:
            if col_name in headers:
                col_index = headers.index(col_name)
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 0, "endRowIndex": 1, # แถวแรกเท่านั้น
                            "startColumnIndex": col_index, "endColumnIndex": col_index + 1
                        },
                        "cell": {"userEnteredFormat": special_header_format},
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                    }
                })

        if len(requests) > 1: # มีคำสั่งมากกว่า 1 (มากกว่าแค่ตั้งค่าสีดำ)
            spreadsheet.batch_update({"requests": requests})
            print(f"✅ จัดรูปแบบชีตสำเร็จ!")
        else:
            print("⚠️ ไม่พบคอลัมน์สำหรับจัดรูปแบบพิเศษ, ตั้งค่าเฉพาะสีตัวอักษรพื้นฐาน")


    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")

if __name__ == "__main__":
    main()
