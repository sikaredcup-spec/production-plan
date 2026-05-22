"""
db_handler.py — จัดการฐานข้อมูลผ่าน Google Sheets
ติดตั้ง: pip install gspread google-auth pandas openpyxl
"""

import os, json
from datetime import datetime, date
from typing import Optional
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ─── GOOGLE SHEETS SETUP ─────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_sheet_client():
    creds = Credentials.from_service_account_file(
        os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json"),
        scopes=SCOPES
    )
    return gspread.authorize(creds)

def open_spreadsheet():
    gc = get_sheet_client()
    return gc.open_by_key(os.getenv("GOOGLE_SHEET_ID"))

# ─── SHEET NAMES ─────────────────────────────────────────────────
SHEET_PLAN       = "ProductionPlan"    # แผนผลิต (import จาก Excel)
SHEET_STATUS     = "StatusLog"         # Log การอัพเดท Status
SHEET_USERS      = "Users"             # LINE user_id → ชื่อ + เครื่อง

# ─── IMPORT EXCEL → GOOGLE SHEETS ────────────────────────────────
def import_excel_to_sheet(excel_path: str):
    """
    อ่าน Excel แผนผลิต แล้วบันทึกลง Google Sheets
    เรียกใช้: python db_handler.py --import test1_.xlsx
    """
    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip()

    # แปลงวันที่ให้เป็น string
    for col in ["Start Date", "Finish Date", "Delivery"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")

    # ทำความสะอาด Big/Small batch
    df["Big/ Small batch"] = df["Big/ Small batch"].str.strip()

    # เพิ่มคอลัมน์ production_status (ค่าเริ่มต้น = waiting)
    if "production_status" not in df.columns:
        df["production_status"] = "waiting"
    if "updated_by" not in df.columns:
        df["updated_by"] = ""
    if "updated_at" not in df.columns:
        df["updated_at"] = ""
    if "issue_note" not in df.columns:
        df["issue_note"] = ""

    df = df.fillna("")

    # บันทึกลง Google Sheets
    ss = open_spreadsheet()
    try:
        ws = ss.worksheet(SHEET_PLAN)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(SHEET_PLAN, rows=1000, cols=30)

    # เขียน header + data
    ws.update([df.columns.tolist()] + df.values.tolist())
    print(f"✅ Import สำเร็จ: {len(df)} รายการ → {SHEET_PLAN}")
    return len(df)

# ─── GET TASKS ────────────────────────────────────────────────────
def get_tasks_for_user(user_id: str, target_date: str = None) -> list[dict]:
    """ดึงงานของ user (กรองตาม machine ที่ user รับผิดชอบ)"""
    if target_date is None:
        target_date = date.today().strftime("%Y-%m-%d")

    user_info = get_user_info(user_id)
    if not user_info:
        return []

    machines = user_info.get("machines", "").split(",")
    machines = [m.strip() for m in machines if m.strip()]

    ss    = open_spreadsheet()
    ws    = ss.worksheet(SHEET_PLAN)
    rows  = ws.get_all_records()
    tasks = []
    for row in rows:
        # กรองตาม Machine และ Start Date
        machine    = str(row.get("Machine", "")).strip()
        start_date = str(row.get("Start Date", ""))
        if machine in machines and start_date == target_date:
            tasks.append(row)
    return tasks

def get_all_assignments_today() -> dict[str, list]:
    """ดึงงานทุกคน จัดกลุ่มตาม user_id"""
    ss       = open_spreadsheet()
    users_ws = ss.worksheet(SHEET_USERS)
    users    = users_ws.get_all_records()

    result = {}
    for u in users:
        uid   = u.get("user_id", "")
        tasks = get_tasks_for_user(uid)
        if uid:
            result[uid] = tasks
    return result

# ─── UPDATE STATUS ────────────────────────────────────────────────
def update_status(batch_no: str, new_status: str, user_id: str,
                  issue_note: str = "") -> tuple[bool, Optional[dict]]:
    """
    อัพเดท production_status ใน Google Sheets
    คืนค่า (success, task_dict)
    """
    ss    = open_spreadsheet()
    ws    = ss.worksheet(SHEET_PLAN)
    rows  = ws.get_all_records()
    headers = ws.row_values(1)

    batch_col       = headers.index("BATCH NO.")          + 1
    status_col      = headers.index("production_status")  + 1
    updated_by_col  = headers.index("updated_by")         + 1
    updated_at_col  = headers.index("updated_at")         + 1
    issue_note_col  = headers.index("issue_note")         + 1

    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user_info = get_user_info(user_id)
    user_name = user_info.get("name", user_id) if user_info else user_id

    for i, row in enumerate(rows, start=2):  # row 1 = header
        row_batch = str(row.get("BATCH NO.", "")).split(".")[0]
        if row_batch == str(batch_no).split(".")[0]:
            ws.update_cell(i, status_col,     new_status)
            ws.update_cell(i, updated_by_col, user_name)
            ws.update_cell(i, updated_at_col, now_str)
            if issue_note:
                ws.update_cell(i, issue_note_col, issue_note)

            # บันทึก Log
            log_status_change(batch_no, new_status, user_name, issue_note)
            return True, row

    return False, None

# ─── STATUS LOG ───────────────────────────────────────────────────
def log_status_change(batch_no, status, user_name, note=""):
    ss = open_spreadsheet()
    try:
        ws = ss.worksheet(SHEET_STATUS)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(SHEET_STATUS, rows=10000, cols=10)
        ws.append_row(["timestamp", "batch_no", "status", "updated_by", "note"])

    ws.append_row([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        str(batch_no),
        status,
        user_name,
        note
    ])

# ─── DAILY SUMMARY ────────────────────────────────────────────────
def get_daily_summary(target_date: str = None) -> dict:
    if target_date is None:
        target_date = date.today().strftime("%Y-%m-%d")

    ss   = open_spreadsheet()
    ws   = ss.worksheet(SHEET_PLAN)
    rows = ws.get_all_records()

    summary = {
        "date": target_date,
        "waiting": 0, "in_progress": 0, "qc": 0,
        "done": 0, "issue": 0, "total": 0
    }

    for row in rows:
        if str(row.get("Start Date", "")) == target_date:
            s = row.get("production_status", "waiting")
            summary[s] = summary.get(s, 0) + 1
            summary["total"] += 1

    return summary

# ─── USER MANAGEMENT ─────────────────────────────────────────────
def get_user_info(user_id: str) -> Optional[dict]:
    ss = open_spreadsheet()
    try:
        ws   = ss.worksheet(SHEET_USERS)
        rows = ws.get_all_records()
        for row in rows:
            if row.get("user_id") == user_id:
                return row
    except Exception:
        pass
    return None

def register_user(user_id: str, name: str, machines: str):
    """ลงทะเบียนลูกน้องใหม่ — เรียกจาก Admin"""
    ss = open_spreadsheet()
    try:
        ws = ss.worksheet(SHEET_USERS)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(SHEET_USERS, rows=200, cols=5)
        ws.append_row(["user_id", "name", "machines", "registered_at"])

    ws.append_row([user_id, name, machines,
                   datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

# ─── PENDING ISSUE (รอรับรายละเอียดปัญหา) ────────────────────────
_pending_issues: dict[str, str] = {}  # user_id → batch_no

def set_pending_issue(user_id: str, batch_no: str):
    _pending_issues[user_id] = batch_no

def get_pending_issue(user_id: str) -> Optional[str]:
    return _pending_issues.get(user_id)

def clear_pending_issue(user_id: str):
    _pending_issues.pop(user_id, None)

# ─── CLI: python db_handler.py --import <file.xlsx> ──────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "--import":
        import_excel_to_sheet(sys.argv[2])
    elif len(sys.argv) >= 2 and sys.argv[1] == "--summary":
        s = get_daily_summary()
        print(json.dumps(s, ensure_ascii=False, indent=2))
    else:
        print("Usage:")
        print("  python db_handler.py --import production_plan.xlsx")
        print("  python db_handler.py --summary")
