# คู่มือติดตั้ง Production Monitor (LINE Bot)

## โครงสร้างไฟล์
```
production_monitor/
├── line_bot.py          ← Backend หลัก (LINE Webhook)
├── db_handler.py        ← จัดการ Google Sheets ฐานข้อมูล
├── scheduler.py         ← ส่งงานประจำวัน + Export Excel
├── .env                 ← API Keys (สร้างจาก .env.example)
└── service_account.json ← Google Service Account Key
```

---

## ขั้นตอนที่ 1 — สมัคร LINE Official Account (ฟรี)

1. ไปที่ https://developers.line.biz
2. Login ด้วย LINE Account
3. กด **Create a new provider** → ตั้งชื่อบริษัท
4. กด **Create a new channel** → เลือก **Messaging API**
5. กรอกข้อมูล:
   - Channel name: `Production Monitor`
   - Channel description: ระบบติดตามการผลิต
6. กด **Create**
7. ไปที่ Tab **Basic settings** → Copy **Channel secret**
8. ไปที่ Tab **Messaging API** → กด **Issue** เพื่อสร้าง **Channel access token**
9. **บันทึก** ทั้ง 2 ค่าไว้ใน `.env`

---

## ขั้นตอนที่ 2 — ตั้ง Google Sheets ฐานข้อมูล

### 2.1 สร้าง Google Sheet
1. ไปที่ https://sheets.google.com → สร้าง Sheet ใหม่
2. ตั้งชื่อ: `Production Monitor DB`
3. Copy **Sheet ID** จาก URL: `https://docs.google.com/spreadsheets/d/**SHEET_ID**/edit`

### 2.2 สร้าง Service Account (Robot สำหรับเขียน Sheet)
1. ไปที่ https://console.cloud.google.com
2. สร้าง Project ใหม่ (หรือใช้ที่มีอยู่)
3. เปิด **Google Sheets API** และ **Google Drive API**
   - ค้นหา "Google Sheets API" → Enable
   - ค้นหา "Google Drive API" → Enable
4. ไปที่ **IAM & Admin → Service Accounts → Create Service Account**
   - ตั้งชื่อ: `production-monitor`
   - กด **Done**
5. คลิกที่ Service Account → Tab **Keys** → **Add Key → JSON**
6. Download ไฟล์ → เปลี่ยนชื่อเป็น `service_account.json`
7. วางไฟล์ไว้ในโฟลเดอร์ `production_monitor/`
8. Copy Email ของ Service Account (ลักษณะ `xxx@xxx.iam.gserviceaccount.com`)

### 2.3 Share Google Sheet ให้ Service Account
1. เปิด Google Sheet ที่สร้าง
2. กด **Share** → วาง Email ของ Service Account
3. เลือกสิทธิ์ **Editor** → Done

---

## ขั้นตอนที่ 3 — Deploy Server (ฟรี บน Render.com)

1. ไปที่ https://render.com → สมัคร/Login
2. กด **New → Web Service**
3. Connect GitHub repo (หรือ upload ไฟล์)
4. ตั้งค่า:
   - **Name**: production-monitor
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn line_bot:app`
5. ไปที่ **Environment** → เพิ่ม variables:
   ```
   LINE_CHANNEL_ACCESS_TOKEN = (ค่าจาก LINE)
   LINE_CHANNEL_SECRET       = (ค่าจาก LINE)
   GOOGLE_SHEET_ID           = (ค่าจาก Google Sheets URL)
   ```
6. กด **Deploy** → รอ 2-3 นาที
7. Copy **URL** เช่น `https://production-monitor.onrender.com`

---

## ขั้นตอนที่ 4 — เชื่อม Webhook URL กับ LINE

1. กลับไปที่ LINE Developers Console
2. Tab **Messaging API** → **Webhook settings**
3. ใส่ URL: `https://production-monitor.onrender.com/webhook`
4. กด **Verify** → ต้องขึ้น Success
5. เปิด **Use webhook** → ON
6. ปิด **Auto-reply messages** → OFF

---

## ขั้นตอนที่ 5 — Import แผนผลิตจาก Excel

```bash
# ครั้งแรก: ติดตั้ง libraries
pip install flask line-bot-sdk gspread google-auth pandas openpyxl python-dotenv gunicorn

# Copy .env.example เป็น .env แล้วใส่ค่า
cp .env.example .env

# Import Excel → Google Sheets
python db_handler.py --import test1_.xlsx
```

---

## ขั้นตอนที่ 6 — ลงทะเบียนลูกน้อง

ลูกน้องต้องทำ:
1. **Add LINE OA** เป็นเพื่อน (แสกน QR Code จาก LINE Developers)
2. พิมพ์ข้อความอะไรก็ได้ → ระบบจะได้ user_id อัตโนมัติ

Admin ลงทะเบียนให้ลูกน้อง ผ่าน Python:
```python
import db_handler
# (user_id, ชื่อ, เครื่องที่รับผิดชอบ คั่นด้วย comma)
db_handler.register_user("Uxxxxxxxxxxxxxxxx", "สมชาย", "HSD3254,HSD-3262")
db_handler.register_user("Uxxxxxxxxxxxxxxxx", "สมหญิง", "T-3035,T-3036")
```

**วิธีหา user_id:** ดูจาก Log ของ Server เมื่อลูกน้องส่งข้อความมา

---

## ขั้นตอนที่ 7 — ทดสอบระบบ

```bash
# ทดสอบส่งงานให้ลูกน้องทันที
python scheduler.py --push-tasks

# Export Status กลับ Excel
python scheduler.py --export-excel

# รัน Daemon (ส่งงานอัตโนมัติทุกเช้า 7:00)
python scheduler.py --daemon
```

---

## วิธีใช้งาน (ลูกน้อง)

| พิมพ์ใน LINE | ผล |
|---|---|
| `งานวันนี้` | แสดงการ์ดงานทั้งหมดพร้อมปุ่ม Status |
| กดปุ่ม Status | อัพเดทสถานะทันที |
| `สรุป` | ดูภาพรวมงานทั้งหมดวันนี้ |

## วิธีใช้งาน (หัวหน้า)

1. เปิด Google Sheets → ดู Sheet `ProductionPlan` แบบ Realtime
2. รัน `python scheduler.py --export-excel` เพื่อได้ไฟล์ Excel

---

## requirements.txt
```
flask==3.0.0
line-bot-sdk==3.11.0
gspread==6.1.2
google-auth==2.29.0
pandas==2.2.1
openpyxl==3.1.2
python-dotenv==1.0.1
gunicorn==21.2.0
```
