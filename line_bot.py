"""
Production Monitor — LINE Bot Backend
ติดตั้ง: pip install flask line-bot-sdk openpyxl pandas python-dotenv
"""

import os, json, hashlib, hmac, base64
from datetime import datetime
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, PushMessageRequest,
    FlexMessage, FlexContainer,
    TextMessage
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent, PostbackEvent
from linebot.v3.exceptions import InvalidSignatureError
from dotenv import load_dotenv
import db_handler  # import โมดูลฐานข้อมูล

load_dotenv()

app = Flask(__name__)

CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
CHANNEL_SECRET       = os.getenv("LINE_CHANNEL_SECRET")

configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler       = WebhookHandler(CHANNEL_SECRET)

# ─── STATUS CONFIG ────────────────────────────────────────────────
STATUS_CONFIG = {
    "waiting":    {"label": "⏳ รอดำเนินการ", "color": "#95A5A6"},
    "in_progress":{"label": "🔄 กำลังผลิต",   "color": "#F39C12"},
    "qc":         {"label": "🔍 ตรวจสอบ QC",  "color": "#3498DB"},
    "done":       {"label": "✅ เสร็จสิ้น",    "color": "#27AE60"},
    "issue":      {"label": "🔴 มีปัญหา",      "color": "#E74C3C"},
}

# ─── WEBHOOK ENDPOINT ─────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# ─── TEXT MESSAGE HANDLER ─────────────────────────────────────────
@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_id = event.source.user_id
    text    = event.message.text.strip()

    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)

        # คำสั่ง: "งานวันนี้" หรือ "งาน"
        if text in ["งานวันนี้", "งาน", "my tasks", "task"]:
            tasks = db_handler.get_tasks_for_user(user_id)
            if not tasks:
                reply_text(line_bot_api, event.reply_token,
                           "ไม่พบงานที่ได้รับมอบหมายวันนี้ครับ")
                return
            messages = [build_task_flex(t) for t in tasks[:5]]  # max 5 cards
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=messages
                )
            )

        # คำสั่ง: "ดูงานวันนี้ทั้งหมด" → ดูงานทุกคนวันนี้
        elif text in ["ดูงานวันนี้ทั้งหมด", "งานทั้งหมด", "all tasks"]:
            all_tasks = db_handler.get_all_tasks_today()
            if not all_tasks:
                reply_text(line_bot_api, event.reply_token,
                           "ไม่พบงานวันนี้ครับ")
                return
            
            # สร้างข้อความสรุป
            msg_lines = [f"📋 งานวันนี้ทั้งหมด ({len(all_tasks)} รายการ)\n"]
            for i, task in enumerate(all_tasks[:20], 1):  # แสดงสูงสุด 20 รายการ
                batch_no = task.get("BATCH NO.", "-")
                machine  = task.get("Machine", "-")
                status   = task.get("production_status", "waiting")
                status_label = STATUS_CONFIG.get(status, {}).get("label", status)
                msg_lines.append(f"{i}. Batch {batch_no} | {machine} | {status_label}")
            
            if len(all_tasks) > 20:
                msg_lines.append(f"\n...และอีก {len(all_tasks) - 20} รายการ")
            
            reply_text(line_bot_api, event.reply_token, "\n".join(msg_lines))

        # คำสั่ง: "สรุป" → ดู Dashboard
        elif text in ["สรุป", "summary", "dashboard"]:
            summary = db_handler.get_daily_summary()
            msg = format_summary(summary)
            reply_text(line_bot_api, event.reply_token, msg)

        else:
            reply_text(line_bot_api, event.reply_token,
                       "พิมพ์ 'งานวันนี้' เพื่อดูงานของคุณ\n'ดูงานวันนี้ทั้งหมด' เพื่อดูงานทุกคน\n'สรุป' เพื่อดูภาพรวม")

# ─── POSTBACK HANDLER (กดปุ่ม Status) ────────────────────────────
@handler.add(PostbackEvent)
def handle_postback(event):
    user_id = event.source.user_id
    data    = dict(item.split("=") for item in event.postback.data.split("&"))

    action     = data.get("action")
    batch_no   = data.get("batch_no")
    new_status = data.get("status")

    if action == "update_status" and batch_no and new_status:
        success, task = db_handler.update_status(batch_no, new_status, user_id)

        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            if success:
                status_label = STATUS_CONFIG[new_status]["label"]
                msg = (f"✅ อัพเดทสำเร็จ!\n"
                       f"Batch: {batch_no}\n"
                       f"Status: {status_label}\n"
                       f"เวลา: {datetime.now().strftime('%H:%M น.')}")
                # ถ้ามีปัญหา ขอรายละเอียดเพิ่ม
                if new_status == "issue":
                    msg += "\n\nกรุณาพิมพ์รายละเอียดปัญหา เพื่อแจ้งหัวหน้าครับ"
                    db_handler.set_pending_issue(user_id, batch_no)
            else:
                msg = "❌ ไม่พบ Batch นี้ กรุณาตรวจสอบอีกครั้ง"
            reply_text(line_bot_api, event.reply_token, msg)

    elif action == "report_issue":
        # รับรายละเอียดปัญหา (ขั้นตอนต่อเนื่อง)
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            reply_text(line_bot_api, event.reply_token,
                       "กรุณาพิมพ์รายละเอียดปัญหาครับ")
            db_handler.set_pending_issue(user_id, batch_no)

# ─── BUILD FLEX MESSAGE (Task Card) ───────────────────────────────
def build_task_flex(task: dict) -> FlexMessage:
    """สร้าง Flex Message การ์ดงาน 1 ใบ"""
    batch_no    = str(task["batch_no"])
    machine     = task.get("machine", "-")
    description = task.get("description", "-")[:50]
    batch_size  = task.get("batch_size_kg", "-")
    start_date  = task.get("start_date", "-")
    current_status = task.get("production_status", "waiting")
    status_label   = STATUS_CONFIG.get(current_status, {}).get("label", current_status)
    status_color   = STATUS_CONFIG.get(current_status, {}).get("color", "#95A5A6")

    # สร้างปุ่ม Status ทั้ง 5 ปุ่ม
    status_buttons = []
    for key, cfg in STATUS_CONFIG.items():
        is_current = key == current_status
        status_buttons.append({
            "type": "button",
            "action": {
                "type": "postback",
                "label": cfg["label"],
                "data": f"action=update_status&batch_no={batch_no}&status={key}"
            },
            "style": "primary" if is_current else "secondary",
            "color": cfg["color"] if is_current else "#AAAAAA",
            "height": "sm",
            "margin": "xs"
        })

    flex_body = {
        "type": "bubble",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": status_color,
            "paddingAll": "md",
            "contents": [
                {
                    "type": "text",
                    "text": f"Batch: {batch_no}",
                    "color": "#FFFFFF",
                    "weight": "bold",
                    "size": "md"
                },
                {
                    "type": "text",
                    "text": status_label,
                    "color": "#FFFFFF",
                    "size": "sm"
                }
            ]
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "sm",
            "contents": [
                row_item("🏭 เครื่อง", machine),
                row_item("📋 สินค้า", description),
                row_item("⚖️ ขนาด Batch", f"{batch_size} KG"),
                row_item("📅 วันเริ่ม", str(start_date)),
                {"type": "separator", "margin": "md"},
                {
                    "type": "text",
                    "text": "เลือก Status:",
                    "size": "sm",
                    "color": "#555555",
                    "margin": "md",
                    "weight": "bold"
                }
            ] + status_buttons
        }
    }

    return FlexMessage(
        alt_text=f"งาน Batch {batch_no} — {status_label}",
        contents=FlexContainer.from_dict(flex_body)
    )

# ─── HELPER FUNCTIONS ─────────────────────────────────────────────
def row_item(label: str, value: str) -> dict:
    return {
        "type": "box",
        "layout": "horizontal",
        "contents": [
            {"type": "text", "text": label,       "size": "sm", "color": "#888888", "flex": 3},
            {"type": "text", "text": str(value),  "size": "sm", "color": "#333333", "flex": 5, "wrap": True}
        ]
    }

def reply_text(api: MessagingApi, reply_token: str, text: str):
    api.reply_message(
        ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=text)]
        )
    )

def format_summary(summary: dict) -> str:
    lines = [f"📊 สรุปการผลิต วันที่ {summary.get('date', '-')}\n"]
    for key, cfg in STATUS_CONFIG.items():
        count = summary.get(key, 0)
        lines.append(f"{cfg['label']}: {count} รายการ")
    lines.append(f"\n📦 รวมทั้งหมด: {summary.get('total', 0)} รายการ")
    return "\n".join(lines)

# ─── DAILY PUSH (เรียกจาก Scheduler) ─────────────────────────────
def push_daily_tasks():
    """ส่งงานประจำวันให้ลูกน้องทุกคน — เรียกผ่าน Scheduler ทุกเช้า 7:00"""
    assignments = db_handler.get_all_assignments_today()
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        for user_id, tasks in assignments.items():
            if not tasks:
                continue
            messages = [TextMessage(text=f"🌅 งานวันนี้ของคุณ ({len(tasks)} รายการ)")]
            messages += [build_task_flex(t) for t in tasks[:5]]
            line_bot_api.push_message(
                PushMessageRequest(to=user_id, messages=messages)
            )

if __name__ == "__main__":
    app.run(port=5000, debug=True)
