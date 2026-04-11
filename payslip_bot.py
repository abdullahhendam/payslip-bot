"""
بوت تيليجرام لإرسال فيش القبض
يحتاج:
  - employees.xlsx  (ملف بيانات الموظفين)
  - payslips.pdf    (ملف الفيش الشهري)
ضع الملفين في نفس مجلد السكريبت
"""

import re
import io
import logging
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)

# ─── إعدادات ────────────────────────────────────────────────
BOT_TOKEN  = "8743584646:AAHN2TIMN47GkMLHayZSy4RK0EkdxQ8ssB8"
EXCEL_FILE = "employees.xlsx"
PDF_FILE   = "payslips.pdf"
MAX_TRIES  = 3      # عدد محاولات الرقم السري قبل الحظر
BLOCK_SECS = 300    # مدة الحظر بالثواني (5 دقائق)

# ─── مراحل المحادثة ──────────────────────────────────────────
ASK_CODE, ASK_PIN = range(2)

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    filename="bot.log",
    encoding="utf-8"
)
logger = logging.getLogger(__name__)


# ─── تحميل البيانات ──────────────────────────────────────────
def load_employees():
    df = pd.read_excel(EXCEL_FILE, header=None)
    header_row = None
    for i, row in df.iterrows():
        if 'الكود' in row.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("لم يتم العثور على هيدر الجدول في ملف Excel")
    df = pd.read_excel(EXCEL_FILE, header=header_row)
    employees = {}
    for _, row in df.iterrows():
        code = str(row.get('الكود', '')).strip()
        pin  = str(row.get('الرقم السري', '')).strip()
        name = str(row.get('الاسم', '')).strip()
        if code and code != 'nan' and pin and pin != 'nan':
            employees[code] = {'pin': pin, 'name': name}
    logger.info(f"تم تحميل {len(employees)} موظف من Excel")
    return employees


def build_pdf_map():
    """خريطة: كود الموظف -> (رقم الصفحة، موضع الظرف 0/1/2)"""
    emp_map = {}
    with pdfplumber.open(PDF_FILE) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            h = page.height
            w = page.width
            thirds = [
                page.within_bbox((0, 0,          w, h / 3)),
                page.within_bbox((0, h / 3,      w, h * 2 / 3)),
                page.within_bbox((0, h * 2 / 3,  w, h)),
            ]
            for pos, section in enumerate(thirds):
                text = section.extract_text() or ''
                codes = re.findall(r'\b([0-9]{3,4})\b', text)
                for code in codes:
                    if code != '2026' and len(code) >= 3 and code not in emp_map:
                        emp_map[code] = (page_num, pos)
    logger.info(f"تم بناء خريطة PDF: {len(emp_map)} موظف")
    return emp_map


def extract_slip(page_num: int, position: int) -> bytes:
    """استخراج ظرف الموظف فقط (position: 0=أول، 1=تاني، 2=تالت)"""
    src = fitz.open(PDF_FILE)
    page = src[page_num - 1]
    h = page.rect.height
    w = page.rect.width
    y0 = (h / 3) * position
    y1 = (h / 3) * (position + 1)
    clip = fitz.Rect(0, y0, w, y1)
    dst = fitz.open()
    new_page = dst.new_page(width=w, height=h / 3)
    new_page.show_pdf_page(new_page.rect, src, page_num - 1, clip=clip)
    data = dst.tobytes()
    src.close()
    dst.close()
    return data


# ─── تحميل عند البدء ─────────────────────────────────────────
try:
    EMPLOYEES = load_employees()
    PDF_MAP   = build_pdf_map()
except Exception as e:
    logger.critical(f"خطأ في تحميل البيانات: {e}")
    raise

user_state: dict = {}


# ─── هاندلرز ──────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً!\n\nأرسل *كود الموظف* الخاص بك:",
        parse_mode="Markdown"
    )
    return ASK_CODE


async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import time
    user_id = update.effective_user.id
    code = update.message.text.strip()

    state = user_state.get(user_id, {})
    blocked_until = state.get('blocked_until', 0)
    if time.time() < blocked_until:
        remaining = int(blocked_until - time.time())
        await update.message.reply_text(
            f"⛔ تم تجاوز عدد المحاولات.\n"
            f"يرجى الانتظار {remaining} ثانية والمحاولة مجددًا."
        )
        return ConversationHandler.END

    if code not in EMPLOYEES:
        logger.warning(f"User {user_id} entered invalid code: {code}")
        await update.message.reply_text("❌ الكود غير صحيح. حاول مرة أخرى أو تواصل مع HR.")
        return ConversationHandler.END

    context.user_data['code'] = code
    await update.message.reply_text("🔒 أرسل *الرقم السري* الخاص بك:", parse_mode="Markdown")
    return ASK_PIN


async def receive_pin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import time
    user_id = update.effective_user.id
    pin_entered = update.message.text.strip()
    code = context.user_data.get('code')

    if not code:
        await update.message.reply_text("حدث خطأ. ابدأ من جديد بـ /start")
        return ConversationHandler.END

    correct_pin = EMPLOYEES[code]['pin']
    state = user_state.setdefault(user_id, {'tries': 0, 'blocked_until': 0})

    if pin_entered != correct_pin:
        state['tries'] += 1
        remaining_tries = MAX_TRIES - state['tries']
        logger.warning(f"User {user_id} wrong PIN for code {code} (attempt {state['tries']})")

        if state['tries'] >= MAX_TRIES:
            state['blocked_until'] = time.time() + BLOCK_SECS
            state['tries'] = 0
            await update.message.reply_text(
                f"⛔ تم تجاوز {MAX_TRIES} محاولات خاطئة.\n"
                f"تم تعليق حسابك مؤقتاً لمدة {BLOCK_SECS // 60} دقائق."
            )
            return ConversationHandler.END

        await update.message.reply_text(
            f"❌ الرقم السري غير صحيح.\n"
            f"لديك {remaining_tries} محاولة{'ات' if remaining_tries > 1 else ''} متبقية."
        )
        return ASK_PIN

    state['tries'] = 0

    if code not in PDF_MAP:
        logger.error(f"Code {code} not found in PDF map")
        await update.message.reply_text(
            "✅ تم التحقق بنجاح، لكن لم يتم العثور على فيش القبض الخاص بك.\n"
            "يرجى التواصل مع HR."
        )
        return ConversationHandler.END

    page_num, position = PDF_MAP[code]
    name = EMPLOYEES[code]['name']
    logger.info(f"User {user_id} - Code {code} - Name {name} - Page {page_num} Pos {position} - SENT")

    await update.message.reply_text(f"✅ مرحباً {name}!\nجاري إرسال فيش القبض...")

    try:
        pdf_bytes = extract_slip(page_num, position)
        await update.message.reply_document(
            document=io.BytesIO(pdf_bytes),
            filename=f"فيش_القبض_{code}.pdf",
            caption="📄 فيش القبض الخاص بك"
        )
    except Exception as e:
        logger.error(f"Error sending slip for code {code}: {e}")
        await update.message.reply_text("حدث خطأ أثناء إرسال فيش القبض. يرجى المحاولة لاحقاً.")

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("تم الإلغاء. ابدأ من جديد بـ /start")
    return ConversationHandler.END


# ─── تشغيل البوت ──────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASK_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_code)],
            ASK_PIN:  [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_pin)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)
    logger.info("البوت يعمل الآن...")
    print("✅ البوت شغال! اضغط Ctrl+C لإيقافه.")
    app.run_polling()


if __name__ == "__main__":
    main()
