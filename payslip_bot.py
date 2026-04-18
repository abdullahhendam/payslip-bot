"""
بوت تيليجرام لإرسال فيش القبض
"""

import re
import io
import logging
import pandas as pd
import pdfplumber
import fitz
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)

# ─── إعدادات ────────────────────────────────────────────────
BOT_TOKEN  = "8743584646:AAHN2TIMN47GkMLHayZSy4RK0EkdxQ8ssB8"
EXCEL_FILE = "employees.xlsx"
PDF_FILE   = "payslips.pdf"
MAX_TRIES  = 3
BLOCK_SECS = 300
ADMIN_IDS  = [1802415105]  # حط الـ Telegram ID بتاعك هنا مثلاً: [123456789]

# ─── مراحل المحادثة ──────────────────────────────────────────
CHOOSE, ASK_CODE, ASK_PIN, ASK_CONTACT, ADMIN_CONTACT = range(5)

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
        raise ValueError("لم يتم العثور على هيدر الجدول")
    df = pd.read_excel(EXCEL_FILE, header=header_row)
    employees = {}
    phone_map = {}  # رقم الموبايل -> كود الموظف
    for _, row in df.iterrows():
        raw_code = row.get('الكود', '')
        raw_pin  = row.get('الرقم السري', '')
        raw_name = row.get('الاسم', '')
        raw_phone = row.get('رقم الموبايل', '')
        try:
            code = str(int(float(str(raw_code)))).strip()
        except:
            code = str(raw_code).strip()
        try:
            pin = str(int(float(str(raw_pin)))).strip()
        except:
            pin = str(raw_pin).strip()
        name = str(raw_name).strip()
        phone = str(raw_phone).strip().replace(' ', '').replace('-', '')
        if code and code != 'nan' and pin and pin != 'nan':
            employees[code] = {'pin': pin, 'name': name, 'phone': phone}
            if phone and phone != 'nan':
                # حفظ بأشكال مختلفة للرقم
                phone_map[phone] = code
                if phone.startswith('0'):
                    phone_map['2' + phone[1:]] = code
                    phone_map['+2' + phone[1:]] = code
                if phone.startswith('20'):
                    phone_map['0' + phone[2:]] = code
                if phone.startswith('+20'):
                    phone_map['0' + phone[3:]] = code
    logger.info(f"تم تحميل {len(employees)} موظف")
    return employees, phone_map


def build_pdf_map():
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
                codes = re.findall(r'\b([0-9]{3,5})\b', text)
                for code in codes:
                    if code != '2026' and len(code) >= 3 and code not in emp_map:
                        emp_map[code] = (page_num, pos)
    logger.info(f"تم بناء خريطة PDF: {len(emp_map)} موظف")
    return emp_map


def extract_slip(page_num: int, position: int) -> bytes:
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


async def send_slip_by_code(update: Update, code: str):
    """إرسال فيش القبض بناءً على الكود"""
    if code not in PDF_MAP:
        await update.message.reply_text(
            "✅ تم التحقق، لكن لم يتم العثور على فيش القبض.\nتواصل مع HR.",
            reply_markup=ReplyKeyboardRemove()
        )
        return
    page_num, position = PDF_MAP[code]
    name = EMPLOYEES[code]['name']
    await update.message.reply_text(
        f"✅ مرحباً {name}!\nجاري إرسال فيش القبض...",
        reply_markup=ReplyKeyboardRemove()
    )
    try:
        pdf_bytes = extract_slip(page_num, position)
        await update.message.reply_document(
            document=io.BytesIO(pdf_bytes),
            filename=f"فيش_القبض_{code}.pdf",
            caption="📄 فيش القبض الخاص بك"
        )
    except Exception as e:
        logger.error(f"Error sending slip for code {code}: {e}")
        await update.message.reply_text("حدث خطأ أثناء إرسال فيش القبض. حاول لاحقاً.")


# ─── تحميل عند البدء ─────────────────────────────────────────
try:
    EMPLOYEES, PHONE_MAP = load_employees()
    PDF_MAP = build_pdf_map()
except Exception as e:
    logger.critical(f"خطأ: {e}")
    raise

user_state: dict = {}


# ─── هاندلرز ──────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    keyboard = [
        ["📄 استلام فيش القبض"],
        ["🔑 معرفة الكود وكلمة السر"]
    ]
    # لو أدمن، يظهر له خيار إضافي
    if user_id in ADMIN_IDS:
        keyboard.append(["👤 عرض بيانات موظف (أدمن)"])

    await update.message.reply_text(
        "👋 أهلاً!\nاختار الخدمة اللي تريدها:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return CHOOSE


async def choose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    choice = update.message.text.strip()

    if choice == "📄 استلام فيش القبض":
        await update.message.reply_text(
            "أرسل *كود الموظف* الخاص بك:",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove()
        )
        return ASK_CODE

    elif choice == "🔑 معرفة الكود وكلمة السر":
        btn = KeyboardButton("📱 مشاركة رقم موبايلي", request_contact=True)
        await update.message.reply_text(
            "اضغط الزر عشان نتعرف عليك من رقم موبايلك:",
            reply_markup=ReplyKeyboardMarkup([[btn]], resize_keyboard=True, one_time_keyboard=True)
        )
        return ASK_CONTACT

    elif choice == "👤 عرض بيانات موظف (أدمن)" and user_id in ADMIN_IDS:
        btn = KeyboardButton("📱 مشاركة رقم الموظف", request_contact=True)
        await update.message.reply_text(
            "شارك رقم الموظف اللي عايز تشوف بياناته:",
            reply_markup=ReplyKeyboardMarkup([[btn]], resize_keyboard=True, one_time_keyboard=True)
        )
        return ADMIN_CONTACT

    else:
        await update.message.reply_text("اختار من القائمة.")
        return CHOOSE


async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import time
    user_id = update.effective_user.id
    code = update.message.text.strip()

    state = user_state.get(user_id, {})
    if time.time() < state.get('blocked_until', 0):
        remaining = int(state['blocked_until'] - time.time())
        await update.message.reply_text(f"⛔ محظور. انتظر {remaining} ثانية.")
        return ConversationHandler.END

    if code not in EMPLOYEES:
        await update.message.reply_text("❌ الكود غير صحيح. تواصل مع HR.")
        return ConversationHandler.END

    context.user_data['code'] = code
    await update.message.reply_text("🔒 أرسل *الرقم السري*:", parse_mode="Markdown")
    return ASK_PIN


async def receive_pin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import time
    user_id = update.effective_user.id
    pin_entered = update.message.text.strip()
    code = context.user_data.get('code')

    if not code:
        await update.message.reply_text("حدث خطأ. ابدأ بـ /start")
        return ConversationHandler.END

    correct_pin = EMPLOYEES[code]['pin']
    state = user_state.setdefault(user_id, {'tries': 0, 'blocked_until': 0})

    if pin_entered != correct_pin:
        state['tries'] += 1
        remaining_tries = MAX_TRIES - state['tries']
        if state['tries'] >= MAX_TRIES:
            state['blocked_until'] = time.time() + BLOCK_SECS
            state['tries'] = 0
            await update.message.reply_text(
                f"⛔ تجاوزت {MAX_TRIES} محاولات. محظور لمدة {BLOCK_SECS // 60} دقائق."
            )
            return ConversationHandler.END
        await update.message.reply_text(
            f"❌ الرقم السري غير صحيح. لديك {remaining_tries} محاولة متبقية."
        )
        return ASK_PIN

    state['tries'] = 0
    logger.info(f"User {user_id} - Code {code} - SLIP SENT")
    await send_slip_by_code(update, code)
    return ConversationHandler.END


async def receive_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """معرفة الكود وكلمة السر عن طريق رقم الموبايل"""
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("لم يتم استلام رقم الموبايل. حاول مرة أخرى.")
        return ConversationHandler.END

    phone = contact.phone_number.replace('+', '').replace(' ', '')
    # تجربة أشكال مختلفة للرقم
    code = None
    for variant in [phone, '0' + phone[-9:], '20' + phone[-9:], '+20' + phone[-9:]]:
        if variant in PHONE_MAP:
            code = PHONE_MAP[variant]
            break

    if not code:
        await update.message.reply_text(
            "❌ رقم موبايلك مش موجود في السجلات.\nتواصل مع HR.",
            reply_markup=ReplyKeyboardRemove()
        )
        return ConversationHandler.END

    emp = EMPLOYEES[code]
    logger.info(f"Contact lookup - Phone {phone} - Code {code} - Name {emp['name']}")
    await update.message.reply_text(
        f"✅ تم التعرف عليك!\n\n"
        f"👤 الاسم: {emp['name']}\n"
        f"🔢 الكود: {code}\n"
        f"🔑 الرقم السري: {emp['pin']}",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


async def admin_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """الأدمن يشوف بيانات أي موظف"""
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("غير مصرح.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    contact = update.message.contact
    if not contact:
        await update.message.reply_text("لم يتم استلام رقم الموبايل.")
        return ConversationHandler.END

    phone = contact.phone_number.replace('+', '').replace(' ', '')
    code = None
    for variant in [phone, '0' + phone[-9:], '20' + phone[-9:], '+20' + phone[-9:]]:
        if variant in PHONE_MAP:
            code = PHONE_MAP[variant]
            break

    if not code:
        await update.message.reply_text(
            "❌ الرقم مش موجود في السجلات.",
            reply_markup=ReplyKeyboardRemove()
        )
        return ConversationHandler.END

    emp = EMPLOYEES[code]
    await update.message.reply_text(
        f"📋 بيانات الموظف:\n\n"
        f"👤 الاسم: {emp['name']}\n"
        f"🔢 الكود: {code}\n"
        f"🔑 الرقم السري: {emp['pin']}\n"
        f"📱 الموبايل: {emp['phone']}",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("تم الإلغاء.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


# ─── تشغيل البوت ──────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            CHOOSE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, choose)],
            ASK_CODE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_code)],
            ASK_PIN:       [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_pin)],
            ASK_CONTACT:   [MessageHandler(filters.CONTACT, receive_contact)],
            ADMIN_CONTACT: [MessageHandler(filters.CONTACT, admin_contact)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)
    logger.info("البوت يعمل...")
    print("✅ البوت شغال!")
    app.run_polling()


if __name__ == "__main__":
    main()
