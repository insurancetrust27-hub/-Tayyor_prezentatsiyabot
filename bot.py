import asyncio
import json
import time

from google import genai
from google.genai import types
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from kalitlar import ADMIN_ID, GEMINI_KALIT, KARTA, NARX, TOKEN
from shablon import SHABLONLAR, prompt_yaratish, shablondan_yaratish

MODEL = "gemini-3.6-flash"
client = genai.Client(api_key=GEMINI_KALIT)

NARX_MATN = f"{NARX:,}".replace(",", " ")
BUYURTMALAR = {}


# ---------------------------------------------------------------------------
# AI: shablon uchun matn yozish
# ---------------------------------------------------------------------------

def matn_generatsiya(shablon_kaliti, mavzu):
    prompt = prompt_yaratish(shablon_kaliti, mavzu)
    oxirgi_xato = None
    for urinish in range(4):
        try:
            javob = client.models.generate_content(
                model=MODEL,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                ),
            )
            malumot = json.loads(javob.text)
            if not isinstance(malumot, dict):
                raise ValueError("AI javobi JSON obyekt emas")
            return malumot
        except Exception as xato:
            oxirgi_xato = xato
            print("Urinish", urinish + 1, "muvaffaqiyatsiz:", xato)
            time.sleep(5 * (urinish + 1))
    raise oxirgi_xato


# ---------------------------------------------------------------------------
# Mijoz bilan suhbat: mavzu -> shablon -> to'lov
# ---------------------------------------------------------------------------

async def boshlash(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Salom! Taqdimot mavzusini yozing, men uni tayyorlab beraman.\n"
        f"Bitta taqdimot narxi: {NARX_MATN} so'm."
    )


async def mavzu_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mavzu = update.message.text.strip()
    context.user_data["mavzu"] = mavzu
    context.user_data.pop("shablon", None)

    tugmalar = InlineKeyboardMarkup(
        [[InlineKeyboardButton(v["nomi"], callback_data=f"sh:{k}")]
         for k, v in SHABLONLAR.items()]
    )
    await update.message.reply_text(
        f"Mavzu: {mavzu}\n\nQaysi dizayn shablonini tanlaysiz?",
        reply_markup=tugmalar,
    )


async def shablon_tanlandi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    mavzu = context.user_data.get("mavzu")
    if not mavzu:
        await q.edit_message_text("Avval taqdimot mavzusini yozing.")
        return
    kalit = q.data.split(":")[1]
    if kalit not in SHABLONLAR:
        await q.edit_message_text("Bunday shablon topilmadi. Mavzuni qayta yozing.")
        return
    context.user_data["shablon"] = kalit

    await q.edit_message_text(
        f"Mavzu: {mavzu}\n"
        f"Shablon: {SHABLONLAR[kalit]['nomi']}\n\n"
        f"Narxi: {NARX_MATN} so'm\n"
        f"Karta: {KARTA}\n\n"
        "To'lovni o'tkazing va chekning skrinshotini shu yerga rasm qilib yuboring. "
        "To'lov tasdiqlangach, taqdimot tayyorlanadi."
    )


# ---------------------------------------------------------------------------
# Chek qabul qilish va tasdiqlash
# ---------------------------------------------------------------------------

async def chek_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ud = context.user_data
    if not ud.get("mavzu") or not ud.get("shablon"):
        await update.message.reply_text(
            "Avval taqdimot mavzusini yozing va shablonni tanlang."
        )
        return

    buyurtma_id = str(int(time.time() * 1000))
    BUYURTMALAR[buyurtma_id] = {
        "chat_id": update.effective_chat.id,
        "mijoz": update.effective_user.full_name,
        "username": update.effective_user.username,
        "mavzu": ud["mavzu"],
        "shablon": ud["shablon"],
        "holat": "kutmoqda",
    }
    for kalit in ("mavzu", "shablon"):
        ud.pop(kalit, None)

    b = BUYURTMALAR[buyurtma_id]
    tugmalar = InlineKeyboardMarkup([[
        InlineKeyboardButton("Tasdiqlash", callback_data="ok:" + buyurtma_id),
        InlineKeyboardButton("Rad etish", callback_data="no:" + buyurtma_id),
    ]])
    await context.bot.send_photo(
        chat_id=ADMIN_ID,
        photo=update.message.photo[-1].file_id,
        caption=(
            "Yangi buyurtma\n"
            f"ID: {buyurtma_id}\n"
            f"Mijoz: {b['mijoz']}\n"
            f"Mavzu: {b['mavzu']}\n"
            f"Shablon: {SHABLONLAR[b['shablon']]['nomi']}\n"
            f"Summa: {NARX_MATN} so'm"
        ),
        reply_markup=tugmalar,
    )
    await update.message.reply_text(
        "Chek qabul qilindi. Tasdiqlanishini kuting, biroz vaqt olishi mumkin."
    )


async def taqdimot_yasab_yuborish(context: ContextTypes.DEFAULT_TYPE, buyurtma_id, b):
    """AI matnini yozdiradi, shablonni to'ldiradi va mijozga yuboradi.
    Xato bo'lsa, adminga to'liq ma'lumot bilan xabar beradi."""
    try:
        kontent = await asyncio.to_thread(matn_generatsiya, b["shablon"], b["mavzu"])
        fayl = await asyncio.to_thread(shablondan_yaratish, b["shablon"], kontent)
        await context.bot.send_document(
            b["chat_id"], document=fayl, filename="taqdimot.pptx"
        )
        b["holat"] = "tayyor"
    except Exception as xato:
        print("XATO:", xato)
        b["holat"] = "xato"
        await context.bot.send_message(
            b["chat_id"],
            "Kechirasiz, taqdimot yasashda xatolik yuz berdi. Tez orada siz bilan bog'lanamiz.",
        )
        username_matn = f"@{b['username']}" if b.get("username") else "(username yo'q)"
        await context.bot.send_message(
            ADMIN_ID,
            "Xato: buyurtma tayyorlanmadi.\n"
            f"Buyurtma ID: {buyurtma_id}\n"
            f"Mijoz: {b['mijoz']} {username_matn}\n"
            f"Chat ID: {b['chat_id']}\n"
            f"Mavzu: {b['mavzu']}\n\n"
            f"Qayta urinish uchun: /retry {buyurtma_id}\n"
            f"Xato matni: {xato}",
        )


async def tugma_bosildi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if str(q.from_user.id) != str(ADMIN_ID):
        await q.answer("Ruxsat yo'q", show_alert=True)
        return
    amal, buyurtma_id = q.data.split(":")
    b = BUYURTMALAR.get(buyurtma_id)
    if not b or b["holat"] != "kutmoqda":
        await q.answer("Bu buyurtma ko'rib chiqilgan yoki topilmadi.", show_alert=True)
        return
    await q.answer()
    asl_matn = q.message.caption or ""

    if amal == "no":
        b["holat"] = "rad"
        await q.edit_message_caption(caption=asl_matn + "\n\nRAD ETILDI")
        await context.bot.send_message(
            b["chat_id"],
            "Kechirasiz, to'lov tasdiqlanmadi. Chekni tekshirib, qayta urinib ko'ring.",
        )
        return

    b["holat"] = "tayyorlanmoqda"
    await q.edit_message_caption(caption=asl_matn + "\n\nTASDIQLANDI, tayyorlanmoqda...")
    await context.bot.send_message(
        b["chat_id"], "To'lov tasdiqlandi. Taqdimot tayyorlanmoqda, biroz kuting..."
    )
    await taqdimot_yasab_yuborish(context, buyurtma_id, b)


async def qayta_urinish(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin uchun: /retry <buyurtma_id> - xato bo'lgan buyurtmani qayta yasaydi."""
    if str(update.effective_user.id) != str(ADMIN_ID):
        return
    if not context.args:
        await update.message.reply_text("Foydalanish: /retry <buyurtma_id>")
        return
    buyurtma_id = context.args[0]
    b = BUYURTMALAR.get(buyurtma_id)
    if not b:
        await update.message.reply_text("Bunday buyurtma topilmadi.")
        return
    await update.message.reply_text(f"Buyurtma {buyurtma_id} qayta urinilmoqda...")
    b["holat"] = "tayyorlanmoqda"
    await taqdimot_yasab_yuborish(context, buyurtma_id, b)
    if b["holat"] == "tayyor":
        await update.message.reply_text("Muvaffaqiyatli yuborildi.")


def main():
    app = ApplicationBuilder().token(TOKEN).concurrent_updates(True).build()
    app.add_handler(CommandHandler("start", boshlash))
    app.add_handler(CommandHandler("retry", qayta_urinish))
    app.add_handler(CallbackQueryHandler(shablon_tanlandi, pattern=r"^sh:\w+$"))
    app.add_handler(CallbackQueryHandler(tugma_bosildi, pattern=r"^(ok|no):\d+$"))
    app.add_handler(MessageHandler(filters.PHOTO, chek_qabul))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mavzu_qabul))
    app.run_polling()


if __name__ == "__main__":
    main()
