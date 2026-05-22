import asyncio
import base64
import io
import json
import logging
import os
import re
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import BufferedInputFile, Message
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook


load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
EXCEL_PATH = DATA_DIR / "customers.xlsx"

HEADERS = [
    "No",
    "Ism familiya",
    "Telefon raqami",
    "Manzil",
    "Qo'shimcha izoh",
    "Tekshirish kerak",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

excel_lock = asyncio.Lock()
openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None


CUSTOMER_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "customers": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "number": {"type": "string"},
                    "full_name": {"type": "string"},
                    "phone": {"type": "string"},
                    "address": {"type": "string"},
                    "note": {"type": "string"},
                    "needs_review": {"type": "string"},
                },
                "required": [
                    "number",
                    "full_name",
                    "phone",
                    "address",
                    "note",
                    "needs_review",
                ],
            },
        }
    },
    "required": ["customers"],
}


SYSTEM_PROMPT = """
Siz mijoz ma'lumotlarini ajratadigan yordamchisiz.
Matn yoki rasmda bir nechta mijoz bo'lishi mumkin. Har bir mijozni alohida obyekt qiling.

Ajratiladigan maydonlar:
- number: asl tartib raqami bor bo'lsa, aks holda bo'sh string
- full_name: ism familiya bor bo'lsa
- phone: telefon raqami asl matndagi ko'rinishida, hech narsa to'qimang
- address: manzil
- note: boshqa foydali izohlar, noaniq yoki yo'qolmasligi kerak bo'lgan bo'laklar
- needs_review: noaniq o'qilgan, telefon raqami shubhali, rasm sifati past, yoki maydonlar aralash bo'lsa qisqa izoh

Qoidalar:
- Ma'lumot yo'q bo'lsa bo'sh string qaytaring.
- Telefon raqamni formatlamang, asl ko'rinishida qaytaring.
- Taxmin qilmang. Ishonchsiz joylarni needs_review maydoniga yozing.
- Javob faqat schema bo'yicha bo'lsin.
""".strip()


def ensure_excel_file() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Customers"
    sheet.append(HEADERS)

    widths = [8, 28, 18, 38, 40, 32]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width

    workbook.save(EXCEL_PATH)


def reset_excel_file() -> None:
    if EXCEL_PATH.exists():
        EXCEL_PATH.unlink()
    ensure_excel_file()


def normalize_phone(raw_phone: str) -> tuple[str, str]:
    raw_phone = (raw_phone or "").strip()
    if not raw_phone:
        return "", ""

    digits = re.sub(r"\D", "", raw_phone)
    review = ""

    if len(digits) == 12 and digits.startswith("998"):
        return digits, review

    if len(digits) == 9:
        return f"998{digits}", review

    if len(digits) == 10 and digits.startswith("0"):
        return f"998{digits[1:]}", review

    return raw_phone, f"Telefon noaniq: {raw_phone}"


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def prepare_rows(customers: list[dict[str, Any]]) -> list[list[str]]:
    rows = []
    for customer in customers:
        normalized_phone, phone_review = normalize_phone(clean_text(customer.get("phone")))
        review_parts = [
            clean_text(customer.get("needs_review")),
            phone_review,
        ]
        review = "; ".join(part for part in review_parts if part)

        rows.append(
            [
                clean_text(customer.get("number")),
                clean_text(customer.get("full_name")),
                normalized_phone,
                clean_text(customer.get("address")),
                clean_text(customer.get("note")),
                review,
            ]
        )
    return rows


async def append_customers(customers: list[dict[str, Any]]) -> int:
    rows = prepare_rows(customers)
    if not rows:
        return 0

    async with excel_lock:
        ensure_excel_file()
        workbook = load_workbook(EXCEL_PATH)
        sheet = workbook.active

        next_number = sheet.max_row
        for row in rows:
            supplied_number = row[0]
            row[0] = supplied_number if supplied_number else str(next_number)
            sheet.append(row)
            next_number += 1

        workbook.save(EXCEL_PATH)

    return len(rows)


async def get_excel_bytes() -> bytes:
    async with excel_lock:
        ensure_excel_file()
        return EXCEL_PATH.read_bytes()


def parse_openai_output(response: Any) -> list[dict[str, Any]]:
    output_text = getattr(response, "output_text", "")
    if not output_text:
        raise ValueError("OpenAI bo'sh javob qaytardi.")

    parsed = json.loads(output_text)
    customers = parsed.get("customers", [])
    if not isinstance(customers, list):
        raise ValueError("OpenAI javobida customers ro'yxati topilmadi.")

    return [item for item in customers if isinstance(item, dict)]


def call_openai_with_text(text: str) -> list[dict[str, Any]]:
    if openai_client is None:
        raise RuntimeError("OPENAI_API_KEY environment variable sozlanmagan.")

    response = openai_client.responses.create(
        model=OPENAI_MODEL,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": f"Quyidagi mijoz ma'lumotlarini ajrating:\n\n{text}",
            },
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "customer_extraction",
                "schema": CUSTOMER_SCHEMA,
                "strict": True,
            }
        },
    )
    return parse_openai_output(response)


def call_openai_with_image(image_bytes: bytes, mime_type: str) -> list[dict[str, Any]]:
    if openai_client is None:
        raise RuntimeError("OPENAI_API_KEY environment variable sozlanmagan.")

    encoded = base64.b64encode(image_bytes).decode("utf-8")
    data_url = f"data:{mime_type};base64,{encoded}"

    response = openai_client.responses.create(
        model=OPENAI_MODEL,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": "Rasmdagi mijoz ma'lumotlarini maksimal aniqlik bilan o'qing va ajrating.",
                    },
                    {
                        "type": "input_image",
                        "image_url": data_url,
                    },
                ],
            },
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "customer_extraction",
                "schema": CUSTOMER_SCHEMA,
                "strict": True,
            }
        },
    )
    return parse_openai_output(response)


async def handle_customers(message: Message, customers: list[dict[str, Any]]) -> None:
    if not customers:
        await message.answer(
            "Mijoz ma'lumotlari topilmadi. Iltimos, matnni aniqroq yuboring yoki rasm sifatini yaxshilang."
        )
        return

    count = await append_customers(customers)
    await message.answer(
        f"{count} ta mijoz Excel faylga qo'shildi.\n"
        "Faylni olish uchun /excel buyrug'ini yuboring."
    )


async def start_handler(message: Message) -> None:
    await message.answer(
        "Assalomu alaykum!\n\n"
        "Men mijoz ma'lumotlarini matn yoki rasm ichidan ajratib, Excel faylga yozib boraman.\n\n"
        "Yuborishingiz mumkin:\n"
        "- oddiy matn\n"
        "- daftar rasmi\n"
        "- skrinshot\n"
        "- qo'lda yozilgan ma'lumot rasmi\n\n"
        "Excel faylni olish: /excel\n"
        "Ro'yxatni tozalash: /clear\n"
        "Yordam: /help"
    )


async def help_handler(message: Message) -> None:
    await message.answer(
        "Foydalanish:\n\n"
        "1. Mijoz ma'lumotlarini matn qilib yuboring yoki rasm jo'nating.\n"
        "2. Bot ism, telefon, manzil va izohlarni ajratadi.\n"
        "3. Telefonlar 998XXXXXXXXX formatiga keltiriladi.\n"
        "4. Noaniq joylar 'Tekshirish kerak' ustuniga yoziladi.\n\n"
        "Komandalar:\n"
        "/excel - Excel faylni yuboradi\n"
        "/clear - ro'yxatni tozalaydi\n"
        "/help - yordam"
    )


async def excel_handler(message: Message) -> None:
    file_bytes = await get_excel_bytes()
    await message.answer_document(
        BufferedInputFile(file_bytes, filename="customers.xlsx"),
        caption="Yangilangan mijozlar ro'yxati.",
    )


async def clear_handler(message: Message) -> None:
    async with excel_lock:
        reset_excel_file()
    await message.answer("Ro'yxat tozalandi. Yangi Excel fayl tayyor.")


async def text_handler(message: Message) -> None:
    text = message.text or ""
    processing = await message.answer("Matn tahlil qilinyapti...")

    try:
        customers = await asyncio.to_thread(call_openai_with_text, text)
        await handle_customers(message, customers)
    except Exception as error:
        logger.exception("Text parsing failed")
        await message.answer(f"Xatolik yuz berdi: {error}")
    finally:
        try:
            await processing.delete()
        except Exception:
            pass


async def photo_handler(message: Message, bot: Bot) -> None:
    processing = await message.answer("Rasm o'qilyapti va tahlil qilinyapti...")

    try:
        photo = message.photo[-1]
        file = await bot.get_file(photo.file_id)
        buffer = io.BytesIO()
        await bot.download_file(file.file_path, destination=buffer)
        image_bytes = buffer.getvalue()

        customers = await asyncio.to_thread(
            call_openai_with_image,
            image_bytes,
            "image/jpeg",
        )
        await handle_customers(message, customers)
    except Exception as error:
        logger.exception("Image parsing failed")
        await message.answer(f"Xatolik yuz berdi: {error}")
    finally:
        try:
            await processing.delete()
        except Exception:
            pass


async def document_image_handler(message: Message, bot: Bot) -> None:
    document = message.document
    if document is None or not (document.mime_type or "").startswith("image/"):
        await message.answer("Iltimos, rasm yoki mijoz ma'lumotlari yozilgan matn yuboring.")
        return

    processing = await message.answer("Rasm fayli o'qilyapti va tahlil qilinyapti...")

    try:
        file = await bot.get_file(document.file_id)
        buffer = io.BytesIO()
        await bot.download_file(file.file_path, destination=buffer)
        image_bytes = buffer.getvalue()

        customers = await asyncio.to_thread(
            call_openai_with_image,
            image_bytes,
            document.mime_type or "image/jpeg",
        )
        await handle_customers(message, customers)
    except Exception as error:
        logger.exception("Document image parsing failed")
        await message.answer(f"Xatolik yuz berdi: {error}")
    finally:
        try:
            await processing.delete()
        except Exception:
            pass


async def unsupported_handler(message: Message) -> None:
    await message.answer("Matn yoki rasm yuboring. Yordam uchun /help buyrug'ini bosing.")


async def main() -> None:
    if not BOT_TOKEN:
        raise RuntimeError(
            "BOT_TOKEN environment variable sozlanmagan. Railway > Variables bo'limiga BOT_TOKEN qo'shing."
        )
    if not OPENAI_API_KEY:
        raise RuntimeError(
            "OPENAI_API_KEY environment variable sozlanmagan. Railway > Variables bo'limiga OPENAI_API_KEY qo'shing."
        )

    ensure_excel_file()

    bot = Bot(token=BOT_TOKEN)
    dispatcher = Dispatcher()

    dispatcher.message.register(start_handler, Command("start"))
    dispatcher.message.register(help_handler, Command("help"))
    dispatcher.message.register(excel_handler, Command("excel"))
    dispatcher.message.register(clear_handler, Command("clear"))
    dispatcher.message.register(photo_handler, F.photo)
    dispatcher.message.register(document_image_handler, F.document)
    dispatcher.message.register(text_handler, F.text)
    dispatcher.message.register(unsupported_handler)

    logger.info("Bot started")
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot stopped")
