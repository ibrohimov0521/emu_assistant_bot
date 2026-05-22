import asyncio
import base64
import io
import json
import logging
import os
import re
import shutil
from difflib import get_close_matches
from copy import copy
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import BotCommand, BufferedInputFile, Message
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook


load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "templates"
TEMPLATE_PATH = TEMPLATE_DIR / "yangi_shablon.xlsx"
EXCEL_PATH = DATA_DIR / "customers.xlsx"

HEADERS = [
    "Номер",
    "Компания-получатель",
    "ФИО получателя",
    "Адрес получателя",
    "Телефон получателя",
    "Шифр клиента",
    "Масса посылки",
    "Поручение",
    "Количество мест",
    "Штрихкод (№ накладной)",
    "Компания-отправитель",
    "ФИО отправителя",
    "Адрес отправителя",
    "Телефон отправителя",
    "Город-отправитель",
    "Город-получатель",
    "Оплата получателем",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

excel_lock = asyncio.Lock()
openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
sender_sessions: dict[int, dict[str, str]] = {}
setup_states: dict[int, dict[str, Any]] = {}


SETUP_STEPS = [
    (
        "sender_full_name",
        "Jo'natuvchining ism familiyasini kiriting.",
    ),
    (
        "sender_phone",
        "Jo'natuvchining telefon raqamini kiriting. Masalan: +998 90 123 45 67",
    ),
    (
        "sender_address",
        "Jo'natuvchining to'liq manzilini kiriting.",
    ),
    (
        "sender_city_ru",
        "Jo'natuvchining shahrini rus tilida kiriting. Masalan: Ташкент, Бухара, Шафиркан",
    ),
    (
        "cipher_prefix",
        "Jo'natmalar uchun qaytarilmaydigan shifr prefixini kiriting. Masalan: ABC",
    ),
    (
        "payment_by_receiver",
        "Оплата получателем bo'ladimi? True yoki False deb yozing.",
    ),
    (
        "parcel_weight",
        "Jo'natma og'irligini kiriting. Masalan: 1.5",
    ),
    (
        "places_count",
        "Bir mijozga nechta jo'natma bo'lishini kiriting. Masalan: 1",
    ),
]

ALLOWED_RECIPIENT_LOCATIONS = [
    "Ташкент",
    "Алмазар",
    "Бектемир",
    "Мирабад",
    "Мирзо-Улугбек",
    "Сергели",
    "Учтепа",
    "Чиланзар",
    "Шайхантахур",
    "Юнусабад",
    "Яккасарай",
    "Янгихаёт",
    "Яшнабад",
    "Андижан",
    "Алтинкул",
    "Асака",
    "Балыкчи",
    "Боз",
    "Булокбоши",
    "Джалакудук",
    "Избаскан",
    "Куйган - яр",
    "Кургантепа",
    "Мархамат",
    "Пахтаабад",
    "Улугнор",
    "Ханабад",
    "Ходжаабад",
    "Шахрихан",
    "Бухара",
    "Алат",
    "Вабкент",
    "Галлаасия",
    "Гиждуван",
    "Джандар",
    "Каган",
    "Каракуль",
    "Караулбазар",
    "Пешку",
    "Ромитан",
    "Шафиркан",
    "Джизак",
    "Арнасай",
    "Балангачкыр",
    "Бахмал",
    "Гагарин",
    "Галляарал",
    "Дустлик",
    "Заамин",
    "Зарбдар",
    "Зафарабад",
    "Пахтакор",
    "Шараф-Рашидов",
    "Янгикишлак",
    "Карши",
    "Бешкент",
    "Гузар",
    "Дехканабадский район",
    "Камаши",
    "Касан",
    "Касбий",
    "Китоб",
    "Кокдала",
    "Миришкор",
    "Мубарек",
    "Нишан",
    "Чиракчи",
    "Шахрисабз",
    "Яккабаг",
    "Навои",
    "Бешрабат",
    "Зарафшан",
    "Канимех",
    "Кармана",
    "Кызылтепа",
    "Нурата",
    "Тамдыбулак",
    "Учкудук",
    "Хатырчи",
    "Наманган",
    "Джумашуй",
    "Касансай",
    "Пап",
    "Ташбулак",
    "Туракурган",
    "Уйчи",
    "Учкурган",
    "Хаккулабад",
    "Чартак",
    "Чуст",
    "Янгикурган",
    "Нукус",
    "Акмангит",
    "Амударья",
    "Беруни",
    "Казакеткен",
    "Канлыкуль",
    "Караузяк",
    "Кегейли",
    "Кунград",
    "Муйнак",
    "Тахиаташ",
    "Тахтакупыр",
    "Турткуль",
    "Ходжейли",
    "Чимбай",
    "Шуманай",
    "Элликкала",
    "Самарканд",
    "Акташ",
    "Булунгур",
    "Гульабад",
    "Джамбай",
    "Джума",
    "Зиадин",
    "Иштыхан",
    "Каттакурган",
    "Кушрабад",
    "Лаиш",
    "Нурабад",
    "Пайарык",
    "Тайлак",
    "Ургут",
    "Термез",
    "Ангор",
    "Байсун",
    "Бандихан",
    "Денау",
    "Джаркурган",
    "Карлук",
    "Кизирик",
    "Кумкурган",
    "Сариасия",
    "Узун",
    "Учкизил",
    "Халкабад",
    "Шерабад",
    "Шурчи",
    "Гулистан",
    "Акалтын",
    "Бахт",
    "Дехканабад",
    "Навруз",
    "Сайхун",
    "Сардоба",
    "Сырдарья",
    "Хаваст",
    "Ширин",
    "Янгиер",
    "Нурафшон",
    "Аккурган",
    "Алмалык",
    "Ангрен",
    "Ахангаран",
    "Бекабад",
    "Бука",
    "Верхне-Чирчикский",
    "Газалкент",
    "Дустабад",
    "Зангиата",
    "Келес",
    "Кибрай",
    "Паркент",
    "Пскент",
    "Чиназ",
    "Чирчик",
    "Янгийоль",
    "Фергана",
    "Алтыарык",
    "Багдад",
    "Бешарык",
    "Бувайда",
    "Водил",
    "Дангара",
    "Коканд",
    "Кува",
    "Кувасай",
    "Куштепа",
    "Маргилан",
    "Навбахор",
    "Риштан",
    "Сох",
    "Ташлак",
    "Учкуприк",
    "Язъяван",
    "Яйпан",
    "Ургенч",
    "Багат",
    "Гурлен",
    "Караул",
    "Кошкупыр",
    "Тупраккала",
    "Хазарасп",
    "Ханка",
    "Хива",
    "Шават",
    "Янгиарык",
    "Янгибазар",
]

LOCATION_LIST_FOR_PROMPT = ", ".join(ALLOWED_RECIPIENT_LOCATIONS)


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
                    "recipient_region_ru": {"type": "string"},
                    "note": {"type": "string"},
                    "needs_review": {"type": "string"},
                },
                "required": [
                    "number",
                    "full_name",
                    "phone",
                    "address",
                    "recipient_region_ru",
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
- recipient_region_ru: oluvchining manzilidan P ustun uchun mos shahar/tuman nomini rus tilida tanlang. Faqat quyidagi ro'yxatdan bittasini yozing, boshqa format yozmang:
{location_list}
- note: boshqa foydali izohlar, noaniq yoki yo'qolmasligi kerak bo'lgan bo'laklar
- needs_review: noaniq o'qilgan, telefon raqami shubhali, rasm sifati past, yoki maydonlar aralash bo'lsa qisqa izoh

Qoidalar:
- Ma'lumot yo'q bo'lsa bo'sh string qaytaring.
- Telefon raqamni formatlamang, asl ko'rinishida qaytaring.
- Taxmin qilmang. Ishonchsiz joylarni needs_review maydoniga yozing.
- Ism yo'q bo'lsa "Mijoz", "Noma'lum", "Customer" kabi placeholder yozmang, full_name bo'sh string bo'lsin.
- Bitta xabarda bitta ism/manzil va bir nechta telefon raqami bo'lsa, buni bitta mijoz deb oling: asosiy telefonni phone maydoniga, qolgan telefonlarni note maydoniga yozing.
- Faqat aniq boshqa-boshqa mijozlar bo'lsa alohida obyekt qiling.
- recipient_region_ru hech qachon "Ферганская область, Учкуприкский район" kabi bo'lmasin; ro'yxatdagi "Учкуприк" kabi bitta qiymat bo'lsin.
- Javob faqat schema bo'yicha bo'lsin.
""".strip().format(location_list=LOCATION_LIST_FOR_PROMPT)


def apply_template_header(sheet: Any) -> None:
    if TEMPLATE_PATH.exists():
        template = load_workbook(TEMPLATE_PATH)
        template_sheet = template.active
        for col in range(1, 18):
            source = template_sheet.cell(1, col)
            target = sheet.cell(1, col)
            target.value = source.value
            if source.has_style:
                target._style = copy(source._style)
            target.number_format = source.number_format
            target.alignment = copy(source.alignment)
            letter = target.column_letter
            sheet.column_dimensions[letter].width = template_sheet.column_dimensions[letter].width
        return

    for col, header in enumerate(HEADERS, start=1):
        sheet.cell(1, col).value = header


def ensure_workbook_schema() -> None:
    workbook = load_workbook(EXCEL_PATH)
    sheet = workbook.active
    current_headers = [sheet.cell(1, col).value for col in range(1, 18)]

    if current_headers == HEADERS:
        return

    first_row_has_data = any(value not in (None, "") for value in current_headers)
    if first_row_has_data:
        sheet.insert_rows(1)

    apply_template_header(sheet)
    workbook.save(EXCEL_PATH)


def ensure_excel_file() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        ensure_workbook_schema()
        return

    if TEMPLATE_PATH.exists():
        shutil.copyfile(TEMPLATE_PATH, EXCEL_PATH)
        workbook = load_workbook(EXCEL_PATH)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=17):
            for cell in row:
                cell.value = None
        workbook.save(EXCEL_PATH)
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Шаблон"
    sheet.append(HEADERS)
    widths = [18, 24, 19, 21, 23, 18, 13, 13, 20, 27, 25, 20, 22, 24, 22, 21, 23]
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


def clean_name(value: Any) -> str:
    name = clean_text(value)
    if name.lower() in {"mijoz", "noma'lum", "nomalum", "unknown", "customer"}:
        return ""
    return name


def normalize_location_key(value: str) -> str:
    value = value.lower().replace("ё", "е")
    replacements = {
        "ў": "у",
        "ғ": "г",
        "ģ": "g",
        "ğ": "g",
        "қ": "к",
        "ҳ": "х",
        "ʼ": "",
        "‘": "",
        "’": "",
        "'": "",
        "`": "",
        " - ": "-",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    return re.sub(r"[^a-zа-я0-9]+", "", value)


LOCATION_BY_KEY = {
    normalize_location_key(location): location for location in ALLOWED_RECIPIENT_LOCATIONS
}

LOCATION_ALIASES = {
    "bogot": "Багат",
    "bogat": "Багат",
    "boot": "Багат",
    "bogot tumani": "Багат",
    "bagat": "Багат",
    "olmazor": "Алмазар",
    "almazar": "Алмазар",
    "yunusobod": "Юнусабад",
    "yunusabad": "Юнусабад",
    "sergeli": "Сергели",
    "bekobod": "Бекабад",
    "bekabad": "Бекабад",
    "boka": "Бука",
    "buka": "Бука",
    "buka tumani": "Бука",
    "zafar": "Бекабад",
    "guliston": "Гулистан",
    "gulistan": "Гулистан",
    "tayloq": "Тайлак",
    "taylak": "Тайлак",
    "tayloqtumani": "Тайлак",
    "taylaktumani": "Тайлак",
    "payariq": "Пайарык",
    "payarik": "Пайарык",
    "payaryk": "Пайарык",
    "payariqtumani": "Пайарык",
    "payariktumani": "Пайарык",
    "samarqand": "Самарканд",
    "samarkand": "Самарканд",
    "rudakiy": "Самарканд",
    "qorakol": "Каракуль",
    "qorakul": "Каракуль",
    "karakul": "Каракуль",
    "buxoro": "Бухара",
    "bukhara": "Бухара",
    "shofirkon": "Шафиркан",
    "shafirkan": "Шафиркан",
    "angor": "Ангор",
    "angor tumani": "Ангор",
    "surxondaryo angor": "Ангор",
    "zarbdor": "Зарбдар",
    "zarbdor tumani": "Зарбдар",
    "jizzax zarbdor": "Зарбдар",
    "sharof rashidov": "Шараф-Рашидов",
    "fargona": "Фергана",
    "fergana": "Фергана",
    "uchkoprik": "Учкуприк",
    "uchkuprik": "Учкуприк",
    "uchkopriktumani": "Учкуприк",
    "uchkupraktumani": "Учкуприк",
    "uchkoprik tumani": "Учкуприк",
    "oltiariq": "Алтыарык",
    "oltiriq": "Алтыарык",
    "oltariq": "Алтыарык",
    "oltiarik": "Алтыарык",
    "oltirik": "Алтыарык",
    "oltiriqtumani": "Алтыарык",
    "oltiariqtumani": "Алтыарык",
    "oltiariq tumani": "Алтыарык",
    "altyaryk": "Алтыарык",
    "andijon": "Андижан",
    "andijan": "Андижан",
    "qorgontepa": "Кургантепа",
    "kurgantepa": "Кургантепа",
    "qurgontepa": "Кургантепа",
    "paxtaobod": "Пахтаабад",
    "paxtaabad": "Пахтаабад",
    "pakhtaabad": "Пахтаабад",
}

LOCATION_ALIAS_BY_KEY = {
    normalize_location_key(alias): location for alias, location in LOCATION_ALIASES.items()
}


def resolve_allowed_recipient_location(customer: dict[str, Any]) -> tuple[str, str]:
    address_values = [
        clean_text(customer.get("address")),
        clean_text(customer.get("note")),
    ]
    fallback_values = [
        clean_text(customer.get("recipient_region_ru")),
    ]

    for value in [candidate for candidate in address_values if candidate]:
        key = normalize_location_key(value)
        if key in LOCATION_BY_KEY:
            return LOCATION_BY_KEY[key], ""
        if key in LOCATION_ALIAS_BY_KEY:
            return LOCATION_ALIAS_BY_KEY[key], ""

    address_key = normalize_location_key(" ".join(candidate for candidate in address_values if candidate))
    fallback_key = normalize_location_key(" ".join(candidate for candidate in fallback_values if candidate))
    combined_key = normalize_location_key(" ".join(address_values + fallback_values))

    for alias_key, location in sorted(LOCATION_ALIAS_BY_KEY.items(), key=lambda item: len(item[0]), reverse=True):
        if alias_key and alias_key in address_key:
            return location, ""

    for location_key, location in sorted(LOCATION_BY_KEY.items(), key=lambda item: len(item[0]), reverse=True):
        if location_key and location_key in address_key:
            return location, ""

    for value in [candidate for candidate in fallback_values if candidate]:
        key = normalize_location_key(value)
        if key in LOCATION_BY_KEY:
            return LOCATION_BY_KEY[key], ""
        if key in LOCATION_ALIAS_BY_KEY:
            return LOCATION_ALIAS_BY_KEY[key], ""

    for alias_key, location in sorted(LOCATION_ALIAS_BY_KEY.items(), key=lambda item: len(item[0]), reverse=True):
        if alias_key and alias_key in fallback_key:
            return location, ""

    for location_key, location in sorted(LOCATION_BY_KEY.items(), key=lambda item: len(item[0]), reverse=True):
        if location_key and location_key in fallback_key:
            return location, ""

    matches = get_close_matches(address_key or combined_key, list(LOCATION_BY_KEY.keys()), n=1, cutoff=0.84)
    if matches:
        return LOCATION_BY_KEY[matches[0]], ""

    return "", "P ustun uchun справочникdagi shahar/tuman topilmadi"


def parse_bool(value: str) -> str | None:
    normalized = value.strip().lower()
    true_values = {"true", "1", "ha", "xa", "yes", "y", "да"}
    false_values = {"false", "0", "yo'q", "yoq", "yuq", "no", "n", "нет"}
    if normalized in true_values:
        return "True"
    if normalized in false_values:
        return "False"
    return None


def normalize_cipher_prefix(value: str) -> str:
    prefix = re.sub(r"\s+", "", value.strip()).upper()
    return re.sub(r"[^A-ZА-ЯЁ0-9_-]", "", prefix)


def find_last_data_row(sheet: Any) -> int:
    for row_index in range(sheet.max_row, 1, -1):
        if any(sheet.cell(row_index, col).value not in (None, "") for col in range(1, 18)):
            return row_index
    return 1


def used_cipher_prefixes(sheet: Any) -> set[str]:
    prefixes: set[str] = set()
    for row_index in range(2, sheet.max_row + 1):
        code = clean_text(sheet.cell(row_index, 6).value)
        match = re.match(r"^(.+?)(\d+)$", code)
        if match:
            prefixes.add(match.group(1).upper())
    return prefixes


def next_cipher_index(sheet: Any, prefix: str) -> int:
    highest = 0
    pattern = re.compile(rf"^{re.escape(prefix)}(\d+)$", re.IGNORECASE)
    for row_index in range(2, sheet.max_row + 1):
        code = clean_text(sheet.cell(row_index, 6).value)
        match = pattern.match(code)
        if match:
            highest = max(highest, int(match.group(1)))
    return highest + 1


def copy_row_style(sheet: Any, source_row: int, target_row: int) -> None:
    for col in range(1, 18):
        source = sheet.cell(source_row, col)
        target = sheet.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)


def is_cipher_prefix_available(prefix: str) -> bool:
    ensure_excel_file()
    workbook = load_workbook(EXCEL_PATH)
    sheet = workbook.active
    existing_prefixes = used_cipher_prefixes(sheet)
    active_prefixes = {
        session.get("cipher_prefix", "").upper()
        for session in sender_sessions.values()
        if session.get("cipher_prefix")
    }
    return prefix.upper() not in existing_prefixes and prefix.upper() not in active_prefixes


def validate_setup_value(chat_id: int, key: str, value: str) -> tuple[str | None, str | None]:
    value = clean_text(value)
    if not value:
        return None, "Bu maydon bo'sh bo'lmasin. Iltimos, qayta kiriting."

    if key == "sender_phone":
        normalized, review = normalize_phone(value)
        if review:
            return None, "Telefon raqam noaniq. Masalan: 998901234567 yoki +998 90 123 45 67"
        return normalized, None

    if key == "cipher_prefix":
        prefix = normalize_cipher_prefix(value)
        if not prefix:
            return None, "Shifr faqat harf/raqamlardan iborat bo'lsin. Masalan: ABC"
        current_session = sender_sessions.get(chat_id, {})
        if current_session.get("cipher_prefix", "").upper() == prefix:
            return prefix, None
        if not is_cipher_prefix_available(prefix):
            return None, f"{prefix} shifri oldin ishlatilgan. Boshqa prefix kiriting."
        return prefix, None

    if key == "payment_by_receiver":
        parsed = parse_bool(value)
        if parsed is None:
            return None, "Faqat True yoki False deb yozing."
        return parsed, None

    if key == "places_count":
        digits = re.sub(r"\D", "", value)
        if not digits or int(digits) < 1:
            return None, "Jo'natma soni 1 yoki undan katta raqam bo'lishi kerak."
        return str(int(digits)), None

    return value, None


def setup_summary(session: dict[str, str]) -> str:
    return (
        "Jo'natuvchi ma'lumotlari saqlandi:\n"
        f"Ism familiya: {session['sender_full_name']}\n"
        f"Telefon: {session['sender_phone']}\n"
        f"Manzil: {session['sender_address']}\n"
        f"Shahar: {session['sender_city_ru']}\n"
        f"Shifr: {session['cipher_prefix']}1, {session['cipher_prefix']}2, ...\n"
        f"Оплата получателем: {session['payment_by_receiver']}\n"
        f"Og'irlik: {session['parcel_weight']}\n"
        f"Количество мест: {session['places_count']}\n\n"
        "Endi mijozlar ro'yxatini matn yoki rasm qilib yuboring."
    )


async def start_setup(message: Message, reset: bool = False) -> None:
    chat_id = message.chat.id
    if reset:
        sender_sessions.pop(chat_id, None)
    setup_states[chat_id] = {"step": 0, "data": {}}
    await message.answer(
        "Excel yaratishdan oldin jo'natuvchi ma'lumotlarini kiritamiz.\n\n"
        f"{SETUP_STEPS[0][1]}"
    )


async def handle_setup_message(message: Message) -> bool:
    chat_id = message.chat.id
    state = setup_states.get(chat_id)
    if state is None:
        return False

    step_index = state["step"]
    key, _question = SETUP_STEPS[step_index]
    parsed_value, error = validate_setup_value(chat_id, key, message.text or "")
    if error:
        await message.answer(error)
        return True

    state["data"][key] = parsed_value or ""
    step_index += 1

    if step_index >= len(SETUP_STEPS):
        sender_sessions[chat_id] = state["data"]
        setup_states.pop(chat_id, None)
        await message.answer(setup_summary(sender_sessions[chat_id]))
        return True

    state["step"] = step_index
    await message.answer(SETUP_STEPS[step_index][1])
    return True


def prepare_rows(customers: list[dict[str, Any]], sender: dict[str, str]) -> list[list[str]]:
    rows = []
    for customer in customers:
        normalized_phone, phone_review = normalize_phone(clean_text(customer.get("phone")))
        recipient_location, location_review = resolve_allowed_recipient_location(customer)
        review_parts = [
            clean_text(customer.get("needs_review")),
            phone_review,
            location_review,
        ]
        review = "; ".join(part for part in review_parts if part)

        rows.append(
            [
                "",
                clean_name(customer.get("full_name")),
                clean_name(customer.get("full_name")),
                clean_text(customer.get("address")),
                normalized_phone,
                "",
                sender["parcel_weight"],
                clean_text(customer.get("note")),
                sender["places_count"],
                "",
                sender["sender_full_name"],
                sender["sender_full_name"],
                sender["sender_address"],
                sender["sender_phone"],
                sender["sender_city_ru"],
                recipient_location,
                sender["payment_by_receiver"],
                review,
            ]
        )
    return rows


async def append_customers(customers: list[dict[str, Any]], sender: dict[str, str]) -> int:
    rows = prepare_rows(customers, sender)
    if not rows:
        return 0

    async with excel_lock:
        ensure_excel_file()
        workbook = load_workbook(EXCEL_PATH)
        sheet = workbook.active

        next_row = find_last_data_row(sheet) + 1
        next_number = next_row - 1
        next_code_index = next_cipher_index(sheet, sender["cipher_prefix"])
        for row in rows:
            copy_row_style(sheet, 2, next_row)
            row[0] = next_number
            row[5] = f"{sender['cipher_prefix']}{next_code_index}"
            review = row.pop()
            if review:
                row[7] = "; ".join(part for part in [row[7], review] if part)
            for column_index, value in enumerate(row, start=1):
                sheet.cell(next_row, column_index).value = value
            next_row += 1
            next_number += 1
            next_code_index += 1

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
    sender = sender_sessions.get(message.chat.id)
    if sender is None:
        await message.answer("Avval jo'natuvchi ma'lumotlarini kiritish kerak.")
        await start_setup(message)
        return

    if not customers:
        await message.answer(
            "Mijoz ma'lumotlari topilmadi. Iltimos, matnni aniqroq yuboring yoki rasm sifatini yaxshilang."
        )
        return

    count = await append_customers(customers, sender)
    await message.answer(
        f"{count} ta mijoz Excel faylga qo'shildi.\n"
        "Faylni olish uchun /excel buyrug'ini yuboring."
    )


async def start_handler(message: Message) -> None:
    await message.answer(
        "Assalomu alaykum!\n\n"
        "Men jo'natuvchi ma'lumotlari asosida mijoz ma'lumotlarini matn yoki rasm ichidan ajratib, Excel shablonga yozib boraman.\n\n"
        "Yuborishingiz mumkin:\n"
        "- oddiy matn\n"
        "- daftar rasmi\n"
        "- skrinshot\n"
        "- qo'lda yozilgan ma'lumot rasmi\n\n"
        "Jo'natuvchi ma'lumotlarini qayta sozlash: /setup\n"
        "Excel faylni olish: /excel\n"
        "Ro'yxatni tozalash: /clear\n"
        "Yordam: /help"
    )
    if message.chat.id not in sender_sessions:
        await start_setup(message)


async def help_handler(message: Message) -> None:
    await message.answer(
        "Foydalanish:\n\n"
        "1. Mijoz ma'lumotlarini matn qilib yuboring yoki rasm jo'nating.\n"
        "2. Agar jo'natuvchi ma'lumotlari kiritilmagan bo'lsa, bot avval ularni so'raydi.\n"
        "3. Bot oluvchi ism, telefon, manzil, rus tilidagi hudud va izohlarni ajratadi.\n"
        "4. Telefonlar 998XXXXXXXXX formatiga keltiriladi.\n"
        "5. Shifrlar prefix bo'yicha ketadi: ABC1, ABC2, ABC3...\n\n"
        "Komandalar:\n"
        "/setup - jo'natuvchi ma'lumotlarini qayta kiritish\n"
        "/excel - Excel faylni yuboradi\n"
        "/clear - ro'yxatni tozalaydi\n"
        "/help - yordam"
    )


async def setup_handler(message: Message) -> None:
    await start_setup(message, reset=True)


async def excel_handler(message: Message) -> None:
    file_bytes = await get_excel_bytes()
    await message.answer_document(
        BufferedInputFile(file_bytes, filename="customers.xlsx"),
        caption="Yangilangan mijozlar ro'yxati.",
    )


async def clear_handler(message: Message) -> None:
    async with excel_lock:
        reset_excel_file()
    sender_sessions.pop(message.chat.id, None)
    setup_states.pop(message.chat.id, None)
    await message.answer("Ro'yxat tozalandi. Yangi Excel fayl shablondan tayyorlanadi.")
    await start_setup(message)


async def text_handler(message: Message) -> None:
    text = message.text or ""
    if await handle_setup_message(message):
        return

    if message.chat.id not in sender_sessions:
        await start_setup(message)
        return

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
    if message.chat.id in setup_states:
        await message.answer("Hozir jo'natuvchi ma'lumotlarini matn ko'rinishida kiriting.")
        return

    if message.chat.id not in sender_sessions:
        await start_setup(message)
        return

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
    if message.chat.id in setup_states:
        await message.answer("Hozir jo'natuvchi ma'lumotlarini matn ko'rinishida kiriting.")
        return

    if message.chat.id not in sender_sessions:
        await start_setup(message)
        return

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


async def setup_bot_commands(bot: Bot) -> None:
    await bot.set_my_commands(
        [
            BotCommand(command="start", description="Botni ishga tushirish"),
            BotCommand(command="setup", description="Jo'natuvchi ma'lumotlarini sozlash"),
            BotCommand(command="help", description="Foydalanish bo'yicha yordam"),
            BotCommand(command="excel", description="Excel faylni yuborish"),
            BotCommand(command="clear", description="Ro'yxatni tozalash"),
        ]
    )


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
    await setup_bot_commands(bot)

    dispatcher.message.register(start_handler, Command("start"))
    dispatcher.message.register(setup_handler, Command("setup"))
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
