import asyncio
import base64
import io
import json
import logging
import os
import re
import shutil
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from difflib import get_close_matches
from copy import copy
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F
from aiogram.exceptions import TelegramRetryAfter
from aiogram.filters import Command
from aiogram.types import (
    BotCommand,
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook


load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
ADMIN_IDS = {
    int(part.strip())
    for part in os.getenv("ADMIN_IDS", "").split(",")
    if part.strip().isdigit()
}

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "templates"
TEMPLATE_PATH = TEMPLATE_DIR / "yangi_shablon.xlsx"
BRANCH_CODES_PATH = TEMPLATE_DIR / "branch_codes.xlsx"
EXCEL_PATH = DATA_DIR / "customers.xlsx"
APPROVED_USERS_PATH = DATA_DIR / "approved_users.json"

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
    "Режим",
    "Компания-отправитель",
    "ФИО отправителя",
    "Адрес отправителя",
    "Телефон отправителя",
    "Город-отправитель",
    "Город-получатель",
    "Оплата получателем",
    "Тип вложение",
    "Код упаковке",
]

EXCEL_COLUMN_COUNT = len(HEADERS)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

excel_lock = asyncio.Lock()
openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
sender_sessions: dict[int, dict[str, str]] = {}
setup_states: dict[int, dict[str, Any]] = {}
branch_code_cache: dict[str, str] | None = None
reply_locks: dict[int, asyncio.Lock] = {}
last_success_notice_at: dict[int, float] = {}
batch_states: dict[int, "BatchState"] = {}
access_request_sent_at: dict[int, float] = {}
service_states: dict[int, dict[str, Any]] = {}
emu_api_cache: dict[str, tuple[float, Any]] = {}

SUCCESS_NOTICE_INTERVAL_SECONDS = 12
BATCH_IDLE_SECONDS = 3
BATCH_CONCURRENCY = max(1, int(os.getenv("BATCH_CONCURRENCY", "10")))
BATCH_PROGRESS_EDIT_INTERVAL_SECONDS = 1.5
ACCESS_REQUEST_INTERVAL_SECONDS = 60

MENU_COLLECT = "📥 Excel ga yig'ish"
MENU_OFFICES = "🏢 Ofislar ro'yxati"
MENU_CALCULATOR = "🧮 Kalkulyator"
MENU_AI_ASSISTANT = "🤖 AI yordamchi"
MENU_ARCHIVE = "🗂 Arxiv"
MENU_SETTINGS = "⚙️ Sozlamalar"
MENU_BACK = "⬅️ Orqaga"
MENU_LEGAL = "🏢 Yuridik mijoz"
MENU_PHYSICAL = "👤 ФИЗ ЛИЦО"
MENU_EXCEL_FILE = "📊 Excel fayl"
MENU_TEMPLATE_FILE = "📄 Shablon"
MENU_CLEAR = "🧹 Ro'yxatni tozalash"
MENU_RESET_SETUP = "✏️ Jo'natuvchi sozlamalari"
MENU_ACCESS_STATUS = "🔐 Ruxsat holati"
MENU_NEXT_PAGE = "➡️ Keyingi"
MENU_PREV_PAGE = "⬅️ Oldingi"
MENU_SEARCH = "🔎 Qidirish"
MENU_CANCEL = "❌ Bekor qilish"
MENU_SKIP = "⏭️ O'tkazib yuborish"
MENU_DO_OFFICE = "🏢 ДО ОФИСА"
MENU_TO_HOME = "🏠 НА ДОМ"
MENU_TEXTS = {
    MENU_COLLECT,
    MENU_OFFICES,
    MENU_CALCULATOR,
    MENU_AI_ASSISTANT,
    MENU_ARCHIVE,
    MENU_SETTINGS,
    MENU_BACK,
    MENU_LEGAL,
    MENU_PHYSICAL,
    MENU_EXCEL_FILE,
    MENU_TEMPLATE_FILE,
    MENU_CLEAR,
    MENU_RESET_SETUP,
    MENU_ACCESS_STATUS,
    "Excel ga yig'ish",
    "Ofislar ro'yxati",
    "Kalkulyator",
    "AI yordamchi",
    "Arxiv",
    "Sozlamalar",
    "Orqaga",
    "Yuridik mijoz",
    "ФИЗ ЛИЦО",
    "Excel fayl",
    "Shablon",
    "Ro'yxatni tozalash",
    "Jo'natuvchi sozlamalari",
    "Ruxsat holati",
}

CLIENT_TYPE_LEGAL = "legal"
CLIENT_TYPE_PHYSICAL = "physical"
EMU_API_BASE_URL = "https://apiv1.emu.uz"
TASHKENT_REGION_ID = 13
TASHKENT_CITY_ID = 198
EMU_CACHE_TTL_SECONDS = 3600
OFFICES_PAGE_SIZE = 12

MENU_ALIASES = {
    "Excel ga yig'ish": MENU_COLLECT,
    "Ofislar ro'yxati": MENU_OFFICES,
    "Kalkulyator": MENU_CALCULATOR,
    "AI yordamchi": MENU_AI_ASSISTANT,
    "Arxiv": MENU_ARCHIVE,
    "Sozlamalar": MENU_SETTINGS,
    "Orqaga": MENU_BACK,
    "Yuridik mijoz": MENU_LEGAL,
    "ФИЗ ЛИЦО": MENU_PHYSICAL,
    "Excel fayl": MENU_EXCEL_FILE,
    "Shablon": MENU_TEMPLATE_FILE,
    "Ro'yxatni tozalash": MENU_CLEAR,
    "Jo'natuvchi sozlamalari": MENU_RESET_SETUP,
    "Ruxsat holati": MENU_ACCESS_STATUS,
}


@dataclass
class BatchItem:
    kind: str
    message: Message
    text: str = ""
    file_id: str = ""
    mime_type: str = ""
    bot: Bot | None = None


@dataclass
class BatchState:
    items: list[BatchItem] = field(default_factory=list)
    task: asyncio.Task | None = None
    last_added_at: float = 0


def load_approved_user_ids() -> set[int]:
    if not APPROVED_USERS_PATH.exists():
        return set()
    try:
        data = json.loads(APPROVED_USERS_PATH.read_text(encoding="utf-8"))
    except Exception as error:
        logger.warning("Approved users file could not be read: %s", error)
        return set()
    if not isinstance(data, list):
        return set()
    return {int(item) for item in data if str(item).isdigit()}


def save_approved_user_ids(user_ids: set[int]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    APPROVED_USERS_PATH.write_text(
        json.dumps(sorted(user_ids), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


approved_user_ids: set[int] = load_approved_user_ids()


def emu_cache_key(path: str, params: dict[str, Any] | None = None) -> str:
    clean_params = {
        key: value
        for key, value in (params or {}).items()
        if value not in (None, "")
    }
    return f"{path}?{urllib.parse.urlencode(sorted(clean_params.items()))}"


def emu_api_get(path: str, params: dict[str, Any] | None = None) -> Any:
    key = emu_cache_key(path, params)
    cached = emu_api_cache.get(key)
    now = time.time()
    if cached and now - cached[0] < EMU_CACHE_TTL_SECONDS:
        return cached[1]

    query = urllib.parse.urlencode(
        {
            name: value
            for name, value in (params or {}).items()
            if value not in (None, "")
        }
    )
    url = f"{EMU_API_BASE_URL}{path}{'?' + query if query else ''}"
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json",
            "Accept-Language": "uz",
            "User-Agent": "EMU Assistant Bot",
        },
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        data = json.loads(response.read().decode("utf-8"))
    emu_api_cache[key] = (now, data)
    return data


def emu_api_post(path: str, payload: dict[str, Any], params: dict[str, Any] | None = None) -> Any:
    query = urllib.parse.urlencode(
        {
            name: value
            for name, value in (params or {}).items()
            if value not in (None, "")
        }
    )
    url = f"{EMU_API_BASE_URL}{path}{'?' + query if query else ''}"
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Accept": "application/json",
            "Accept-Language": "uz",
            "Content-Type": "application/json",
            "User-Agent": "EMU Assistant Bot",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


async def get_emu_regions() -> list[dict[str, Any]]:
    return await asyncio.to_thread(emu_api_get, "/api/v1/regions")


async def get_emu_cities(region_id: int | None = None) -> list[dict[str, Any]]:
    params = {"region_id": region_id} if region_id else None
    return await asyncio.to_thread(emu_api_get, "/api/v1/cities", params)


async def get_emu_branches(region_id: int | None = None, city_id: int | None = None) -> list[dict[str, Any]]:
    params = {"region_id": region_id, "city_id": city_id}
    return await asyncio.to_thread(emu_api_get, "/api/v1/branches", params)


async def calculate_emu_delivery(
    sender_city_id: int,
    receiver_city_id: int,
    weight: float,
    service_id: int | None = None,
) -> dict[str, Any]:
    payload = {
        "sender_city_id": sender_city_id,
        "receiver_city_id": receiver_city_id,
        "service": service_id,
        "packages": [
            {
                "mass": weight,
                "length": None,
                "width": None,
                "height": None,
            }
        ],
    }
    return await asyncio.to_thread(emu_api_post, "/api/v1/calculator", payload, {"platform": "app"})


def localized_name(item: dict[str, Any], locale: str = "UZ") -> str:
    i18n = item.get("i18n_name")
    if isinstance(i18n, dict):
        return clean_text(i18n.get(locale)) or clean_text(i18n.get("UZ")) or clean_text(item.get("name"))
    return clean_text(item.get("name"))


def chunked(items: list[Any], size: int) -> list[list[Any]]:
    return [items[index : index + size] for index in range(0, len(items), size)]


def reply_keyboard(labels: list[str], row_size: int = 2, add_back: bool = True) -> ReplyKeyboardMarkup:
    rows = [
        [KeyboardButton(text=label) for label in row]
        for row in chunked(labels, row_size)
    ]
    if add_back:
        rows.append([KeyboardButton(text=MENU_BACK)])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def remember_options(state: dict[str, Any], items: list[dict[str, Any]], locale: str = "UZ") -> list[str]:
    labels: list[str] = []
    options: dict[str, int] = {}
    for item in items:
        label = localized_name(item, locale)
        if not label:
            continue
        labels.append(label)
        options[label] = int(item.get("id") or 0)
    state["options"] = options
    return labels


def selected_option_id(state: dict[str, Any], text: str) -> int | None:
    options = state.get("options") or {}
    if text in options:
        return int(options[text])
    normalized = clean_text(text).casefold()
    for label, value in options.items():
        if clean_text(label).casefold() == normalized:
            return int(value)
    return None


def region_reply_keyboard(regions: list[dict[str, Any]], state: dict[str, Any]) -> ReplyKeyboardMarkup:
    return reply_keyboard(remember_options(state, regions), row_size=2)


def city_reply_keyboard(cities: list[dict[str, Any]], state: dict[str, Any]) -> ReplyKeyboardMarkup:
    return reply_keyboard(remember_options(state, cities), row_size=2)


def calculator_service_reply_keyboard() -> ReplyKeyboardMarkup:
    return reply_keyboard([MENU_DO_OFFICE, MENU_TO_HOME], row_size=2)


def offices_page_keyboard(page: int, total: int) -> ReplyKeyboardMarkup:
    labels: list[str] = []
    if page > 0:
        labels.append(MENU_PREV_PAGE)
    if (page + 1) * OFFICES_PAGE_SIZE < total:
        labels.append(MENU_NEXT_PAGE)
    labels.append(MENU_SEARCH)
    return reply_keyboard(labels, row_size=2)


def region_keyboard(regions: list[dict[str, Any]], prefix: str) -> InlineKeyboardMarkup:
    buttons = [
        InlineKeyboardButton(text=localized_name(region), callback_data=f"{prefix}:{region['id']}")
        for region in regions
    ]
    rows = chunked(buttons, 2)
    rows.append([InlineKeyboardButton(text=MENU_BACK, callback_data="emu:back")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def city_keyboard(cities: list[dict[str, Any]], prefix: str) -> InlineKeyboardMarkup:
    buttons = [
        InlineKeyboardButton(text=localized_name(city), callback_data=f"{prefix}:{city['id']}")
        for city in cities
    ]
    rows = chunked(buttons, 2)
    rows.append([InlineKeyboardButton(text=MENU_BACK, callback_data="emu:back")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def service_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="ДО ОФИСА", callback_data="emu:calc_service:1"),
                InlineKeyboardButton(text="НА ДОМ", callback_data="emu:calc_service:3"),
            ],
            [InlineKeyboardButton(text=MENU_BACK, callback_data="emu:back")],
        ]
    )


def format_branch_card(branch: dict[str, Any], index: int) -> str:
    schedule = branch.get("work_schedule") or []
    active_days = [day for day in schedule if day.get("is_active")]
    work_time = ""
    if active_days:
        first = active_days[0]
        work_time = f"{clean_text(first.get('start_time'))[:5]}-{clean_text(first.get('end_time'))[:5]}"
    open_status = "ochiq" if branch.get("is_open_now") else "yopiq"
    phone = clean_text(branch.get("phone")) or "telefon yo'q"
    address_parts = [clean_text(branch.get("address")), clean_text(branch.get("address_ref"))]
    address = " | ".join(part for part in address_parts if part)
    return (
        f"{index}. {clean_text(branch.get('name'))}\n"
        f"   Hudud: {clean_text(branch.get('region_name'))}, {clean_text(branch.get('city_name'))}\n"
        f"   Manzil: {address or 'manzil korsatilmagan'}\n"
        f"   Tel: {phone}\n"
        f"   Ish vaqti: {work_time or 'korsatilmagan'} | Hozir: {open_status}"
    )


def format_branches_list(branches: list[dict[str, Any]], title: str, limit: int = OFFICES_PAGE_SIZE) -> str:
    if not branches:
        return f"{title}\n\nBu hudud uchun ofis topilmadi."

    visible = branches[:limit]
    lines = [title, f"Jami: {len(branches)} ta ofis", ""]
    lines.extend(format_branch_card(branch, index) for index, branch in enumerate(visible, start=1))
    if len(branches) > limit:
        lines.append(f"\nYana {len(branches) - limit} ta ofis bor. Aniq tuman bo'yicha qidirsak, ro'yxat qisqaradi.")
    return "\n\n".join(lines)


def format_branches_page(branches: list[dict[str, Any]], title: str, page: int = 0) -> str:
    if not branches:
        return f"{title}\n\n😕 Bu hudud uchun ofis topilmadi."

    total = len(branches)
    page_count = max(1, (total + OFFICES_PAGE_SIZE - 1) // OFFICES_PAGE_SIZE)
    page = max(0, min(page, page_count - 1))
    start = page * OFFICES_PAGE_SIZE
    visible = branches[start : start + OFFICES_PAGE_SIZE]

    lines = [
        f"🏢 {title}",
        f"📍 Jami: {total} ta ofis",
        f"📄 Sahifa: {page + 1}/{page_count}",
        "",
    ]
    lines.extend(
        format_branch_card(branch, index)
        for index, branch in enumerate(visible, start=start + 1)
    )
    if page + 1 < page_count:
        lines.append(f"\n➡️ Yana {total - (start + len(visible))} ta ofis bor. Pastdan {MENU_NEXT_PAGE} ni bosing.")
    return "\n\n".join(lines)


def filter_branches(branches: list[dict[str, Any]], query: str) -> list[dict[str, Any]]:
    query = clean_text(query).casefold()
    if not query:
        return branches
    return [
        branch
        for branch in branches
        if query
        in " ".join(
            [
                clean_text(branch.get("name")),
                clean_text(branch.get("address")),
                clean_text(branch.get("address_ref")),
                clean_text(branch.get("city_name")),
                clean_text(branch.get("region_name")),
                clean_text(branch.get("phone")),
            ]
        ).casefold()
    ]


def format_calculator_result(
    result: dict[str, Any],
    service_id: int,
    receiver_branches: list[dict[str, Any]],
) -> str:
    results = result.get("results") if isinstance(result, dict) else []
    selected = None
    for item in results or []:
        if int(item.get("service_id") or 0) == service_id:
            selected = item
            break
    if selected is None and results:
        selected = results[0]

    if not selected:
        price_text = "Narx topilmadi"
    else:
        price = selected.get("price")
        currency = selected.get("currency") or "UZS"
        days = ""
        if selected.get("min_delivery_days") and selected.get("max_delivery_days"):
            days = f"\nMuddat: {selected['min_delivery_days']}-{selected['max_delivery_days']} kun"
        price_value = f"{float(price):,.0f}".replace(",", " ") if price is not None else "topilmadi"
        price_text = (
            f"Xizmat: {clean_text(selected.get('service_name'))}\n"
            f"Narx: {price_value} {currency}"
            + days
        )

    branch_lines = []
    if receiver_branches:
        branch_lines.append("Mavjud ofislar:")
        for branch in receiver_branches[:5]:
            branch_lines.append(f"- {clean_text(branch.get('name'))}: {clean_text(branch.get('address'))}")
        if len(receiver_branches) > 5:
            branch_lines.append(f"... yana {len(receiver_branches) - 5} ta ofis bor")
    else:
        branch_lines.append("Bu tuman/shahar uchun ofis topilmadi.")

    return "Hisob-kitob natijasi:\n\n" + price_text + "\n\n" + "\n".join(branch_lines)


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def has_bot_access(user_id: int) -> bool:
    if not ADMIN_IDS:
        return True
    return is_admin(user_id) or user_id in approved_user_ids


def main_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=MENU_COLLECT)],
            [KeyboardButton(text=MENU_OFFICES), KeyboardButton(text=MENU_CALCULATOR)],
            [KeyboardButton(text=MENU_AI_ASSISTANT)],
            [KeyboardButton(text=MENU_ARCHIVE), KeyboardButton(text=MENU_SETTINGS)],
        ],
        resize_keyboard=True,
    )


def collect_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=MENU_LEGAL), KeyboardButton(text=MENU_PHYSICAL)],
            [KeyboardButton(text=MENU_BACK)],
        ],
        resize_keyboard=True,
    )


def archive_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=MENU_EXCEL_FILE), KeyboardButton(text=MENU_TEMPLATE_FILE)],
            [KeyboardButton(text=MENU_CLEAR)],
            [KeyboardButton(text=MENU_BACK)],
        ],
        resize_keyboard=True,
    )


def settings_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=MENU_RESET_SETUP), KeyboardButton(text=MENU_ACCESS_STATUS)],
            [KeyboardButton(text=MENU_BACK)],
        ],
        resize_keyboard=True,
    )


def collect_active_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=MENU_EXCEL_FILE), KeyboardButton(text=MENU_TEMPLATE_FILE)],
            [KeyboardButton(text=MENU_BACK)],
        ],
        resize_keyboard=True,
    )


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
        "Jo'natuvchi qaysi tuman/shahardanligini rus tilida kiriting. Masalan: Ташкент, Бухара, Шафиркан",
    ),
    (
        "cipher_prefix",
        "Jo'natmalar uchun qaytarilmaydigan shifr prefixini kiriting. Masalan: ABC",
    ),
    (
        "delivery_type",
        "Yetkazib berish turini tanlang.",
    ),
    (
        "payment_by_receiver",
        "Оплата получателем bo'ladimi?",
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
    "Пахтачи",
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

BUTTON_SETUP_OPTIONS = {
    "delivery_type": [
        (MENU_DO_OFFICE, "ДО ОФИСА"),
        (MENU_TO_HOME, "НА ДОМ"),
    ],
    "payment_by_receiver": [
        ("✅ qo'yilsin", "✅ qo'yilsin"),
        ("⬜ qo'yilmasin", "⬜ qo'yilmasin"),
    ],
    "cipher_prefix": [
        (MENU_SKIP, ""),
    ],
}


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
- phone: barcha telefon raqamlari asl matndagi ko'rinishida; bir nechta bo'lsa hammasini shu maydonga yozing
- address: manzil
- recipient_region_ru: oluvchining manzilidan P ustun uchun mos shahar/tuman nomini rus tilida tanlang. Faqat quyidagi ro'yxatdan bittasini yozing, boshqa format yozmang:
{location_list}
- note: boshqa foydali izohlar, noaniq yoki yo'qolmasligi kerak bo'lgan bo'laklar; telefon raqamlarini bu maydonga yozmang
- needs_review: noaniq o'qilgan, telefon raqami shubhali, rasm sifati past, yoki maydonlar aralash bo'lsa qisqa izoh

Qoidalar:
- Ma'lumot yo'q bo'lsa bo'sh string qaytaring.
- Telefon raqamlarni formatlamang, asl ko'rinishida qaytaring.
- Taxmin qilmang. Ishonchsiz joylarni needs_review maydoniga yozing.
- Ism yo'q bo'lsa "Mijoz", "Noma'lum", "Customer" kabi placeholder yozmang, full_name bo'sh string bo'lsin.
- Bitta xabarda bitta ism/manzil va bir nechta telefon raqami bo'lsa, buni bitta mijoz deb oling: hamma telefon raqamlarni phone maydoniga yozing, note maydoniga telefon yozmang.
- Faqat aniq boshqa-boshqa mijozlar bo'lsa alohida obyekt qiling.
- recipient_region_ru hech qachon "Ферганская область, Учкуприкский район" kabi bo'lmasin; ro'yxatdagi "Учкуприк" kabi bitta qiymat bo'lsin.
- Agar manzilda viloyat/tuman/shahar nomi bor bo'lsa, recipient_region_ru ni bo'sh qoldirmang; ro'yxatdan eng yaqin mos qiymatni tanlang.
- "Samarqand viloyati Paxtachi tumani" bo'lsa recipient_region_ru uchun "Пахтачи" yozing.
- Tuman yoki shahar nomi viloyatdan muhimroq: "Farg'ona viloyati Oltiariq tumani" uchun "Фергана" emas, "Алтыарык" yozing.
- Lotin yozuvidagi O'zbekcha nomlarni ruscha ro'yxatga moslang: Qorako'l -> Каракуль, Qo'rg'ontepa -> Кургантепа, Bo'ka -> Бука, Tayloq -> Тайлак.
- Javob faqat schema bo'yicha bo'lsin.
""".strip().format(location_list=LOCATION_LIST_FOR_PROMPT)


def apply_template_header(sheet: Any) -> None:
    if TEMPLATE_PATH.exists():
        template = load_workbook(TEMPLATE_PATH)
        try:
            template_sheet = template.active
            for col in range(1, EXCEL_COLUMN_COUNT + 1):
                source = template_sheet.cell(1, col)
                target = sheet.cell(1, col)
                target.value = source.value
                if source.has_style:
                    target._style = copy(source._style)
                target.number_format = source.number_format
                target.alignment = copy(source.alignment)
                letter = target.column_letter
                sheet.column_dimensions[letter].width = template_sheet.column_dimensions[letter].width
        finally:
            template.close()
        return

    for col, header in enumerate(HEADERS, start=1):
        sheet.cell(1, col).value = header


def ensure_workbook_schema() -> None:
    workbook = load_workbook(EXCEL_PATH)
    try:
        sheet = workbook.active
        current_headers = [sheet.cell(1, col).value for col in range(1, EXCEL_COLUMN_COUNT + 1)]

        if current_headers == HEADERS:
            return

        if (
            current_headers[:17] == HEADERS[:17]
            and current_headers[17] == "Код упаковке"
            and (len(current_headers) < 19 or current_headers[18] in (None, ""))
        ):
            sheet.insert_cols(18)
            apply_template_header(sheet)
            workbook.save(EXCEL_PATH)
            return

        if current_headers[: EXCEL_COLUMN_COUNT - 1] == HEADERS[:-1] and current_headers[-1] in (None, ""):
            apply_template_header(sheet)
            workbook.save(EXCEL_PATH)
            return

        first_row_has_data = any(value not in (None, "") for value in current_headers)
        if first_row_has_data:
            sheet.insert_rows(1)

        apply_template_header(sheet)
        workbook.save(EXCEL_PATH)
    finally:
        workbook.close()


def ensure_excel_file() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        ensure_workbook_schema()
        return

    if TEMPLATE_PATH.exists():
        shutil.copyfile(TEMPLATE_PATH, EXCEL_PATH)
        workbook = load_workbook(EXCEL_PATH)
        try:
            sheet = workbook.active
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=EXCEL_COLUMN_COUNT):
                for cell in row:
                    cell.value = None
            workbook.save(EXCEL_PATH)
        finally:
            workbook.close()
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Шаблон"
    sheet.append(HEADERS)
    widths = [18, 24, 19, 21, 23, 18, 13, 13, 20, 27, 25, 20, 22, 24, 22, 21, 23, 18, 16]
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


PHONE_CANDIDATE_RE = re.compile(r"\+?\d[\d\s\-()]{7,}\d")


def extract_phone_candidates(*values: str) -> list[str]:
    candidates: list[str] = []
    for value in values:
        for match in PHONE_CANDIDATE_RE.findall(value or ""):
            candidate = match.strip(" -()")
            if candidate:
                candidates.append(candidate)
    return candidates


def normalize_phone_list(*values: str) -> tuple[str, str]:
    normalized_numbers: list[str] = []
    reviews: list[str] = []

    for candidate in extract_phone_candidates(*values):
        normalized, review = normalize_phone(candidate)
        if review:
            reviews.append(review)
            continue
        if normalized not in normalized_numbers:
            normalized_numbers.append(normalized)

    if normalized_numbers:
        return "; ".join(normalized_numbers), "; ".join(reviews)

    raw_phone = clean_text(values[0]) if values else ""
    if not raw_phone:
        return "", ""
    return normalize_phone(raw_phone)


def strip_phone_candidates(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = PHONE_CANDIDATE_RE.sub("", text)
    text = re.sub(r"\b(qolgan|ikkinchi|2-?chi|telefon|tel|raqamlar|raqami)\b\s*:?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*[,;|/]\s*", " ", text)
    return re.sub(r"\s{2,}", " ", text).strip(" -,:;|/")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def clean_name(value: Any) -> str:
    name = clean_text(value)
    if not name or name.lower() in {"mijoz", "noma'lum", "nomalum", "unknown", "customer"}:
        return "MIJOZ"
    return name


def clean_address(value: Any, recipient_location: str) -> str:
    address = clean_text(value)
    if address:
        return address
    if recipient_location:
        return f"{recipient_location} markazi"
    return ""


def load_branch_codes() -> dict[str, str]:
    global branch_code_cache
    if branch_code_cache is not None:
        return branch_code_cache

    branch_code_cache = {}
    if not BRANCH_CODES_PATH.exists():
        logger.warning("Branch code file not found: %s", BRANCH_CODES_PATH)
        return branch_code_cache

    workbook = load_workbook(BRANCH_CODES_PATH, data_only=True)
    try:
        sheet = workbook.active
        for row_index in range(2, sheet.max_row + 1):
            code = clean_text(sheet.cell(row_index, 1).value)
            city = clean_text(sheet.cell(row_index, 2).value)
            key = normalize_location_key(city)
            if code and key and key not in branch_code_cache:
                branch_code_cache[key] = code
    finally:
        workbook.close()

    return branch_code_cache


def branch_code_for_location(recipient_location: str) -> str:
    if not recipient_location:
        return ""
    return load_branch_codes().get(normalize_location_key(recipient_location), "")


def region_center_for_location(recipient_location: str) -> str:
    return REGION_CENTER_BY_LOCATION.get(recipient_location, recipient_location)


def format_recipient_address(value: Any, recipient_location: str, delivery_type: str) -> tuple[str, str]:
    if delivery_type == "ДО ОФИСА":
        region_center = region_center_for_location(recipient_location)
        branch_code = branch_code_for_location(region_center)
        if branch_code:
            return branch_code, ""
        if recipient_location:
            return "", f"{region_center} uchun filial kodi topilmadi"
        return "", ""

    address = clean_text(value)
    if address:
        return address, ""
    if recipient_location:
        return f"{recipient_location} markazi", ""
    return "", ""


CYRILLIC_TO_LATIN = str.maketrans(
    {
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "е": "e",
        "ё": "yo",
        "ж": "j",
        "з": "z",
        "и": "i",
        "й": "y",
        "к": "k",
        "л": "l",
        "м": "m",
        "н": "n",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "у": "u",
        "ф": "f",
        "х": "x",
        "ц": "s",
        "ч": "ch",
        "ш": "sh",
        "щ": "sh",
        "ъ": "",
        "ы": "i",
        "ь": "",
        "э": "e",
        "ю": "yu",
        "я": "ya",
        "қ": "q",
        "ғ": "g",
        "ҳ": "h",
        "ў": "o",
    }
)


def transliterate_location_text(value: str) -> str:
    return value.casefold().translate(CYRILLIC_TO_LATIN)


def normalize_location_key(value: str) -> str:
    value = transliterate_location_text(value)
    replacements = {
        "ģ": "g",
        "ğ": "g",
        "ʼ": "",
        "ʻ": "",
        "‘": "",
        "’": "",
        "'": "",
        "`": "",
        " - ": "-",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    return re.sub(r"[^a-z0-9]+", "", value)


LOCATION_BY_KEY = {
    normalize_location_key(location): location for location in ALLOWED_RECIPIENT_LOCATIONS
}

LOCATION_ALIASES = {
    "bogot": "Багат",
    "bogat": "Багат",
    "boot": "Багат",
    "bogot tumani": "Багат",
    "bagat": "Багат",
    "toshkent": "Ташкент",
    "tashkent": "Ташкент",
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
    "paxtachi": "Пахтачи",
    "paxtachitumani": "Пахтачи",
    "paxtachi tumani": "Пахтачи",
    "pakhtachi": "Пахтачи",
    "pakhtachitumani": "Пахтачи",
    "pakhtachi tumani": "Пахтачи",
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
    "oltariqtumani": "Алтыарык",
    "oltiriqtumani": "Алтыарык",
    "oltiariqtumani": "Алтыарык",
    "oltiariq tumani": "Алтыарык",
    "altyaryk": "Алтыарык",
    "andijon": "Андижан",
    "andijan": "Андижан",
    "jizzax": "Джизак",
    "jizzakh": "Джизак",
    "djizak": "Джизак",
    "qashqadaryo": "Карши",
    "kashkadarya": "Карши",
    "qarshi": "Карши",
    "karshi": "Карши",
    "navoiy": "Навои",
    "navoi": "Навои",
    "namangan": "Наманган",
    "qoraqalpogiston": "Нукус",
    "qoraqalpog'iston": "Нукус",
    "karakalpakstan": "Нукус",
    "nukus": "Нукус",
    "sirdaryo": "Гулистан",
    "syrdarya": "Гулистан",
    "surxondaryo": "Термез",
    "surkhandarya": "Термез",
    "termiz": "Термез",
    "termez": "Термез",
    "xorazm": "Ургенч",
    "khorezm": "Ургенч",
    "urganch": "Ургенч",
    "urgench": "Ургенч",
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

REGION_CENTER_BY_LOCATION = {
    **dict.fromkeys(
        [
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
        ],
        "Ташкент",
    ),
    **dict.fromkeys(
        [
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
        ],
        "Андижан",
    ),
    **dict.fromkeys(
        ["Бухара", "Алат", "Вабкент", "Галлаасия", "Гиждуван", "Джандар", "Каган", "Каракуль", "Караулбазар", "Пешку", "Ромитан", "Шафиркан"],
        "Бухара",
    ),
    **dict.fromkeys(
        ["Джизак", "Арнасай", "Балангачкыр", "Бахмал", "Гагарин", "Галляарал", "Дустлик", "Заамин", "Зарбдар", "Зафарабад", "Пахтакор", "Шараф-Рашидов", "Янгикишлак"],
        "Джизак",
    ),
    **dict.fromkeys(
        ["Карши", "Бешкент", "Гузар", "Дехканабадский район", "Камаши", "Касан", "Касбий", "Китоб", "Кокдала", "Миришкор", "Мубарек", "Нишан", "Чиракчи", "Шахрисабз", "Яккабаг"],
        "Карши",
    ),
    **dict.fromkeys(
        ["Навои", "Бешрабат", "Зарафшан", "Канимех", "Кармана", "Кызылтепа", "Нурата", "Тамдыбулак", "Учкудук", "Хатырчи"],
        "Навои",
    ),
    **dict.fromkeys(
        ["Наманган", "Джумашуй", "Касансай", "Пап", "Ташбулак", "Туракурган", "Уйчи", "Учкурган", "Хаккулабад", "Чартак", "Чуст", "Янгикурган"],
        "Наманган",
    ),
    **dict.fromkeys(
        ["Нукус", "Акмангит", "Амударья", "Беруни", "Казакеткен", "Канлыкуль", "Караузяк", "Кегейли", "Кунград", "Муйнак", "Тахиаташ", "Тахтакупыр", "Турткуль", "Ходжейли", "Чимбай", "Шуманай", "Элликкала"],
        "Нукус",
    ),
    **dict.fromkeys(
        ["Самарканд", "Акташ", "Булунгур", "Гульабад", "Джамбай", "Джума", "Зиадин", "Иштыхан", "Каттакурган", "Кушрабад", "Лаиш", "Нурабад", "Пайарык", "Пахтачи", "Тайлак", "Ургут"],
        "Самарканд",
    ),
    **dict.fromkeys(
        ["Термез", "Ангор", "Байсун", "Бандихан", "Денау", "Джаркурган", "Карлук", "Кизирик", "Кумкурган", "Сариасия", "Узун", "Учкизил", "Халкабад", "Шерабад", "Шурчи"],
        "Термез",
    ),
    **dict.fromkeys(
        ["Гулистан", "Акалтын", "Бахт", "Дехканабад", "Навруз", "Сайхун", "Сардоба", "Сырдарья", "Хаваст", "Ширин", "Янгиер"],
        "Гулистан",
    ),
    **dict.fromkeys(
        ["Нурафшон", "Аккурган", "Алмалык", "Ангрен", "Ахангаран", "Бекабад", "Бука", "Верхне-Чирчикский", "Газалкент", "Дустабад", "Зангиата", "Келес", "Кибрай", "Паркент", "Пскент", "Чиназ", "Чирчик", "Янгийоль"],
        "Нурафшон",
    ),
    **dict.fromkeys(
        ["Фергана", "Алтыарык", "Багдад", "Бешарык", "Бувайда", "Водил", "Дангара", "Коканд", "Кува", "Кувасай", "Куштепа", "Маргилан", "Навбахор", "Риштан", "Сох", "Ташлак", "Учкуприк", "Язъяван", "Яйпан"],
        "Фергана",
    ),
    **dict.fromkeys(
        ["Ургенч", "Багат", "Гурлен", "Караул", "Кошкупыр", "Тупраккала", "Хазарасп", "Ханка", "Хива", "Шават", "Янгиарык", "Янгибазар"],
        "Ургенч",
    ),
}

LOCATION_STOP_WORDS = {
    "viloyati",
    "viloyat",
    "tumani",
    "tuman",
    "shahar",
    "shahri",
    "shaharchasi",
    "kocha",
    "kochasi",
    "oblast",
    "rayon",
    "gorod",
    "mahalla",
    "mahallasi",
    "mfy",
    "uy",
    "dom",
    "kv",
    "n",
}


def location_tokens(value: str) -> list[str]:
    normalized = transliterate_location_text(value)
    replacements = {
        "ģ": "g",
        "ğ": "g",
        "ʼ": "",
        "ʻ": "",
        "‘": "",
        "’": "",
        "'": "",
        "`": "",
    }
    for source, target in replacements.items():
        normalized = normalized.replace(source, target)

    tokens = re.findall(r"[a-z0-9]+", normalized)
    return [token for token in tokens if token not in LOCATION_STOP_WORDS and len(token) > 1]


def token_window_keys(value: str) -> list[str]:
    tokens = location_tokens(value)
    keys: list[str] = []
    for size in range(min(3, len(tokens)), 0, -1):
        for start in range(0, len(tokens) - size + 1):
            keys.append(normalize_location_key(" ".join(tokens[start : start + size])))
    return keys


def fuzzy_location_from_key(key: str) -> str:
    if not key:
        return ""

    searchable = {**LOCATION_BY_KEY, **LOCATION_ALIAS_BY_KEY}
    if key in searchable:
        return searchable[key]

    matches = get_close_matches(key, list(searchable.keys()), n=1, cutoff=0.76)
    if matches:
        return searchable[matches[0]]

    return ""


def location_specificity(location: str) -> int:
    if not location:
        return 0
    center = region_center_for_location(location)
    return 1 if center == location else 2


def fuzzy_location_from_text(value: str) -> str:
    keys = token_window_keys(value)
    candidates: list[tuple[int, int, str]] = []
    searchable = {**LOCATION_BY_KEY, **LOCATION_ALIAS_BY_KEY}

    for index, key in enumerate(keys):
        location = searchable.get(key)
        if location:
            candidates.append((location_specificity(location), -index, location))

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][2]

    for key in keys:
        location = fuzzy_location_from_key(key)
        if location:
            return location
    return ""


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

    for value in [candidate for candidate in address_values if candidate]:
        location = fuzzy_location_from_text(value)
        if location:
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

    for value in [candidate for candidate in fallback_values if candidate]:
        location = fuzzy_location_from_text(value)
        if location:
            return location, ""

    location = fuzzy_location_from_key(address_key or combined_key)
    if location:
        return location, ""

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


def chat_reply_lock(chat_id: int) -> asyncio.Lock:
    lock = reply_locks.get(chat_id)
    if lock is None:
        lock = asyncio.Lock()
        reply_locks[chat_id] = lock
    return lock


async def safe_answer(message: Message, text: str, **kwargs: Any) -> Any:
    async with chat_reply_lock(message.chat.id):
        while True:
            try:
                return await message.answer(text, **kwargs)
            except TelegramRetryAfter as error:
                await asyncio.sleep(error.retry_after + 1)


async def safe_send_message(bot: Bot, chat_id: int, text: str, **kwargs: Any) -> Any:
    async with chat_reply_lock(chat_id):
        while True:
            try:
                return await bot.send_message(chat_id, text, **kwargs)
            except TelegramRetryAfter as error:
                await asyncio.sleep(error.retry_after + 1)


async def safe_answer_document(message: Message, document: BufferedInputFile, **kwargs: Any) -> Any:
    async with chat_reply_lock(message.chat.id):
        while True:
            try:
                return await message.answer_document(document, **kwargs)
            except TelegramRetryAfter as error:
                await asyncio.sleep(error.retry_after + 1)


async def safe_edit_text(message: Message, text: str, **kwargs: Any) -> None:
    async with chat_reply_lock(message.chat.id):
        while True:
            try:
                await message.edit_text(text, **kwargs)
                return
            except TelegramRetryAfter as error:
                await asyncio.sleep(error.retry_after + 1)
            except Exception as error:
                logger.warning("Progress message edit failed: %s", error)
                return


def remember_setup_message(chat_id: int, message: Message | None) -> None:
    if message is None:
        return

    state = setup_states.get(chat_id)
    if state is None:
        return

    messages = state.setdefault("cleanup_messages", [])
    if any(saved.message_id == message.message_id for saved in messages):
        return
    messages.append(message)


async def cleanup_setup_messages(chat_id: int) -> None:
    state = setup_states.get(chat_id)
    if state is None:
        return

    messages = state.get("cleanup_messages", [])
    state["cleanup_messages"] = []
    for message in messages:
        while True:
            try:
                await message.delete()
                break
            except TelegramRetryAfter as error:
                await asyncio.sleep(error.retry_after + 1)
            except Exception as error:
                logger.debug("Setup message cleanup failed: %s", error)
                break


async def maybe_send_success_notice(message: Message, count: int) -> None:
    now = asyncio.get_running_loop().time()
    last_sent = last_success_notice_at.get(message.chat.id, 0)
    if now - last_sent < SUCCESS_NOTICE_INTERVAL_SECONDS:
        return

    last_success_notice_at[message.chat.id] = now
    await safe_answer(
        message,
        f"{count} ta mijoz Excel faylga qo'shildi.\n"
        "Ko'p xabar yuborilganda bot javoblarni kamaytiradi. Faylni olish uchun /excel yuboring.",
    )


def user_label(message: Message) -> str:
    user = message.from_user
    if user is None:
        return f"chat_id={message.chat.id}"
    name = " ".join(part for part in [user.first_name, user.last_name] if part)
    username = f"@{user.username}" if user.username else "username yo'q"
    return f"{name or 'Nomalum'} ({username}, id={user.id})"


async def request_access(message: Message) -> None:
    if not ADMIN_IDS:
        return

    user_id = message.from_user.id if message.from_user else message.chat.id
    now = asyncio.get_running_loop().time()
    last_sent = access_request_sent_at.get(user_id, 0)

    if now - last_sent >= ACCESS_REQUEST_INTERVAL_SECONDS:
        access_request_sent_at[user_id] = now
        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="Ruxsat berish", callback_data=f"access:approve:{user_id}"),
                    InlineKeyboardButton(text="Rad etish", callback_data=f"access:deny:{user_id}"),
                ]
            ]
        )
        for admin_id in ADMIN_IDS:
            await safe_send_message(
                message.bot,
                admin_id,
                "Botdan foydalanish uchun yangi so'rov:\n"
                f"{user_label(message)}",
                reply_markup=keyboard,
            )

    await safe_answer(
        message,
        "Sizga hali botdan foydalanish uchun ruxsat berilmagan.\n"
        "Admin tasdiqlagandan keyin /start ni qayta bosing.",
    )


async def ensure_user_access(message: Message) -> bool:
    user_id = message.from_user.id if message.from_user else message.chat.id
    if has_bot_access(user_id):
        return True
    await request_access(message)
    return False


async def send_main_menu(message: Message, text: str | None = None) -> None:
    await safe_answer(
        message,
        text
        or (
            "Asosiy menyu.\n\n"
            "Excel ga yig'ish - yangi jo'natmalar ro'yxatini yig'ish.\n"
            "Ofislar ro'yxati - viloyat bo'yicha filiallarni ko'rish.\n"
            "Kalkulyator - EMU API orqali narx hisoblash.\n"
            "AI yordamchi - savol berib kerakli bo'limdan ma'lumot olish.\n"
            "Arxiv - tayyor Excel va shablon fayllar.\n"
            "Sozlamalar - jo'natuvchi ma'lumotlari va ruxsat holati."
        ),
        reply_markup=main_menu_keyboard(),
    )


async def access_callback_handler(callback: CallbackQuery) -> None:
    data = callback.data or ""
    parts = data.split(":", 2)
    admin_id = callback.from_user.id

    if len(parts) != 3 or parts[0] != "access":
        await callback.answer()
        return
    if not is_admin(admin_id):
        await callback.answer("Bu amal faqat admin uchun.", show_alert=True)
        return
    if not parts[2].isdigit():
        await callback.answer("User ID noto'g'ri.", show_alert=True)
        return

    action = parts[1]
    user_id = int(parts[2])
    if action == "approve":
        approved_user_ids.add(user_id)
        save_approved_user_ids(approved_user_ids)
        await callback.answer("Ruxsat berildi.")
        if callback.message:
            await callback.message.edit_text(f"Ruxsat berildi: {user_id}")
        await safe_send_message(
            callback.bot,
            user_id,
            "Ruxsat berildi.\n\n"
            "Endi botdan foydalanishingiz mumkin. Boshlash uchun /start ni bosing.",
        )
        return

    if action == "deny":
        approved_user_ids.discard(user_id)
        save_approved_user_ids(approved_user_ids)
        await callback.answer("So'rov rad etildi.")
        if callback.message:
            await callback.message.edit_text(f"So'rov rad etildi: {user_id}")
        await safe_send_message(callback.bot, user_id, "Botdan foydalanish so'rovingiz rad etildi.")
        return

    await callback.answer("Noma'lum amal.", show_alert=True)


def normalize_cipher_prefix(value: str) -> str:
    prefix = re.sub(r"\s+", "", value.strip()).upper()
    return re.sub(r"[^A-ZА-ЯЁ0-9_-]", "", prefix)


def find_last_data_row(sheet: Any) -> int:
    for row_index in range(sheet.max_row, 1, -1):
        if any(sheet.cell(row_index, col).value not in (None, "") for col in range(1, EXCEL_COLUMN_COUNT + 1)):
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
    if source_row > sheet.max_row:
        source_row = 1
    for col in range(1, EXCEL_COLUMN_COUNT + 1):
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
    try:
        sheet = workbook.active
        existing_prefixes = used_cipher_prefixes(sheet)
    finally:
        workbook.close()
    active_prefixes = {
        session.get("cipher_prefix", "").upper()
        for session in sender_sessions.values()
        if session.get("cipher_prefix")
    }
    return prefix.upper() not in existing_prefixes and prefix.upper() not in active_prefixes


def validate_setup_value(chat_id: int, key: str, value: str) -> tuple[str | None, str | None]:
    value = clean_text(value)
    if not value and key in {"cipher_prefix"}:
        return "", None
    if not value:
        return None, "Bu maydon bo'sh bo'lmasin. Iltimos, qayta kiriting."

    if key in BUTTON_SETUP_OPTIONS:
        option_by_label = {
            option_label: option_value
            for option_label, option_value in BUTTON_SETUP_OPTIONS[key]
        }
        option_values = {option_value for _label, option_value in BUTTON_SETUP_OPTIONS[key]}
        if value in option_by_label:
            return option_by_label[value], None
        if value in option_values:
            return value, None
        if key == "cipher_prefix" and value:
            pass
        else:
            return None, "Iltimos, pastdagi tugmalardan birini tanlang."

    if key == "sender_phone":
        normalized, review = normalize_phone(value)
        if review:
            return None, "Telefon raqam noaniq. Masalan: 998901234567 yoki +998 90 123 45 67"
        return normalized, None

    if key == "cipher_prefix":
        if not value:
            return "", None
        prefix = normalize_cipher_prefix(value)
        if not prefix:
            return None, "Shifr faqat harf/raqamlardan iborat bo'lsin. Masalan: ABC"
        current_session = sender_sessions.get(chat_id, {})
        if current_session.get("cipher_prefix", "").upper() == prefix:
            return prefix, None
        if not is_cipher_prefix_available(prefix):
            return None, f"{prefix} shifri oldin ishlatilgan. Boshqa prefix kiriting."
        return prefix, None

    if key == "places_count":
        digits = re.sub(r"\D", "", value)
        if not digits or int(digits) < 1:
            return None, "Jo'natma soni 1 yoki undan katta raqam bo'lishi kerak."
        return str(int(digits)), None

    return value, None


def setup_step_keyboard(key: str) -> ReplyKeyboardMarkup | None:
    options = BUTTON_SETUP_OPTIONS.get(key)
    if not options:
        return reply_keyboard([], add_back=True)

    return reply_keyboard([label for label, _value in options], row_size=2)


async def ask_setup_step(message_or_query: Message | CallbackQuery, step_index: int) -> None:
    key, question = SETUP_STEPS[step_index]
    keyboard = setup_step_keyboard(key)
    sent_message = None
    chat_id = None

    if key == "cipher_prefix":
        question = f"{question}\n\nShifr kerak bo'lmasa, pastdagi tugmani bosing."

    if isinstance(message_or_query, CallbackQuery):
        if message_or_query.message:
            chat_id = message_or_query.message.chat.id
            sent_message = await safe_answer(message_or_query.message, question, reply_markup=keyboard)
    else:
        chat_id = message_or_query.chat.id
        sent_message = await safe_answer(message_or_query, question, reply_markup=keyboard)

    if chat_id is not None:
        remember_setup_message(chat_id, sent_message)


def setup_summary(session: dict[str, str]) -> str:
    client_type = "Yuridik mijoz" if session.get("client_type") == CLIENT_TYPE_LEGAL else "ФИЗ ЛИЦО"
    return (
        "Jo'natuvchi ma'lumotlari saqlandi:\n"
        f"Yo'nalish: {client_type}\n"
        f"Ism familiya: {session['sender_full_name']}\n"
        f"Telefon: {session['sender_phone']}\n"
        f"Manzil: {session['sender_address']}\n"
        f"Shahar: {session['sender_city_ru']}\n"
        f"Shifr: {session['cipher_prefix'] + '1, ' + session['cipher_prefix'] + '2, ...' if session['cipher_prefix'] else 'yoq'}\n"
        f"Yetkazib berish turi: {session['delivery_type']}\n"
        f"Оплата получателем: {session['payment_by_receiver']}\n"
        f"Og'irlik: {session['parcel_weight']}\n"
        f"Количество мест: {session['places_count']}\n\n"
        "Endi mijozlar ro'yxatini matn yoki rasm qilib yuboring."
    )


def legal_sender_defaults() -> dict[str, str]:
    return {
        "client_type": CLIENT_TYPE_LEGAL,
        "sender_full_name": "",
        "sender_phone": "",
        "sender_address": "",
        "sender_city_ru": "",
    }


async def start_setup(
    message: Message,
    reset: bool = False,
    client_type: str = CLIENT_TYPE_PHYSICAL,
    start_step: int = 0,
    initial_data: dict[str, str] | None = None,
) -> None:
    chat_id = message.chat.id
    if reset:
        sender_sessions.pop(chat_id, None)
    setup_states[chat_id] = {
        "step": start_step,
        "data": initial_data or {"client_type": client_type},
        "cleanup_messages": [],
    }
    intro_text = (
        "Yuridik mijoz tanlandi.\n"
        "Jo'natuvchi ma'lumotlari so'ralmaydi. Endi jo'natma sozlamalarini kiritamiz."
        if client_type == CLIENT_TYPE_LEGAL
        else "ФИЗ ЛИЦО tanlandi.\nExcel yaratishdan oldin jo'natuvchi ma'lumotlarini kiritamiz."
    )
    intro_message = await safe_answer(
        message,
        intro_text,
    )
    remember_setup_message(chat_id, intro_message)
    await ask_setup_step(message, start_step)


async def start_legal_setup(message: Message, reset: bool = True) -> None:
    await start_setup(
        message,
        reset=reset,
        client_type=CLIENT_TYPE_LEGAL,
        start_step=4,
        initial_data=legal_sender_defaults(),
    )


async def save_setup_value_and_advance(
    message_or_query: Message | CallbackQuery,
    chat_id: int,
    key: str,
    value: str,
) -> None:
    state = setup_states[chat_id]
    state["data"][key] = value
    step_index = state["step"] + 1

    if step_index >= len(SETUP_STEPS):
        sender_sessions[chat_id] = state["data"]
        setup_states.pop(chat_id, None)
        summary = setup_summary(sender_sessions[chat_id])
        if isinstance(message_or_query, CallbackQuery):
            if message_or_query.message:
                await safe_answer(message_or_query.message, summary, reply_markup=collect_active_keyboard())
        else:
            await safe_answer(message_or_query, summary, reply_markup=collect_active_keyboard())
        return

    state["step"] = step_index
    await ask_setup_step(message_or_query, step_index)


async def handle_setup_message(message: Message) -> bool:
    chat_id = message.chat.id
    state = setup_states.get(chat_id)
    if state is None:
        return False

    step_index = state["step"]
    key, _question = SETUP_STEPS[step_index]
    parsed_value, error = validate_setup_value(chat_id, key, message.text or "")
    if error:
        await safe_answer(message, error)
        if key in BUTTON_SETUP_OPTIONS:
            await ask_setup_step(message, step_index)
        return True

    remember_setup_message(chat_id, message)
    await cleanup_setup_messages(chat_id)
    await save_setup_value_and_advance(message, chat_id, key, parsed_value or "")
    return True


async def setup_callback_handler(callback: CallbackQuery) -> None:
    data = callback.data or ""
    parts = data.split(":", 2)
    if len(parts) != 3 or parts[0] != "setup":
        await callback.answer()
        return

    chat_id = callback.message.chat.id if callback.message else callback.from_user.id
    state = setup_states.get(chat_id)
    if state is None:
        await callback.answer("Sozlash jarayoni topilmadi. /setup ni bosing.", show_alert=True)
        return

    expected_key, _question = SETUP_STEPS[state["step"]]
    callback_key, value = parts[1], parts[2]
    if callback_key != expected_key:
        await callback.answer("Bu tugma eski savol uchun. Hozirgi savolga javob bering.", show_alert=True)
        return

    parsed_value, error = validate_setup_value(chat_id, expected_key, value)
    if error:
        await callback.answer(error, show_alert=True)
        return

    if callback.message:
        remember_setup_message(chat_id, callback.message)
    await callback.answer(f"Tanlandi: {parsed_value}")
    await cleanup_setup_messages(chat_id)
    await save_setup_value_and_advance(callback, chat_id, expected_key, parsed_value or "")


def prepare_rows(customers: list[dict[str, Any]], sender: dict[str, str]) -> list[list[str]]:
    rows = []
    for customer in customers:
        note = strip_phone_candidates(customer.get("note"))
        normalized_phone, phone_review = normalize_phone_list(
            clean_text(customer.get("phone")),
            clean_text(customer.get("note")),
        )
        recipient_location, location_review = resolve_allowed_recipient_location(customer)
        recipient_address, branch_code_review = format_recipient_address(
            customer.get("address"),
            recipient_location,
            sender["delivery_type"],
        )
        review_parts = [
            clean_text(customer.get("needs_review")),
            phone_review,
            location_review,
            branch_code_review,
        ]
        review = "; ".join(part for part in review_parts if part)

        rows.append(
            [
                "",
                clean_name(customer.get("full_name")),
                clean_name(customer.get("full_name")),
                recipient_address,
                normalized_phone,
                "",
                sender["parcel_weight"],
                note,
                sender["places_count"],
                sender["delivery_type"],
                sender["sender_full_name"],
                sender["sender_full_name"],
                sender["sender_address"],
                sender["sender_phone"],
                sender["sender_city_ru"],
                recipient_location,
                sender["payment_by_receiver"],
                "",
                "",
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
        try:
            sheet = workbook.active

            next_row = find_last_data_row(sheet) + 1
            next_number = next_row - 1
            next_code_index = next_cipher_index(sheet, sender["cipher_prefix"]) if sender["cipher_prefix"] else 1
            for row in rows:
                copy_row_style(sheet, 2, next_row)
                row[0] = next_number
                row[5] = f"{sender['cipher_prefix']}{next_code_index}" if sender["cipher_prefix"] else ""
                review = row.pop()
                if review:
                    row[7] = "; ".join(part for part in [row[7], review] if part)
                for column_index, value in enumerate(row, start=1):
                    sheet.cell(next_row, column_index).value = value
                next_row += 1
                next_number += 1
                if sender["cipher_prefix"]:
                    next_code_index += 1

            workbook.save(EXCEL_PATH)
        finally:
            workbook.close()

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
        await safe_answer(message, "Avval jo'natuvchi ma'lumotlarini kiritish kerak.")
        await start_setup(message)
        return

    if not customers:
        await safe_answer(
            message,
            "Mijoz ma'lumotlari topilmadi. Iltimos, matnni aniqroq yuboring yoki rasm sifatini yaxshilang."
        )
        return

    count = await append_customers(customers, sender)
    await maybe_send_success_notice(message, count)


async def enqueue_batch_item(item: BatchItem) -> None:
    chat_id = item.message.chat.id
    state = batch_states.get(chat_id)
    if state is None:
        state = BatchState()
        batch_states[chat_id] = state

    state.items.append(item)
    state.last_added_at = asyncio.get_running_loop().time()

    if state.task is None or state.task.done():
        state.task = asyncio.create_task(process_batch(chat_id))


async def wait_for_batch_idle(chat_id: int) -> BatchState | None:
    while True:
        state = batch_states.get(chat_id)
        if state is None:
            return None

        elapsed = asyncio.get_running_loop().time() - state.last_added_at
        if elapsed >= BATCH_IDLE_SECONDS:
            return state

        await asyncio.sleep(max(0.2, BATCH_IDLE_SECONDS - elapsed))


async def extract_customers_from_batch_item(item: BatchItem) -> list[dict[str, Any]]:
    if item.kind == "text":
        return await asyncio.to_thread(call_openai_with_text, item.text)

    if item.bot is None:
        raise RuntimeError("Bot instance topilmadi.")

    file = await item.bot.get_file(item.file_id)
    buffer = io.BytesIO()
    await item.bot.download_file(file.file_path, destination=buffer)
    image_bytes = buffer.getvalue()
    return await asyncio.to_thread(
        call_openai_with_image,
        image_bytes,
        item.mime_type or "image/jpeg",
    )


def batch_progress_text(total: int, processed: int, added: int, errors: list[str]) -> str:
    return (
        f"{total} ta ma'lumot qabul qilindi\n"
        f"{processed}/{total} tahlil qilindi\n"
        f"Excelga qo'shilgan mijozlar: {added}\n"
        f"Xatoliklar: {len(errors)}"
    )


def batch_final_text(total: int, added: int, errors: list[str]) -> str:
    lines = [
        "Tahlil tugadi.",
        f"Qabul qilingan xabarlar: {total}",
        f"Excelga qo'shilgan mijozlar: {added}",
    ]
    if errors:
        lines.append(f"Xatoliklar: {len(errors)}")
        lines.extend(errors[:10])
        if len(errors) > 10:
            lines.append(f"... yana {len(errors) - 10} ta xatolik bor")
    else:
        lines.append("Xatoliklar: yo'q")
    lines.append("Excel faylni olish uchun /excel yuboring.")
    return "\n".join(lines)


async def process_batch(chat_id: int) -> None:
    state = await wait_for_batch_idle(chat_id)
    if state is None or not state.items:
        batch_states.pop(chat_id, None)
        return

    items = list(state.items)
    state.items.clear()

    first_message = items[0].message
    progress_message = await safe_answer(
        first_message,
        batch_progress_text(len(items), 0, 0, []),
    )

    sender = sender_sessions.get(chat_id)
    errors: list[str] = []
    processed_total = 0
    added_total = 0
    extracted_by_index: list[list[dict[str, Any]]] = [[] for _ in items]
    last_progress_edit_at = 0.0

    if sender is None:
        await safe_edit_text(
            progress_message,
            batch_final_text(
                len(items),
                0,
                ["Jo'natuvchi ma'lumotlari sozlanmagan. /setup ni bosing."],
            ),
        )
        batch_states.pop(chat_id, None)
        return

    semaphore = asyncio.Semaphore(BATCH_CONCURRENCY)

    async def extract_with_index(index: int, item: BatchItem) -> tuple[int, list[dict[str, Any]] | None, str | None]:
        async with semaphore:
            try:
                customers = await extract_customers_from_batch_item(item)
                if not customers:
                    raise RuntimeError("mijoz ma'lumotlari topilmadi")
                return index, customers, None
            except Exception as error:
                logger.exception("Batch item failed")
                return index, None, str(error)

    tasks = [
        asyncio.create_task(extract_with_index(index, item))
        for index, item in enumerate(items, start=1)
    ]

    for task in asyncio.as_completed(tasks):
        index, customers, error = await task
        processed_total += 1
        if customers is not None:
            extracted_by_index[index - 1] = customers
        if error:
            errors.append(f"{index}-xabar: {error}")

        now = asyncio.get_running_loop().time()
        if processed_total == len(items) or now - last_progress_edit_at >= BATCH_PROGRESS_EDIT_INTERVAL_SECONDS:
            last_progress_edit_at = now
            await safe_edit_text(
                progress_message,
                batch_progress_text(len(items), processed_total, added_total, errors),
            )

    customers_to_append = [
        customer
        for customers in extracted_by_index
        for customer in customers
    ]
    if customers_to_append:
        try:
            added_total = await append_customers(customers_to_append, sender)
        except Exception as error:
            logger.exception("Batch append failed")
            errors.append(f"Excelga yozishda xatolik: {error}")

    await safe_edit_text(
        progress_message,
        batch_final_text(len(items), added_total, errors),
    )

    state = batch_states.get(chat_id)
    if state and state.items:
        state.task = asyncio.create_task(process_batch(chat_id))
    else:
        batch_states.pop(chat_id, None)


async def start_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await send_main_menu(
        message,
        "Assalomu alaykum!\n\n"
        "Bot mijoz ma'lumotlarini matn yoki rasm ichidan ajratib, Excel shablonga yozadi.\n"
        "Quyidagi bo'limlardan birini tanlang.",
    )


async def help_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await safe_answer(
        message,
        "Foydalanish yo'riqnomasi:\n\n"
        "1. Excel ga yig'ish bo'limiga kiring.\n"
        "2. Yuridik mijoz yoki ФИЗ ЛИЦО yo'nalishini tanlang.\n"
        "3. Bot so'ragan sozlamalarga javob bering.\n"
        "4. Mijoz ma'lumotlarini matn yoki rasm qilib yuboring.\n"
        "5. Tayyor faylni Arxiv bo'limidan oling.\n\n"
        "Har bir ichki bo'limda Orqaga tugmasi bor.",
        reply_markup=main_menu_keyboard(),
    )


async def setup_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await start_setup(message, reset=True)


async def offices_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await show_offices_menu(message)


async def calculator_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await show_calculator_menu(message)


async def ai_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await show_ai_assistant(message)


async def excel_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    file_bytes = await get_excel_bytes()
    await safe_answer_document(
        message,
        BufferedInputFile(file_bytes, filename="customers.xlsx"),
        caption="Yangilangan mijozlar ro'yxati.",
    )


async def template_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    if not TEMPLATE_PATH.exists():
        await safe_answer(message, "Shablon fayl topilmadi.")
        return

    await safe_answer_document(
        message,
        BufferedInputFile(TEMPLATE_PATH.read_bytes(), filename="yangi_shablon.xlsx"),
        caption="Excel shablon fayli.",
    )


async def show_offices_menu(message: Message) -> None:
    state = {"mode": "offices", "step": "region"}
    service_states[message.chat.id] = state
    try:
        regions = await get_emu_regions()
    except Exception as error:
        logger.exception("EMU regions loading failed")
        await safe_answer(message, f"Ofislar ro'yxatini olishda xatolik: {error}")
        return

    await safe_answer(
        message,
        "🏢 Ofislar ro'yxati\n\n"
        "📍 Viloyatni tanlang. Keyin shu hududdagi ofislar sahifalab chiqadi.",
        reply_markup=region_reply_keyboard(regions, state),
    )


async def show_calculator_menu(message: Message) -> None:
    state = {"mode": "calculator", "step": "sender_region"}
    service_states[message.chat.id] = state
    try:
        regions = await get_emu_regions()
    except Exception as error:
        logger.exception("EMU regions loading failed")
        await safe_answer(message, f"Kalkulyator ma'lumotlarini olishda xatolik: {error}")
        return

    await safe_answer(
        message,
        "🧮 Kalkulyator\n\n"
        "📦 Jo'natilish nuqtasining viloyatini tanlang.",
        reply_markup=region_reply_keyboard(regions, state),
    )


async def show_ai_assistant(message: Message) -> None:
    service_states[message.chat.id] = {"mode": "ai"}
    await safe_answer(
        message,
        "AI yordamchi bo'limi.\n\n"
        "EMU bo'yicha savolingizni yozing. Masalan:\n"
        "- Samarqand ofislari qayerda?\n"
        "- Andijondan Toshkentga 2 kg qancha?\n"
        "- Do ofisa va Na dom farqi nima?\n\n"
        f"Chiqish uchun {MENU_BACK} tugmasini bosing.",
        reply_markup=reply_keyboard([], add_back=True),
    )


async def emu_callback_handler(callback: CallbackQuery) -> None:
    data = callback.data or ""
    parts = data.split(":")
    chat_id = callback.message.chat.id if callback.message else callback.from_user.id

    if data == "emu:back":
        service_states.pop(chat_id, None)
        await callback.answer()
        if callback.message:
            await safe_edit_text(callback.message, "Asosiy menyuga qaytdingiz.")
            await safe_send_message(callback.bot, chat_id, "Asosiy menyu", reply_markup=main_menu_keyboard())
        return

    try:
        if len(parts) == 3 and parts[1] == "office_region":
            region_id = int(parts[2])
            branches = await get_emu_branches(region_id=region_id)
            regions = await get_emu_regions()
            region = next((item for item in regions if int(item.get("id") or 0) == region_id), {})
            title = f"{localized_name(region)} ofislari"
            await callback.answer("Ofislar yuklandi.")
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    format_branches_list(branches, title),
                    reply_markup=InlineKeyboardMarkup(
                        inline_keyboard=[[InlineKeyboardButton(text=MENU_BACK, callback_data="emu:back")]]
                    ),
                )
            return

        if len(parts) == 3 and parts[1] == "calc_sender_region":
            region_id = int(parts[2])
            state = service_states.setdefault(chat_id, {"mode": "calculator"})
            state.update({"sender_region_id": region_id})
            await callback.answer()
            if region_id == TASHKENT_REGION_ID:
                state.update({"sender_city_id": TASHKENT_CITY_ID, "step": "receiver_region"})
                regions = await get_emu_regions()
                if callback.message:
                    await safe_edit_text(
                        callback.message,
                        "Jo'natilish nuqtasi: Toshkent.\n\nEndi olish nuqtasining viloyatini tanlang.",
                        reply_markup=region_keyboard(regions, "emu:calc_receiver_region"),
                    )
                return
            cities = await get_emu_cities(region_id)
            state["step"] = "sender_city"
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    "Jo'natilish nuqtasining tuman/shahrini tanlang.",
                    reply_markup=city_keyboard(cities, "emu:calc_sender_city"),
                )
            return

        if len(parts) == 3 and parts[1] == "calc_sender_city":
            city_id = int(parts[2])
            state = service_states.setdefault(chat_id, {"mode": "calculator"})
            state.update({"sender_city_id": city_id, "step": "receiver_region"})
            regions = await get_emu_regions()
            await callback.answer()
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    "Endi olish nuqtasining viloyatini tanlang.",
                    reply_markup=region_keyboard(regions, "emu:calc_receiver_region"),
                )
            return

        if len(parts) == 3 and parts[1] == "calc_receiver_region":
            region_id = int(parts[2])
            state = service_states.setdefault(chat_id, {"mode": "calculator"})
            state.update({"receiver_region_id": region_id})
            await callback.answer()
            if region_id == TASHKENT_REGION_ID:
                state.update({"receiver_city_id": TASHKENT_CITY_ID, "step": "service"})
                if callback.message:
                    await safe_edit_text(
                        callback.message,
                        "Olish nuqtasi: Toshkent.\n\nYetkazib berish turini tanlang.",
                        reply_markup=service_keyboard(),
                    )
                return
            cities = await get_emu_cities(region_id)
            state["step"] = "receiver_city"
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    "Olish nuqtasining tuman/shahrini tanlang.",
                    reply_markup=city_keyboard(cities, "emu:calc_receiver_city"),
                )
            return

        if len(parts) == 3 and parts[1] == "calc_receiver_city":
            city_id = int(parts[2])
            state = service_states.setdefault(chat_id, {"mode": "calculator"})
            state.update({"receiver_city_id": city_id, "step": "service"})
            await callback.answer()
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    "Olish turi: ofisgachami yoki uygachami?",
                    reply_markup=service_keyboard(),
                )
            return

        if len(parts) == 3 and parts[1] == "calc_service":
            service_id = int(parts[2])
            state = service_states.setdefault(chat_id, {"mode": "calculator"})
            state.update({"service_id": service_id, "step": "weight"})
            await callback.answer()
            if callback.message:
                await safe_edit_text(
                    callback.message,
                    "Jo'natmaning og'irligini kiriting.\n\n"
                    "Agar gabaritda o'lchangan og'irligi kattaroq bo'lsa, shuni kiriting. Masalan: 1.5",
                )
            return
    except Exception as error:
        logger.exception("EMU callback failed")
        await callback.answer("Xatolik yuz berdi.", show_alert=True)
        if callback.message:
            await safe_answer(callback.message, f"Xatolik: {error}")
        return

    await callback.answer()


async def show_collect_menu(message: Message) -> None:
    await safe_answer(
        message,
        "Jo'natmalarni yig'ish bo'limi.\n\n"
        "Yuridik mijoz - jo'natuvchi ma'lumotlari so'ralmaydi.\n"
        "ФИЗ ЛИЦО - jo'natuvchi ma'lumotlari odatdagidek so'raladi.\n\n"
        "Kerakli yo'nalishni tanlang.",
        reply_markup=collect_menu_keyboard(),
    )


async def show_archive_menu(message: Message) -> None:
    await safe_answer(
        message,
        "Arxiv bo'limi.\n\n"
        "Excel fayl - yig'ilgan mijozlar ro'yxatini yuboradi.\n"
        "Shablon - hozirgi Excel shablonni yuboradi.\n"
        "Ro'yxatni tozalash - Excel ro'yxatini boshidan boshlaydi.",
        reply_markup=archive_menu_keyboard(),
    )


async def show_settings_menu(message: Message) -> None:
    await safe_answer(
        message,
        "Sozlamalar bo'limi.\n\n"
        "Jo'natuvchi sozlamalari - ФИЗ ЛИЦО uchun ma'lumotlarni qayta kiritish.\n"
        "Ruxsat holati - botdan foydalanish ruxsatini ko'rsatadi.",
        reply_markup=settings_menu_keyboard(),
    )


async def handle_menu_message(message: Message) -> bool:
    text = MENU_ALIASES.get((message.text or "").strip(), (message.text or "").strip())
    if not text:
        return False

    if text == MENU_BACK:
        setup_states.pop(message.chat.id, None)
        service_states.pop(message.chat.id, None)
        await send_main_menu(message)
        return True

    if text == MENU_COLLECT:
        await show_collect_menu(message)
        return True

    if text == MENU_OFFICES:
        await show_offices_menu(message)
        return True

    if text == MENU_CALCULATOR:
        await show_calculator_menu(message)
        return True

    if text == MENU_AI_ASSISTANT:
        await show_ai_assistant(message)
        return True

    if text == MENU_ARCHIVE:
        await show_archive_menu(message)
        return True

    if text == MENU_SETTINGS:
        await show_settings_menu(message)
        return True

    if text == MENU_LEGAL:
        await start_legal_setup(message, reset=True)
        return True

    if text == MENU_PHYSICAL:
        await start_setup(message, reset=True, client_type=CLIENT_TYPE_PHYSICAL)
        return True

    if text == MENU_EXCEL_FILE:
        await excel_handler(message)
        return True

    if text == MENU_TEMPLATE_FILE:
        await template_handler(message)
        return True

    if text == MENU_CLEAR:
        await clear_handler(message, start_next_setup=False)
        return True

    if text == MENU_RESET_SETUP:
        await start_setup(message, reset=True, client_type=CLIENT_TYPE_PHYSICAL)
        return True

    if text == MENU_ACCESS_STATUS:
        user_id = message.from_user.id if message.from_user else message.chat.id
        status = "admin" if is_admin(user_id) else "ruxsat berilgan"
        if not ADMIN_IDS:
            status = "ruxsat tekshiruvi o'chirilgan"
        await safe_answer(message, f"Ruxsat holati: {status}", reply_markup=settings_menu_keyboard())
        return True

    return False


async def answer_ai_question(message: Message, question: str) -> None:
    lowered = question.lower()
    context_parts: list[str] = []

    try:
        if any(word in lowered for word in ["ofis", "filial", "office", "branch"]):
            branches = await get_emu_branches()
            query_words = {
                word
                for word in re.findall(r"[\w'`-]+", lowered)
                if len(word) >= 4 and word not in {"ofis", "filial", "office", "branch", "qayerda"}
            }
            matching = [
                branch
                for branch in branches
                if query_words
                and query_words
                & set(
                    re.findall(
                        r"[\w'`-]+",
                        " ".join(
                            [
                                clean_text(branch.get("name")),
                                clean_text(branch.get("address")),
                                clean_text(branch.get("city_name")),
                                clean_text(branch.get("region_name")),
                            ]
                        ).lower(),
                    )
                )
            ]
            if not matching:
                matching = branches[:10]
            context_parts.append(format_branches_list(matching[:10], "Ofislar bo'yicha topilgan ma'lumot", limit=10))

        if any(word in lowered for word in ["viloyat", "tuman", "shahar", "city", "region"]):
            cities = await get_emu_cities()
            sample = [
                f"{city.get('id')}: {localized_name(city)} ({clean_text((city.get('region_name') or {}).get('UZ'))})"
                for city in cities[:60]
            ]
            context_parts.append("EMU shahar/tuman ma'lumotlari:\n" + "\n".join(sample))

        if any(word in lowered for word in ["narx", "kalk", "qancha", "kg", "sum", "so'm", "som"]):
            context_parts.append(
                "Kalkulyator uchun aniq hisoblashda jo'natuvchi shahar/tuman, oluvchi shahar/tuman, "
                "yetkazish turi va og'irlik kerak. Botdagi Kalkulyator bo'limi shu ma'lumotlar asosida EMU API'dan narx oladi."
            )

        if not context_parts:
            context_parts.append(
                "EMU Express saytida ofislar, shahar/tumanlar, filial telefonlari, kalkulyator narxlari, "
                "Do ofisa/Na dom tariflari va indeks ma'lumotlari mavjud."
            )

        context = "\n\n".join(context_parts)[:8000]
        if openai_client is None:
            await safe_answer(message, context[:3500])
            return

        response = await asyncio.to_thread(
            openai_client.responses.create,
            model=OPENAI_MODEL,
            input=[
                {
                    "role": "system",
                    "content": (
                        "Siz EMU botining yordamchisisiz. Faqat berilgan kontekstga tayanib, "
                        "qisqa, amaliy va o'zbek tilida javob bering. Agar ma'lumot yetmasa, "
                        "qaysi bo'limdan foydalanish kerakligini ayting."
                    ),
                },
                {
                    "role": "user",
                    "content": f"Kontekst:\n{context}\n\nSavol:\n{question}",
                },
            ],
        )
        answer = clean_text(getattr(response, "output_text", ""))
        await safe_answer(message, answer or "Javob topilmadi. Savolni aniqroq yozing.")
    except Exception as error:
        logger.exception("AI assistant failed")
        await safe_answer(message, f"AI yordamchida xatolik: {error}")


async def send_offices_page(message: Message, state: dict[str, Any], page: int = 0) -> None:
    branches = state.get("filtered_branches") or state.get("branches") or []
    title = state.get("title") or "Ofislar"
    page_count = max(1, (len(branches) + OFFICES_PAGE_SIZE - 1) // OFFICES_PAGE_SIZE)
    page = max(0, min(page, page_count - 1))
    state["page"] = page
    await safe_answer(
        message,
        format_branches_page(branches, title, page),
        reply_markup=offices_page_keyboard(page, len(branches)),
    )


async def handle_service_text(message: Message) -> bool:
    state = service_states.get(message.chat.id)
    if not state:
        return False

    mode = state.get("mode")
    text = (message.text or "").strip()

    if text == MENU_CANCEL:
        service_states.pop(message.chat.id, None)
        await send_main_menu(message, "❌ Amal bekor qilindi.")
        return True

    if mode == "offices":
        step = state.get("step")
        if step == "region":
            region_id = selected_option_id(state, text)
            if region_id is None:
                await safe_answer(message, "📍 Pastdagi tugmalardan viloyatni tanlang.")
                return True
            try:
                branches = await get_emu_branches(region_id=region_id)
                regions = await get_emu_regions()
                region = next((item for item in regions if int(item.get("id") or 0) == region_id), {})
                state.update(
                    {
                        "step": "browser",
                        "region_id": region_id,
                        "branches": branches,
                        "filtered_branches": branches,
                        "title": f"{localized_name(region)} ofislari",
                        "page": 0,
                    }
                )
                await send_offices_page(message, state, 0)
            except Exception as error:
                logger.exception("Office list failed")
                await safe_answer(message, f"⚠️ Ofislar ro'yxatini olishda xatolik: {error}")
            return True

        if step == "browser":
            if text == MENU_NEXT_PAGE:
                await send_offices_page(message, state, int(state.get("page") or 0) + 1)
                return True
            if text == MENU_PREV_PAGE:
                await send_offices_page(message, state, int(state.get("page") or 0) - 1)
                return True
            if text == MENU_SEARCH:
                state["step"] = "search"
                await safe_answer(
                    message,
                    "🔎 Qidirish uchun tuman, filial nomi yoki manzil bo'lagini yozing.\n"
                    "Masalan: Sergeli, Olmazor, Qoraqamish",
                    reply_markup=reply_keyboard([MENU_CANCEL], row_size=1),
                )
                return True
            await safe_answer(message, "Pastdagi tugmalardan amal tanlang.")
            return True

        if step == "search":
            filtered = filter_branches(state.get("branches") or [], text)
            state["filtered_branches"] = filtered
            state["title"] = f"{state.get('title', 'Ofislar')} | qidiruv: {text}"
            state["step"] = "browser"
            await send_offices_page(message, state, 0)
            return True

    if mode == "ai":
        await answer_ai_question(message, text)
        return True

    if mode == "calculator" and state.get("step") != "weight":
        step = state.get("step")
        try:
            if step == "sender_region":
                region_id = selected_option_id(state, text)
                if region_id is None:
                    await safe_answer(message, "📍 Jo'natilish viloyatini pastdagi tugmalardan tanlang.")
                    return True
                state["sender_region_id"] = region_id
                if region_id == TASHKENT_REGION_ID:
                    state.update({"sender_city_id": TASHKENT_CITY_ID, "step": "receiver_region"})
                    regions = await get_emu_regions()
                    await safe_answer(
                        message,
                        "✅ Jo'natilish nuqtasi: Toshkent.\n\n📍 Endi olish nuqtasining viloyatini tanlang.",
                        reply_markup=region_reply_keyboard(regions, state),
                    )
                    return True
                cities = await get_emu_cities(region_id)
                state["step"] = "sender_city"
                await safe_answer(
                    message,
                    "🏙 Jo'natilish nuqtasining tuman/shahrini tanlang.",
                    reply_markup=city_reply_keyboard(cities, state),
                )
                return True

            if step == "sender_city":
                city_id = selected_option_id(state, text)
                if city_id is None:
                    await safe_answer(message, "🏙 Tuman/shaharni pastdagi tugmalardan tanlang.")
                    return True
                state.update({"sender_city_id": city_id, "step": "receiver_region"})
                regions = await get_emu_regions()
                await safe_answer(
                    message,
                    "📍 Endi olish nuqtasining viloyatini tanlang.",
                    reply_markup=region_reply_keyboard(regions, state),
                )
                return True

            if step == "receiver_region":
                region_id = selected_option_id(state, text)
                if region_id is None:
                    await safe_answer(message, "📍 Olish viloyatini pastdagi tugmalardan tanlang.")
                    return True
                state["receiver_region_id"] = region_id
                if region_id == TASHKENT_REGION_ID:
                    state.update({"receiver_city_id": TASHKENT_CITY_ID, "step": "service"})
                    await safe_answer(
                        message,
                        "✅ Olish nuqtasi: Toshkent.\n\n🚚 Yetkazib berish turini tanlang.",
                        reply_markup=calculator_service_reply_keyboard(),
                    )
                    return True
                cities = await get_emu_cities(region_id)
                state["step"] = "receiver_city"
                await safe_answer(
                    message,
                    "🏙 Olish nuqtasining tuman/shahrini tanlang.",
                    reply_markup=city_reply_keyboard(cities, state),
                )
                return True

            if step == "receiver_city":
                city_id = selected_option_id(state, text)
                if city_id is None:
                    await safe_answer(message, "🏙 Olish tuman/shahrini pastdagi tugmalardan tanlang.")
                    return True
                state.update({"receiver_city_id": city_id, "step": "service"})
                await safe_answer(
                    message,
                    "🚚 Olish turi: ofisgachami yoki uygachami?",
                    reply_markup=calculator_service_reply_keyboard(),
                )
                return True

            if step == "service":
                if text == MENU_DO_OFFICE or "ДО ОФИСА" in text:
                    service_id = 1
                elif text == MENU_TO_HOME or "НА ДОМ" in text:
                    service_id = 3
                else:
                    await safe_answer(message, "🚚 Yetkazib berish turini pastdagi tugmalardan tanlang.")
                    return True
                state.update({"service_id": service_id, "step": "weight"})
                await safe_answer(
                    message,
                    "⚖️ Jo'natmaning og'irligini kiriting.\n\n"
                    "Agar gabaritda o'lchangan og'irligi kattaroq bo'lsa, shuni kiriting. Masalan: 1.5",
                    reply_markup=reply_keyboard([], add_back=True),
                )
                return True
        except Exception as error:
            logger.exception("Calculator step failed")
            await safe_answer(message, f"⚠️ Kalkulyator xatoligi: {error}")
            return True

    if mode == "calculator" and state.get("step") == "weight":
        normalized = text.replace(",", ".")
        try:
            weight = float(normalized)
        except ValueError:
            await safe_answer(message, "Og'irlikni raqamda kiriting. Masalan: 1.5")
            return True
        if weight <= 0:
            await safe_answer(message, "Og'irlik 0 dan katta bo'lishi kerak.")
            return True

        try:
            sender_city_id = int(state["sender_city_id"])
            receiver_city_id = int(state["receiver_city_id"])
            service_id = int(state["service_id"])
            result = await calculate_emu_delivery(sender_city_id, receiver_city_id, weight, service_id)
            receiver_branches = await get_emu_branches(city_id=receiver_city_id)
            if not receiver_branches and state.get("receiver_region_id"):
                receiver_branches = await get_emu_branches(region_id=int(state["receiver_region_id"]))
            await safe_answer(
                message,
                format_calculator_result(result, service_id, receiver_branches),
                reply_markup=main_menu_keyboard(),
            )
            service_states.pop(message.chat.id, None)
        except Exception as error:
            logger.exception("Calculator failed")
            await safe_answer(message, f"Kalkulyator xatoligi: {error}")
        return True

    return False


async def clear_handler(message: Message, start_next_setup: bool = True) -> None:
    if not await ensure_user_access(message):
        return
    async with excel_lock:
        reset_excel_file()
    sender_sessions.pop(message.chat.id, None)
    setup_states.pop(message.chat.id, None)
    await safe_answer(
        message,
        "Ro'yxat tozalandi. Yangi Excel fayl shablondan tayyorlanadi.",
        reply_markup=archive_menu_keyboard() if not start_next_setup else None,
    )
    if start_next_setup:
        await start_setup(message)


async def text_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    text = message.text or ""
    normalized_menu_text = MENU_ALIASES.get(text.strip(), text.strip())
    if normalized_menu_text in MENU_TEXTS:
        setup_states.pop(message.chat.id, None)
        if normalized_menu_text != MENU_BACK:
            service_states.pop(message.chat.id, None)
        await handle_menu_message(message)
        return

    if await handle_service_text(message):
        return

    if await handle_setup_message(message):
        return

    if await handle_menu_message(message):
        return

    if message.chat.id not in sender_sessions:
        await show_collect_menu(message)
        return

    await enqueue_batch_item(BatchItem(kind="text", message=message, text=text))


async def photo_handler(message: Message, bot: Bot) -> None:
    if not await ensure_user_access(message):
        return
    if message.chat.id in setup_states:
        await safe_answer(message, "Hozir jo'natuvchi ma'lumotlarini matn ko'rinishida kiriting.")
        return

    if message.chat.id not in sender_sessions:
        await show_collect_menu(message)
        return

    photo = message.photo[-1]
    await enqueue_batch_item(
        BatchItem(
            kind="image",
            message=message,
            file_id=photo.file_id,
            mime_type="image/jpeg",
            bot=bot,
        )
    )


async def document_image_handler(message: Message, bot: Bot) -> None:
    if not await ensure_user_access(message):
        return
    if message.chat.id in setup_states:
        await safe_answer(message, "Hozir jo'natuvchi ma'lumotlarini matn ko'rinishida kiriting.")
        return

    if message.chat.id not in sender_sessions:
        await show_collect_menu(message)
        return

    document = message.document
    if document is None or not (document.mime_type or "").startswith("image/"):
        await safe_answer(message, "Iltimos, rasm yoki mijoz ma'lumotlari yozilgan matn yuboring.")
        return

    await enqueue_batch_item(
        BatchItem(
            kind="image",
            message=message,
            file_id=document.file_id,
            mime_type=document.mime_type or "image/jpeg",
            bot=bot,
        )
    )


async def unsupported_handler(message: Message) -> None:
    if not await ensure_user_access(message):
        return
    await safe_answer(message, "Matn yoki rasm yuboring. Yordam uchun /help buyrug'ini bosing.")


async def setup_bot_commands(bot: Bot) -> None:
    await bot.set_my_commands(
        [
            BotCommand(command="start", description="Botni ishga tushirish"),
            BotCommand(command="setup", description="ФИЗ ЛИЦО jo'natuvchi sozlamalari"),
            BotCommand(command="ofislar", description="EMU ofislar ro'yxati"),
            BotCommand(command="kalkulyator", description="Yetkazib berish narxini hisoblash"),
            BotCommand(command="ai", description="EMU bo'yicha AI yordamchi"),
            BotCommand(command="help", description="Foydalanish bo'yicha yordam"),
            BotCommand(command="excel", description="Excel faylni yuborish"),
            BotCommand(command="shablon", description="Excel shablonni yuborish"),
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

    dispatcher.callback_query.register(access_callback_handler, F.data.startswith("access:"))
    dispatcher.callback_query.register(emu_callback_handler, F.data.startswith("emu:"))
    dispatcher.callback_query.register(setup_callback_handler, F.data.startswith("setup:"))
    dispatcher.message.register(start_handler, Command("start"))
    dispatcher.message.register(setup_handler, Command("setup"))
    dispatcher.message.register(offices_handler, Command("ofislar"))
    dispatcher.message.register(calculator_handler, Command("kalkulyator"))
    dispatcher.message.register(ai_handler, Command("ai"))
    dispatcher.message.register(help_handler, Command("help"))
    dispatcher.message.register(excel_handler, Command("excel"))
    dispatcher.message.register(template_handler, Command("shablon"))
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
