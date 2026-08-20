"""Uzbekiston shahar/tumanlarini tanib, EMU server formatiga keltirish.

Kelib tushgan matn (Excel, rasm, matn xabar) qanday yozilgan bo'lsa ham -
lotin, kirill, ruscha yoki eski nom bilan - shu modul uni EMU справочникdagi
yagona shahar/tuman yozuviga (`Place.server`) va shu tumandagi ofis nomlariga
bog'lab beradi.

Ma'lumot manbasi: `location_data.py` (EMU API dan avtomatik yaratiladi,
`python tools/refresh_locations.py`). Qo'lda kiritilgan qo'shimcha nomlar
(`Zafar`, `Toʻytepa`, `Chelak` ...) shu fayldagi `EXTRA_NAMES` va `WEAK_NAMES`
jadvallarida turadi.

Aniqlash tartibi:

1. Aniq moslik   - to'liq nom ("Denov tumani"), qisqa nom ("Denov"), taxallus.
2. Fonetik moslik - lotin/kirill/ruscha imlo farqlari ("Khiva" = "Xiva").
3. Taxminiy moslik - xato yozilgan nomlar uchun ("Deneu" -> "Denov").
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import get_close_matches

from location_data import CITIES, REGIONS

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
        "ъ": "",
        "і": "i",
        "ї": "i",
        "є": "e",
    }
)

APOSTROPHES = ("ʼ", "ʻ", "‘", "’", "'", "`", "´", "ʹ")
LETTER_REPLACEMENTS = (("ģ", "g"), ("ğ", "g"), ("ŏ", "o"), ("ō", "o"), ("ū", "u"))

# Tuman/shahar ko'rsatkichlari - nomdan keyin kelsa moslikni kuchaytiradi.
DISTRICT_MARKERS = {
    "tumani",
    "tuman",
    "tumanida",
    "tumaniga",
    "tumandagi",
    "rayon",
    "rayoni",
    "rayona",
    "raion",
    "shahri",
    "shahar",
    "shahrida",
    "shahriga",
    "shaxri",
    "shaxar",
    "shaharchasi",
    "shaharcha",
    "gorod",
    "goroda",
    "gorodok",
}

# Nomdan oldin kelib, shahar/tumanni ko'rsatadigan so'zlar.
DISTRICT_PREFIX_MARKERS = {"gorod", "goroda", "g", "gor", "sh", "shahri", "tumani", "pgt"}

# Viloyat ko'rsatkichlari - bunday nom tuman emas, viloyat sifatida o'qiladi.
REGION_MARKERS = {
    "viloyati",
    "viloyat",
    "viloyatida",
    "viloyatining",
    "vil",
    "oblast",
    "oblasti",
    "oblastn",
    "obl",
    "respublikasi",
    "respublika",
    "respublikasida",
    "resp",
}

# Ko'cha/uy ko'rsatkichlari - nom manzil qismi bo'lsa, uni hudud deb olmaymiz.
STREET_SUFFIX_MARKERS = {
    "kochasi",
    "kocha",
    "koch",
    "kochada",
    "kochasida",
    "yoli",
    "yolida",
    "yol",
    "shox",
    "shoh",
    "tor",
    "tupigi",
    "berk",
    "massivi",
    "mavzesi",
    "dahasi",
    "kvartali",
    "kvartal",
    "bozori",
    "bekati",
    "metrosi",
    "koprigi",
    "maktabi",
    "shifoxonasi",
    "majmuasi",
    "uy",
    "uyi",
    "korpusi",
}
STREET_PREFIX_MARKERS = {
    # kirillcha "улица" translit qilinganda "ulisa" bo'ladi (ц -> s), shuning uchun ikki xil yozuv ham bor.
    "ulitsa",
    "ulisa",
    "ulisasi",
    "ul",
    "prospekt",
    "prospekti",
    "pr",
    "prt",
    "massiv",
    "massivi",
    "mavze",
    "daha",
    "kvartal",
    "kvartira",
    "kv",
    "dom",
    "pereulok",
    "shosse",
    "tupik",
    "proezd",
    "bulvar",
    "naberejnaya",
    "mikrorayon",
    "mkr",
    "korpus",
    "xonadon",
    "poselok",
    "posyolok",
    "posilok",
    "gorodskoy",
    "gorodskaya",
    "selo",
    "sovxoz",
    "kolxoz",
}

# Excel/matnda uchraydigan, hudud nomi bo'lib ko'rinadigan lekin hudud emas so'zlar.
BASE_NAME_BLOCKLIST = {"uzbekiston", "ozbekiston", "respublika", "markaz"}

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

FULL_WEIGHT = 9
ALIAS_WEIGHT = 9
BASE_WEIGHT = 4
PHONETIC_WEIGHT = 3
FUZZY_WEIGHT = 1

WINDOW_SCORE = 12
DISTRICT_MARKER_BONUS = 10
REGION_HINT_BONUS = 14
STREET_PENALTY = -40
MAX_WINDOW_WORDS = 4
FUZZY_CUTOFF = 0.86

# Qo'shimcha nomlar: EMU ro'yxatida yo'q, lekin hayotda ishlatiladigan yozuvlar.
# Kalit - `location_data.CITIES` dagi o'zbekcha nom.
EXTRA_NAMES: dict[str, tuple[str, ...]] = {
    # Toshkent shahri
    "Toshkent": ("Toshkent shahri", "Toshkent shahar", "Tashkent city", "Toshkent sh"),
    "Mirabad tumani": ("Mirobod", "Mirobod tumani"),
    "Shayxontohur tumani": ("Shayxontoxur", "Shaykhantakhur", "Shayhontohur"),
    "Yashnobod tumani": ("Hamza", "Hamza tumani"),
    "Yakkasaroy tumani": ("Yakkasarai",),
    "Yangihayot tumani": ("Yangihayat",),
    # Andijon
    "Andijon tumani": ("Kuyganyor", "Quyganyor", "Kuyganyar", "Kuygan yor"),
    "Izboskan tumani": ("Poytug", "Poytugʻ", "Izbosgan"),
    "Jalaquduq tumani": ("Jalolquduq", "Jalakuduk"),
    "Xonobod shahri": ("Qorasuv", "Karasuu", "Qora suv", "Xonobod"),
    "Ulug‘nor tumani": ("Ulugnor",),
    "Asaka tumani": ("Leninsk",),
    "Bo‘z tumani": ("Buz", "Boz shahri"),
    "Xo‘jaobod tumani": ("Hojiobod", "Xojaobod"),
    # Buxoro
    "Buxoro tumani": ("Galaosiyo", "Gʻalaosiyo", "Gallaosiyo", "Karvonbozor"),
    "Kogon tumani": ("Kogon shahri", "Kagan"),
    "G‘ijduvon tumani": ("Gijduvon shahri",),
    # Jizzax
    "Forish tumani": ("Yangiqishloq", "Yangiqishloq shahri"),
    "Mirzacho‘l tumani": ("Mirzachol", "Mirzachoʻl", "Gagarin shahri"),
    "Yangiobod tumani": ("Yangiobod Jizzax",),
    "Zomin tumani": ("Zaamin",),
    # Qashqadaryo
    "Qarshi tumani": ("Beshkent shahri",),
    "Kasbi tumani": ("Muglan", "Mugʻlon"),
    "Mirishkor tumani": ("Yangi Mirishkor", "Pomuq"),
    "Shahrisabz tumani": ("Shahrisabz shahri",),
    "Dehqonobod tumani": ("Dehqonobod Qashqadaryo", "Karashina"),
    # Navoiy
    "Karmana tumani": ("Kermine", "Navoiy tumani"),
    "Nurota tumani": ("Gʻozgʻon", "Gozgon", "Gazgan"),
    "Xatirchi tumani": ("Yangirabot", "Yangirabod"),
    "Tomdi tumani": ("Tomdibuloq", "Tomdi shahri"),
    # Namangan
    "Mingbuloq tumani": ("Jumashuy", "Jomashuy"),
    "Norin tumani": ("Xaqqulobod", "Haqqulobod"),
    "Namangan tumani": ("Toshbuloq",),
    # Qoraqalpogʻiston
    "Amudaryo tumani": ("Mangʻit", "Mangit"),
    "Ellikqala tumani": ("Boʻston", "Boston", "Ellikqalʼa", "Ellikqala shahri"),
    "Taqiyatosh tumani": ("Taxiatosh", "Taxtiatosh"),
    "Bo‘zatov tumani": ("Bozatov", "Qazaqketken", "Qozoqketken"),
    "Nukus tumani": ("Aqmangit", "Akmangit"),
    "Mo‘ynoq tumani": ("Moynoq", "Muynak"),
    # Samarqand
    "Payariq tumani": ("Chelak", "Chelek", "Payariq shahri"),
    "Kattaqo‘rg‘on tumani": ("Payshanba", "Kattaqorgon shahri"),
    "Paxtachi tumani": ("Ziyovuddin", "Paxtachi"),
    "Narpay tumani": ("Oqtosh", "Oktosh"),
    "Oqdaryo tumani": ("Loyish", "Layish", "Oqdaryo"),
    "Pastdarg‘om tumani": ("Juma", "Jumabozor"),
    "Samarqand shahri": ("Samarqant",),
    "Samarqand tumani": ("Gulobod", "Gulbahor", "Seliski"),
    "Qo‘shrabot tumani": ("Qoshrabod", "Qoshrabot"),
    "Ishtixon tumani": ("Ishtihon",),
    # Surxondaryo
    "Denov tumani": ("Denau", "Denov", "Denou", "Denov shahri", "Beshkapa", "Denov Beshkapa"),
    "Oltinsoy tumani": ("Qorluq", "Korluk", "Oltinsoy", "Oltinsay"),
    "Muzrabot tumani": ("Xalqobod", "Halqabad", "Muzrabod"),
    "Termiz tumani": ("Uchqizil", "Uchkizil"),
    "Sho‘rchi tumani": ("Shorchi", "Shurchi"),
    # Sirdaryo
    "Sirdaryo tumani": ("Baxt", "Bakht", "Sirdaryo shahri"),
    "Guliston tumani": ("Dehqonobod Sirdaryo",),
    "Mirzaobod tumani": ("Navruz", "Navroʻz", "Mirzaobod"),
    "Sayxunobod tumani": ("Sayxun", "Saykhun"),
    "Xovos tumani": ("Xavast", "Xovos"),
    "Boyovut tumani": ("Boyovut", "Bayaut"),
    # Toshkent viloyati
    "O‘rtachirchiq tumani": ("Toʻytepa", "Toytepa", "Nurafshon shahri", "Ortachirchiq"),
    "Quyichirchiq tumani": ("Doʻstobod", "Dostobod", "Quyi Chirchiq"),
    "Yuqorichirchiq tumani": ("Yuqori Chirchiq", "Yangibozor Toshkent"),
    "Bo‘stonliq tumani": ("Gʻazalkent", "Gazalkent", "Bostonliq", "Chorvoq", "Chimyon"),
    "Zangiota tumani": ("Eshonguzar", "Zangiota shahri"),
    "Bekobod tumani": ("Zafar", "Bekobod shahri"),
    "Toshkent tumani": ("Keles", "Kelas"),
    "Ohangaron tumani": ("Axangaron", "Ohangaron shahri", "Ahangaran"),
    "Yangiyo‘l tumani": ("Yangiyol", "Yangiyul", "Yangiyoʻl shahri"),
    "Akkurgan tumani": ("Oqqoʻrgʻon", "Oqqorgon"),
    "Pskent tumani": ("Piskent",),
    # Farg‘ona
    "Furqat tumani": ("Furqat", "Navbahor Fargona"),
    "Buvayda tumani": ("Ibrat", "Buvayda shahri"),
    "O‘zbekiston tumani": ("Yaypan shahri",),
    "Farg‘ona tumani": ("Vodil", "Vuadil"),
    "Kuva tumani": ("Quva", "Quva shahri"),
    "Qo‘qon shahri": ("Quqon", "Qoqon", "Kokand"),
    "Marg‘ilon shahri": ("Margilon", "Margilan"),
    "So‘x tumani": ("Sox", "Sokh"),
    "Dangara tumani": ("Dangʻara", "Dangara shahri"),
    # Xorazm
    "Tuproqqal’a tumani": ("Pitnak", "Tuproqqala", "Turpoqqala"),
    "Xazorasp tumani": ("Hazorasp", "Hazarasp"),
    "Urganch tumani": ("Qorovul", "Korovul"),
    "Yangibozor tumani": ("Yangibozor Xorazm",),
    "Shovot tumani": ("Shovot shahri",),
}

# Ikki xil hududda uchraydigan yoki manzil ichida chalkashishi mumkin bo'lgan nomlar:
# aniq nomlardan kuchsizroq baholanadi.
WEAK_NAMES: dict[str, tuple[str, ...]] = {
    # Oqoltin: Sirdaryodagi tuman (asosiy) va Andijon Ulug'nor markazi.
    "Ulug‘nor tumani": ("Oqoltin", "Oq oltin"),
    # Dehqonobod: Qashqadaryodagi tuman (asosiy) va Sirdaryo Guliston tumani markazi.
    "Guliston tumani": ("Dehqonobod", "Dehkanabad"),
    # Yangibozor: Xorazmdagi tuman (asosiy) va Toshkent viloyati Yuqorichirchiq markazi.
    "Yuqorichirchiq tumani": ("Yangibozor",),
    "Baliqchi tumani": ("Chinobod", "Chinabad"),
    "Samarqand shahri": ("Samarqand Qorasuv",),
}


def transliterate_location_text(value: str) -> str:
    return str(value or "").casefold().translate(CYRILLIC_TO_LATIN)


def _clean_latin(value: str) -> str:
    text = transliterate_location_text(value)
    for mark in APOSTROPHES:
        text = text.replace(mark, "")
    for source, target in LETTER_REPLACEMENTS:
        text = text.replace(source, target)
    return text


def normalize_location_key(value: str) -> str:
    """Har qanday yozuvni taqqoslash uchun yagona kalitga keltiradi."""

    return re.sub(r"[^a-z0-9]+", "", _clean_latin(value))


def location_words(value: str) -> list[str]:
    """Matnni tartibi saqlangan so'zlarga ajratadi (ko'rsatkich so'zlar ham qoladi)."""

    return re.findall(r"[a-z0-9]+", _clean_latin(value))


def location_tokens(value: str) -> list[str]:
    """Hudud nomiga aloqasi yo'q so'zlarni tashlab, qolganini qaytaradi."""

    normalized = _clean_latin(value).replace("ozbekiston", "uzbekiston")
    tokens = re.findall(r"[a-z0-9]+", normalized)
    return [token for token in tokens if token not in LOCATION_STOP_WORDS and len(token) > 1]


_PHONETIC_DIGRAPHS = (
    ("shch", "sh"),
    ("sch", "sh"),
    ("dzh", "j"),
    ("dj", "j"),
    ("kh", "h"),
    ("ph", "f"),
    ("ts", "s"),
)
_PHONETIC_LETTERS = (("x", "h"), ("q", "k"), ("w", "v"), ("c", "s"), ("y", "i"), ("e", "i"), ("o", "u"))


def phonetic_key(value: str) -> str:
    """Lotin/kirill/ruscha imlo farqlarini yo'qotadigan taqqoslash kaliti."""

    key = value if value.isalnum() else normalize_location_key(value)
    for source, target in _PHONETIC_DIGRAPHS:
        key = key.replace(source, target)
    for source, target in _PHONETIC_LETTERS:
        key = key.replace(source, target)
    return re.sub(r"(.)\1+", r"\1", key)


_RU_ADJECTIVE_KEY_TAILS = ("skiy", "skoy", "skaya", "skie", "skogo", "skom", "skoi", "skix", "tskiy", "skogo")


def _ru_adjective_stem_key(key: str) -> str:
    """'andijanskiy' -> 'andijan' (ruscha sifat shaklidan asosiy nomni ajratadi)."""

    for tail in _RU_ADJECTIVE_KEY_TAILS:
        if key.endswith(tail) and len(key) - len(tail) >= 4:
            return key[: -len(tail)]
    return ""


@dataclass(frozen=True)
class Region:
    region_id: int
    uz: str
    ru: str
    center_city_id: int


@dataclass(frozen=True)
class Place:
    """EMU справочникdagi bitta shahar yoki tuman."""

    city_id: int
    region_id: int
    uz: str
    ru: str
    server: str
    offices: tuple[str, ...]
    is_district: bool
    is_region_center: bool

    @property
    def region(self) -> Region:
        return REGION_BY_ID[self.region_id]

    def display_name(self) -> str:
        return f"{self.server} ({self.uz})"


@dataclass(frozen=True)
class LocationMatch:
    """Matndan aniqlangan hudud."""

    server: str = ""
    place: Place | None = None
    region: Region | None = None
    note: str = ""
    approximate: bool = False
    score: int = 0

    def __bool__(self) -> bool:
        return bool(self.server)


REGION_BY_ID: dict[int, Region] = {}
PLACES: list[Place] = []
PLACE_BY_ID: dict[int, Place] = {}
PLACE_BY_SERVER: dict[str, Place] = {}

_name_index: dict[str, dict[int, int]] = {}
_keys_by_place: dict[int, set[str]] = {}
_phonetic_index: dict[str, int] = {}
_phonetic_blocked: set[str] = set()
_region_index: dict[str, int] = {}

REGION_EXTRA_NAMES: dict[int, tuple[str, ...]] = {
    1: ("Andijan", "Andijon"),
    2: ("Bukhara", "Buxoro"),
    3: ("Fergana", "Fargona", "Fargʻona"),
    4: ("Jizzakh", "Djizak", "Jizzax"),
    5: ("Kashkadarya", "Qashqadaryo", "Kashkadaryo"),
    6: ("Khorezm", "Xorazm", "Horazm"),
    7: ("Namangan",),
    8: ("Navoi", "Navoiy"),
    9: (
        "Qoraqalpogʻiston",
        "Qoraqalpogiston",
        "Qoraqalpoqstan",
        "Karakalpakstan",
        "Karakalpakiya",
        "Qoraqalpogʻiston Respublikasi",
    ),
    10: ("Samarkand", "Samarqand"),
    11: ("Surkhandarya", "Surxondaryo", "Surhondaryo"),
    12: ("Syrdarya", "Sirdaryo"),
    13: ("Toshkent shahri", "Tashkent city"),
    14: ("Tashkent region", "Toshkent viloyat"),
}

_RU_ADJECTIVE_SUFFIXES = ("ский", "ской", "ская", "ские", "цкий", "ий", "ый", "ой", "ая")
_UZ_TAIL_WORDS = ("tumani", "shahri", "shahar", "shaharchasi", "respublikasi", "viloyati")


def _uz_base_name(name: str) -> str:
    words = name.split()
    while words and normalize_location_key(words[-1]) in _UZ_TAIL_WORDS:
        words.pop()
    return " ".join(words)


def _ru_base_name(name: str) -> str:
    text = re.sub(r"^\s*(город|г\.|гор\.|пос\.|пгт)\s+", "", name.strip(), flags=re.IGNORECASE)
    text = re.sub(r"\s*(район|области|область|города)\s*$", "", text, flags=re.IGNORECASE).strip()
    lowered = text.casefold()
    for suffix in _RU_ADJECTIVE_SUFFIXES:
        if lowered.endswith(suffix) and len(lowered) > len(suffix) + 2:
            return text[: -len(suffix)]
    return text


def _add_name(place_id: int, name: str, weight: int) -> None:
    key = normalize_location_key(name)
    if len(key) < 3 or key.isdigit():
        return
    if weight <= BASE_WEIGHT and key in BASE_NAME_BLOCKLIST:
        return
    slot = _name_index.setdefault(key, {})
    if slot.get(place_id, 0) < weight:
        slot[place_id] = weight
    _keys_by_place.setdefault(place_id, set()).add(key)


def _add_phonetic(place_id: int, name: str) -> None:
    """Fonetik kalitni indeksga qo'shadi; ikki xil hududga to'g'ri kelsa - o'chiradi."""

    key = phonetic_key(normalize_location_key(name))
    if len(key) < 4 or key in _phonetic_blocked:
        return
    current = _phonetic_index.get(key)
    if current is None:
        _phonetic_index[key] = place_id
    elif current != place_id:
        del _phonetic_index[key]
        _phonetic_blocked.add(key)


def _register_region_name(name: str, region_id: int, override: bool) -> None:
    key = normalize_location_key(name)
    if len(key) < 3:
        return
    if override or key not in _region_index:
        _region_index[key] = region_id


def _sorted_offices(uz_name: str, offices: tuple[str, ...]) -> tuple[str, ...]:
    """Tuman nomi bilan atalgan ofis birinchi bo'ladi (manzil aniq bo'lmasa ham to'g'ri kod chiqadi)."""

    base_key = normalize_location_key(_uz_base_name(uz_name))

    def rank(name: str) -> int:
        key = normalize_location_key(name)
        if base_key and key == base_key:
            return 0
        if base_key and base_key in key:
            return 1
        return 2

    return tuple(sorted(offices, key=rank))


def _build() -> None:
    for region_id, uz, ru, center_city_id in REGIONS:
        REGION_BY_ID[region_id] = Region(region_id=region_id, uz=uz, ru=ru, center_city_id=center_city_id)

    centers = {region.center_city_id for region in REGION_BY_ID.values()}
    for city_id, region_id, uz, ru, server, offices in CITIES:
        place = Place(
            city_id=city_id,
            region_id=region_id,
            uz=uz,
            ru=ru,
            server=server,
            offices=_sorted_offices(uz, offices),
            is_district=normalize_location_key(uz.split()[-1]) == "tumani",
            is_region_center=city_id in centers,
        )
        PLACES.append(place)
        PLACE_BY_ID[city_id] = place
        PLACE_BY_SERVER.setdefault(server, place)

    for place in PLACES:
        for name in (place.uz, place.ru, place.server):
            _add_name(place.city_id, name, FULL_WEIGHT)
            _add_phonetic(place.city_id, name)
        for name in (_uz_base_name(place.uz), _ru_base_name(place.ru), _ru_base_name(place.server)):
            if name:
                _add_name(place.city_id, name, BASE_WEIGHT)
                _add_phonetic(place.city_id, name)

    for uz_name, names in EXTRA_NAMES.items():
        place = _place_by_uz(uz_name)
        for name in names:
            _add_name(place.city_id, name, ALIAS_WEIGHT)
            _add_phonetic(place.city_id, name)

    for uz_name, names in WEAK_NAMES.items():
        place = _place_by_uz(uz_name)
        for name in names:
            _add_name(place.city_id, name, BASE_WEIGHT)

    for region in REGION_BY_ID.values():
        _register_region_name(region.uz, region.region_id, override=True)
        _register_region_name(region.ru, region.region_id, override=True)
    for region in REGION_BY_ID.values():
        for name in (_uz_base_name(region.uz), _ru_base_name(region.ru)):
            if name:
                _register_region_name(name, region.region_id, override=False)
    for region_id, names in REGION_EXTRA_NAMES.items():
        for name in names:
            _register_region_name(name, region_id, override=False)


_uz_lookup: dict[str, Place] = {}


def _place_by_uz(uz_name: str) -> Place:
    if not _uz_lookup:
        _uz_lookup.update({normalize_location_key(place.uz): place for place in PLACES})
    place = _uz_lookup.get(normalize_location_key(uz_name))
    if place is None:
        raise KeyError(f"location_data.CITIES ichida '{uz_name}' topilmadi")
    return place


_build()

SERVER_LOCATIONS: tuple[str, ...] = tuple(dict.fromkeys(place.server for place in PLACES))
_fuzzy_keys: tuple[str, ...] = tuple(_phonetic_index)


def place_for_server(value: str) -> Place | None:
    place = PLACE_BY_SERVER.get(str(value or "").strip())
    if place is not None:
        return place
    key = normalize_location_key(value)
    for candidate in PLACES:
        if normalize_location_key(candidate.server) == key:
            return candidate
    return None


def name_keys_for_server(value: str) -> set[str]:
    """Shu hududga tegishli barcha nom kalitlari (filial nomini solishtirish uchun)."""

    place = place_for_server(value)
    if place is None:
        key = normalize_location_key(value)
        return {key} if key else set()
    return set(_keys_by_place.get(place.city_id, ()))


def offices_for_server(value: str) -> tuple[str, ...]:
    place = place_for_server(value)
    return place.offices if place else ()


def region_center_server(value: str) -> str:
    place = place_for_server(value)
    if place is None:
        return str(value or "")
    return PLACE_BY_ID[place.region.center_city_id].server


def region_offices_for_server(value: str) -> tuple[str, ...]:
    """Viloyatdagi barcha ofis nomlari: markaz ofislari birinchi bo'ladi."""

    place = place_for_server(value)
    if place is None:
        return ()
    center_id = place.region.center_city_id
    names: list[str] = list(PLACE_BY_ID[center_id].offices)
    for other in PLACES:
        if other.region_id == place.region_id and other.city_id != center_id:
            names.extend(other.offices)
    return tuple(dict.fromkeys(names))


def _windows(words: list[str]):
    total = len(words)
    for size in range(min(MAX_WINDOW_WORDS, total), 0, -1):
        for start in range(0, total - size + 1):
            yield start, size, "".join(words[start : start + size])


def _region_hints(words: list[str]) -> set[int]:
    hints: set[int] = set()
    for start, size, key in _windows(words):
        region_id = _region_index.get(key)
        if region_id is None:
            continue
        following = words[start + size] if start + size < len(words) else ""
        if following in DISTRICT_MARKERS:
            # "Toshkent tumani" - bu viloyat emas, tuman nomi.
            continue
        if size > 1 or following in REGION_MARKERS or normalize_location_key(REGION_BY_ID[region_id].uz) == key:
            hints.add(region_id)
    return hints


def _window_score(words: list[str], start: int, size: int, weight: int, place: Place, hints: set[int]) -> int:
    previous = words[start - 1] if start > 0 else ""
    following = words[start + size] if start + size < len(words) else ""
    score = WINDOW_SCORE * size + weight
    if following in DISTRICT_MARKERS or previous in DISTRICT_PREFIX_MARKERS:
        score += DISTRICT_MARKER_BONUS
    if following in STREET_SUFFIX_MARKERS or previous in STREET_PREFIX_MARKERS:
        score += STREET_PENALTY
    if place.region_id in hints:
        score += REGION_HINT_BONUS
    return score


def _rank_key(score: int, place: Place, approximate: bool) -> tuple:
    return (
        score,
        0 if approximate else 1,
        1 if place.is_region_center else 0,
        0 if place.is_district else 1,
        -place.city_id,
    )


def _prefer_district_over_center(
    winner: LocationMatch,
    by_place: dict[int, tuple[tuple, LocationMatch, tuple[int, int]]],
) -> LocationMatch:
    """Matnda ham viloyat markazi, ham o'sha viloyatning tumani bo'lsa - tuman tanlanadi.

    Masalan "Yunusobod 12-kvartal, Toshkent" -> Юнусабад (Ташкент emas).
    Ikkalasi matnning bir joyidan topilgan bo'lsa ("Buxoro shahri" -> Buxoro shahri /
    Buxoro tumani) markaz qoladi.
    """

    if winner.place is None or not winner.place.is_region_center:
        return winner

    winner_window = by_place.get(winner.place.city_id, (None, None, None))[2]
    if winner_window is None:
        return winner

    def overlaps(window: tuple[int, int]) -> bool:
        start, size = window
        winner_start, winner_size = winner_window
        return start < winner_start + winner_size and winner_start < start + size

    alternatives = [
        entry
        for city_id, entry in by_place.items()
        if city_id != winner.place.city_id
        and PLACE_BY_ID[city_id].region_id == winner.place.region_id
        and entry[1].score > 0
        and not entry[1].approximate
        and not overlaps(entry[2])
    ]
    if not alternatives:
        return winner
    return max(alternatives, key=lambda entry: entry[0])[1]


def _match_words(words: list[str]) -> tuple[LocationMatch | None, set[int]]:
    hints = _region_hints(words)
    best: tuple[tuple, LocationMatch] | None = None
    # Har bir hudud bo'yicha eng yaxshi moslik: (rank, match, matn ichidagi o'rni)
    by_place: dict[int, tuple[tuple, LocationMatch, tuple[int, int]]] = {}
    unmatched: list[tuple[int, int, str]] = []

    for start, size, key in _windows(words):
        if len(key) < 3 or key.isdigit():
            continue
        following = words[start + size] if start + size < len(words) else ""
        if following in REGION_MARKERS and key in _region_index:
            continue

        candidates: list[tuple[int, int]] = []
        slot = _name_index.get(key)
        if slot:
            candidates.extend(slot.items())
        else:
            place_id = _phonetic_index.get(phonetic_key(key))
            stem_slot = _name_index.get(_ru_adjective_stem_key(key)) if not place_id else None
            if place_id:
                candidates.append((place_id, PHONETIC_WEIGHT))
            elif stem_slot:
                candidates.extend((stem_place_id, BASE_WEIGHT) for stem_place_id in stem_slot)
            elif len(key) >= 4:
                unmatched.append((start, size, key))

        for place_id, weight in candidates:
            place = PLACE_BY_ID[place_id]
            score = _window_score(words, start, size, weight, place, hints)
            match = LocationMatch(
                server=place.server,
                place=place,
                region=place.region,
                score=score,
                approximate=weight <= FUZZY_WEIGHT,
            )
            rank = _rank_key(score, place, match.approximate)
            if best is None or rank > best[0]:
                best = (rank, match)
            known = by_place.get(place_id)
            if known is None or rank > known[0]:
                by_place[place_id] = (rank, match, (start, size))

    if best is not None and best[1].score > 0:
        return _prefer_district_over_center(best[1], by_place), hints

    for start, size, key in unmatched:
        near = get_close_matches(phonetic_key(key), _fuzzy_keys, n=1, cutoff=FUZZY_CUTOFF)
        if not near:
            continue
        place_id = _phonetic_index.get(near[0])
        if not place_id:
            continue
        place = PLACE_BY_ID[place_id]
        score = _window_score(words, start, size, FUZZY_WEIGHT, place, hints)
        if score <= 0:
            # Ko'cha/uy nomiga o'xshagan taxminiy moslikni hudud sifatida olmaymiz.
            continue
        match = LocationMatch(
            server=place.server,
            place=place,
            region=place.region,
            score=score,
            approximate=True,
            note=f"'{key}' nomi taxminan '{place.server}' deb olindi, tekshirish kerak",
        )
        rank = _rank_key(score, place, True)
        if best is None or rank > best[0]:
            best = (rank, match)

    return (best[1] if best else None), hints


def resolve_location(*texts: str) -> LocationMatch:
    """Berilgan matnlardan hududni aniqlaydi (birinchi ishonchli moslik yutadi)."""

    fallback: LocationMatch | None = None
    hints: set[int] = set()

    for text in texts:
        words = location_words(text or "")
        if not words:
            continue
        match, text_hints = _match_words(words)
        hints |= text_hints
        if match is None:
            continue
        if match.score > 0 and not match.approximate:
            return match
        if fallback is None or _rank_key(match.score, match.place, match.approximate) > _rank_key(
            fallback.score, fallback.place, fallback.approximate
        ):
            fallback = match

    if fallback is not None:
        note = fallback.note
        if not note and fallback.score <= 0:
            note = f"{fallback.server} manzil ichidan aniq ajratilmadi, tekshirish kerak"
        return LocationMatch(
            server=fallback.server,
            place=fallback.place,
            region=fallback.region,
            note=note,
            approximate=True,
            score=fallback.score,
        )

    if hints:
        region = REGION_BY_ID[sorted(hints)[0]]
        center = PLACE_BY_ID[region.center_city_id]
        return LocationMatch(
            server=center.server,
            place=center,
            region=region,
            note=f"faqat {region.uz} aniqlandi, {center.server} markazi qo'yildi",
            approximate=True,
            score=0,
        )

    return LocationMatch(note="P ustun uchun справочникdagi shahar/tuman topilmadi")


def resolve_server_location(*texts: str) -> str:
    return resolve_location(*texts).server


def validate() -> list[str]:
    """Ma'lumot jadvallarini tekshiradi (testlar uchun)."""

    problems: list[str] = []
    for place in PLACES:
        if not place.server:
            problems.append(f"{place.uz}: server nomi bo'sh")
        if place.region_id not in REGION_BY_ID:
            problems.append(f"{place.uz}: viloyat {place.region_id} topilmadi")
    for region in REGION_BY_ID.values():
        if region.center_city_id not in PLACE_BY_ID:
            problems.append(f"{region.uz}: markaz shahri topilmadi")
    duplicates = [server for server in {place.server for place in PLACES} if sum(1 for p in PLACES if p.server == server) > 1]
    for server in duplicates:
        problems.append(f"'{server}' bir nechta tumanga tegishli")
    return problems
