"""Excel qatorlarini tayyorlash testlari (main.prepare_rows).

`main.py` telegram/openai kutubxonalarini import qiladi - test uchun ular
o'rniga bo'sh (mock) modullar qo'yiladi, chunki tekshirilayotgan mantiq
faqat hudud + filial kodiga tegishli.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

for _name in ("aiogram", "aiogram.exceptions", "aiogram.filters", "aiogram.types", "dotenv", "openai"):
    sys.modules.setdefault(_name, MagicMock())

import main  # noqa: E402

OFFICE_SENDER = {
    "client_type": "physical",
    "delivery_type": "ДО ОФИСА",
    "parcel_weight": "1",
    "places_count": "1",
    "payment_by_receiver": "False",
    "sender_full_name": "JO'NATUVCHI",
    "sender_address": "Toshkent",
    "sender_phone": "998901112233",
    "sender_city_ru": "Ташкент",
    "cipher_prefix": "ABC",
}

ADDRESS_COLUMN = 4
LOCATION_COLUMN = 16
REVIEW_COLUMN = -1


def customer(address: str, note: str = "", region: str = "") -> dict[str, str]:
    return {
        "full_name": "MIJOZ",
        "phone": "901112233",
        "address": address,
        "note": note,
        "recipient_region_ru": region,
        "source_cipher": "",
        "needs_review": "",
    }


class PrepareRowsTests(unittest.TestCase):
    def row(self, address: str, note: str = "", region: str = "", sender: dict | None = None) -> list:
        rows = main.prepare_rows([customer(address, note, region)], sender or OFFICE_SENDER)
        self.assertEqual(1, len(rows))
        return rows[0]

    def test_allowed_locations_come_from_emu_data(self):
        self.assertEqual(198, len(main.ALLOWED_RECIPIENT_LOCATIONS))
        self.assertIn("Денау", main.ALLOWED_RECIPIENT_LOCATIONS)

    def test_denov_spellings_give_the_same_row(self):
        for address in (
            "denov tumani, Sharof Rashidov ko'chasi 12",
            "денов",
            "ДЕНОУ",
            "Denau shahri",
            "Surxondaryo viloyati Denov tumani",
        ):
            with self.subTest(address=address):
                row = self.row(address)
                self.assertEqual("Денау", row[LOCATION_COLUMN])
                self.assertEqual("23", row[ADDRESS_COLUMN])
                self.assertEqual("", row[REVIEW_COLUMN])

    def test_office_code_for_various_districts(self):
        cases = {
            "Toshkent, Chilonzor tumani, massiv Chilonzor 5-kvartal 29": ("Чиланзар", "21"),
            "Хоразм вилояти Хива тумани": ("Хива", "45"),
            "Andijon viloyati, Asaka tumani, Umid ko'chasi 90": ("Асака", "102"),
            "Наманганская область, Чустский район": ("Чуст", "99"),
            "Farg'ona viloyati Quva tumani": ("Кува", "235"),
            "Qashqadaryo, Koson tumani": ("Касан", "86"),
        }
        for address, (location, code) in cases.items():
            with self.subTest(address=address):
                row = self.row(address)
                self.assertEqual(location, row[LOCATION_COLUMN])
                self.assertEqual(code, row[ADDRESS_COLUMN])

    def test_district_without_office_is_flagged(self):
        row = self.row("Samarqand viloyati, Ishtixon tumani, Chelak")
        self.assertEqual("Иштыхан", row[LOCATION_COLUMN])
        self.assertTrue(row[ADDRESS_COLUMN])
        self.assertIn("aniq filial topilmadi", row[REVIEW_COLUMN])

    def test_home_delivery_keeps_the_address(self):
        sender = dict(OFFICE_SENDER, delivery_type="НА ДОМ")
        row = self.row("Denov tumani, Sharof Rashidov ko'chasi 12", sender=sender)
        self.assertEqual("Денау", row[LOCATION_COLUMN])
        self.assertIn("Sharof Rashidov", row[ADDRESS_COLUMN])

    def test_region_column_used_when_address_has_no_district(self):
        row = self.row("5-uy, 12-xonadon", region="Денау")
        self.assertEqual("Денау", row[LOCATION_COLUMN])
        self.assertEqual("23", row[ADDRESS_COLUMN])

    def test_unknown_location_is_reported(self):
        row = self.row("qwerty 123")
        self.assertEqual("", row[LOCATION_COLUMN])
        self.assertIn("topilmadi", row[REVIEW_COLUMN])

    def test_legal_client_row_layout(self):
        sender = dict(OFFICE_SENDER, client_type=main.CLIENT_TYPE_LEGAL)
        rows = main.prepare_rows([customer("Денов тумани")], sender)
        self.assertEqual("Денау", rows[0][11])
        self.assertEqual("23", rows[0][ADDRESS_COLUMN])


class ExcelImportTests(unittest.TestCase):
    def test_excel_row_location_is_resolved(self):
        self.assertEqual("Денау", main.resolve_server_location("denov tumani, 5-uy"))
        self.assertEqual("Ташкент", main.resolve_server_location("Toshkent shahri, Amir Temur ko'chasi"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
