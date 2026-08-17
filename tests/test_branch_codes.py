"""Filial kodi tanlash testlari (branch_codes.py)."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import branch_codes as B  # noqa: E402
import locations as L  # noqa: E402


class BranchFileTests(unittest.TestCase):
    def test_file_is_loaded(self):
        records = B.load_branch_code_records()
        self.assertGreater(len(records), 200)
        self.assertTrue(all(record.code and record.name for record in records))

    def test_codes_are_unique_per_office(self):
        records = B.load_branch_code_records()
        names = [L.normalize_location_key(record.name) for record in records]
        self.assertEqual(len(names), len(set(names)), "branch_codes.xlsx da nomlar takrorlangan")


class BranchCodeTests(unittest.TestCase):
    def code(self, location: str, address: str = "") -> str:
        code, _ = B.branch_code_for_address(location, address)
        return code

    def test_every_district_gets_a_code(self):
        for place in L.PLACES:
            with self.subTest(place=place.server):
                code, note = B.branch_code_for_address(place.server, place.uz)
                self.assertTrue(code, f"{place.server} ({place.uz}) uchun kod topilmadi: {note}")

    def test_districts_with_own_office_have_no_note(self):
        for place in L.PLACES:
            if not L.offices_for_server(place.server):
                continue
            with self.subTest(place=place.server):
                _, note = B.branch_code_for_address(place.server, place.uz)
                self.assertEqual("", note)

    def test_known_office_codes(self):
        cases = {
            "Денау": "23",          # Denov O'zbegim
            "Карлук": "71",         # Oltinsoy
            "Халкабад": "337",      # Muzrabot
            "Янгикишлак": "356",    # Forish
            "Навбахор": "334",      # Furqat
            "Зиадин": "365",        # Ziyovuddin
            "Пайарык": "304",       # Chelak
            "Келес": "322",         # Keles
            "Нурафшон": "69",       # Nurafshon
            "Баяут": "106",         # Boyovut
            "Тупраккала": "362",    # Pitnak
            "Элликкала": "269",     # Ellikqal'a
            "Ханабад": "363",       # Xonobod
            "Улугнор": "366",       # Oqoltin (Andijon)
            "Хива": "45",
            "Термез": "10",         # Termiz Kurant
        }
        for location, expected in cases.items():
            with self.subTest(location=location):
                self.assertEqual(expected, self.code(location, location))

    def test_address_picks_the_right_office_in_a_city(self):
        # Toshkent shahri: manzil bo'yicha tuman ofisi tanlanadi.
        self.assertEqual("21", self.code("Чиланзар", "массив Чиланзор, 5-й квартал, 29"))
        self.assertEqual("78", self.code("Мирзо-Улугбек", "улица Мухаммада Юсуфа, 1"))
        self.assertEqual("36", self.code("Сергели", "массив Сергели-VIIIА, 17"))
        # Andijon shahrida ikki ofis bor - manzil hal qiladi.
        self.assertEqual("3", self.code("Андижан", "проспект Амира Тимура, 11"))
        self.assertEqual("27", self.code("Андижан", "Узбекистанская улица, 55"))

    def test_district_without_office_falls_back_to_region_center(self):
        # Ishtixon (Samarqand) da EMU ofisi yo'q.
        code, note = B.branch_code_for_address("Иштыхан", "Ishtixon tumani")
        self.assertTrue(code)
        self.assertIn("Самарканд", note)
        # Xovos (Sirdaryo) da ofis yo'q.
        code, note = B.branch_code_for_address("Хаваст", "Xovos tumani")
        self.assertTrue(code)
        self.assertIn("Гулистан", note)

    def test_unknown_location_returns_note(self):
        code, note = B.branch_code_for_address("Yo'q shahar", "manzil")
        self.assertEqual("", code)
        self.assertTrue(note)

    def test_empty_location_is_quiet(self):
        self.assertEqual(("", ""), B.branch_code_for_address("", "manzil"))

    def test_new_office_name_is_found_without_regenerated_data(self):
        """branch_codes.xlsx ga tuman nomi bilan ofis qo'shilsa, kod shu ofisdan olinadi."""

        records = list(B.load_branch_code_records())
        records.append(B.BranchCodeRecord(code="999", parent="Guliston Oxunboboyev", name="Xovos", address="test"))
        original = B._records_cache
        try:
            B._records_cache = records
            code, note = B.branch_code_for_address("Хаваст", "Xovos tumani")
            self.assertEqual("999", code)
            self.assertEqual("", note)
        finally:
            B._records_cache = original


if __name__ == "__main__":
    unittest.main(verbosity=2)
