import tempfile
import unittest
from pathlib import Path
from itertools import product
from subprocess import run
import sys

from openpyxl import Workbook

from purchase import price, quantity, read_items, report


class PurchaseTests(unittest.TestCase):
    def test_cli_input_formats(self):
        root = Path(__file__).resolve().parents[1]
        for request_ext, prices_ext in product(("csv", "xlsx"), repeat=2):
            with self.subTest(request=request_ext, prices=prices_ext):
                result = run(
                    [sys.executable, "-X", "utf8", "purchase.py",
                     f"examples/request.{request_ext}", f"examples/prices.{prices_ext}"],
                    cwd=root, capture_output=True, encoding="utf-8", check=True,
                )
                self.assertIn("LABEL-03: 2 шт. — нет в прайсе", result.stdout)
                self.assertTrue(result.stdout.rstrip().endswith("742,00 ₽"))

    def test_excel_codes_prices_and_error_location(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "прайс.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.append(["артикул", "цена"])
            sheet.append(["001", 10.05])
            sheet.append(["1", "0,01"])
            book.save(path)
            self.assertEqual(read_items(path, "цена", price),
                             {"001": price("10,05"), "1": price("0,01")})
            sheet.append(["002", -1])
            book.save(path)
            with self.assertRaises(ValueError) as error:
                read_items(path, "цена", price)
            self.assertIn(f"{path}, строка 4", str(error.exception))

    def test_partial_quote(self):
        result = report({"001": 4, "1": 2}, {"001": price("125,50")})
        self.assertIn("1: 2 шт. — нет в прайсе", result)
        self.assertTrue(result.endswith("По найденным позициям (без цены: 1): 502,00 ₽"))

    def test_zero_price(self):
        self.assertTrue(report({"A": 2}, {"A": price("0")}).endswith("Итого: 0,00 ₽"))

    def test_duplicate_sku(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "request.csv"
            path.write_text("артикул;количество\n001;1\n001;2\n", encoding="utf-8-sig")
            with self.assertRaisesRegex(ValueError, "повторяющийся"):
                read_items(path, "количество", quantity)

    def test_invalid_numbers(self):
        for parse, value in [(quantity, "0"), (quantity, "1.5"), (price, "-1"), (price, "1.001")]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse(value)
