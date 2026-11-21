import tempfile
import unittest
from pathlib import Path

from purchase import price, quantity, read_items, report


class PurchaseTests(unittest.TestCase):
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
