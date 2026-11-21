import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from subprocess import run

from openpyxl import load_workbook

from allocation import Offer, build_quote
from excel_report import save_report


class ReportTests(unittest.TestCase):
    def test_export_from_cli(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "закупка.xlsx"
            run([sys.executable, "-X", "utf8", "purchase.py", "examples/request.csv",
                 "examples/prices.csv", "examples/Склад.csv", "-o", str(path)],
                cwd=Path(__file__).resolve().parents[1], check=True, capture_output=True)
            book = load_workbook(path, data_only=True)
            self.assertEqual(set(book.sheetnames), {"Сводка", "Требуют внимания", "Склад", "prices"})
            summary = book["Сводка"]
            self.assertIn("дефицитом: 1", summary["C2"].value)
            total = Decimal(str(summary["B2"].value))
            self.assertEqual(total, Decimal("652.10"))
            orders_total = Decimal(0)
            for name in ("Склад", "prices"):
                sheet = book[name]
                rows = list(sheet.iter_rows(min_row=4, values_only=True))
                amount = sum(Decimal(str(row[5])) * row[4] for row in rows)
                self.assertEqual(amount, sum(Decimal(str(row[6])) for row in rows))
                self.assertEqual(amount, Decimal(str(sheet["B2"].value)))
                orders_total += amount
            self.assertEqual(orders_total, total)
            self.assertEqual(total, sum(Decimal(str(row[5])) for row in
                                        summary.iter_rows(min_row=4, values_only=True)))
            self.assertEqual(book["Требуют внимания"]["A4"].value, "LABEL-03")
            book.close()

    def test_unlisted_surplus_and_supplier_names(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "результат.xlsx"
            quote = build_quote({"001": 1, "=B2": 1, "C": 1}, [
                {"001": Offer("Сводка", Decimal(1), 2)},
                {"=B2": Offer("Длинное имя" * 4, Decimal(0))},
            ])
            save_report(path, quote)
            book = load_workbook(path)
            self.assertEqual(book["Сводка (2)"]["A4"].value, "001")
            self.assertEqual(book["Сводка"]["A5"].data_type, "s")
            self.assertEqual(book["Требуют внимания"]["D4"].value, 1)
            self.assertEqual(book["Требуют внимания"]["G5"].value, "Нет в прайсах")
            self.assertLessEqual(max(map(len, book.sheetnames)), 31)
            book.close()
