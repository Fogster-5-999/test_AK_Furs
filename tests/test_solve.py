import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import solve


class SolveTests(unittest.TestCase):
    def test_invoice_number_is_not_order_number(self) -> None:
        existing = {828, 243}
        self.assertEqual(
            solve.extract_payment_order_ids("по счету №828 от 13.12.2025", existing),
            set(),
        )
        self.assertEqual(
            solve.extract_payment_order_ids("Оплата заказа A-00243", existing),
            {243},
        )

    def test_empty_required_number_is_an_error(self) -> None:
        with self.assertRaises(ValueError):
            solve.parse_decimal("", field="payment amount")
        with self.assertRaises(ValueError):
            solve.parse_decimal("нет данных", field="payment amount")

    def test_discount_is_applied_to_whole_line(self) -> None:
        row = {"_unit_price": Decimal(1000), "_quantity": Decimal(2), "Скидка": "100"}
        unit, total = solve.amount_after_discount(row)
        self.assertEqual(unit, Decimal(950))
        self.assertEqual(total, Decimal(1900))

    def test_chat_uses_manager_events_and_keeps_evidence(self) -> None:
        content = (
            "Чат C9999 | клиент CL-1 | заказ A-1 | менеджер Марина\n"
            "[01.01.2025 10:00] Клиент: Стоит ли это 50 000 ₽?\n"
            "[01.01.2025 10:10] Менеджер: Приглашаю на примерку.\n"
            "[01.01.2025 10:20] Менеджер: Стоимость модели — 50 000 ₽.\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9999.txt"
            path.write_text(content, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertTrue(event["violation"])
        self.assertEqual(event["first_invitation_text"], "Приглашаю на примерку.")
        self.assertEqual(event["first_price_text"], "Стоимость модели — 50 000 ₽.")

    def test_c0010_price_is_selected_for_chat_date(self) -> None:
        data_dir = Path(__file__).resolve().parents[1] / "Initial project data"
        sheets = solve.read_xlsx_sheets(data_dir / "03_catalog_prices_reference.xlsx")
        prices, _, _ = solve.workbook_price_data(sheets)
        on_chat_date = solve.find_minimum_price("SOB-ZH-01", date(2025, 9, 4), prices, {})
        latest = max(prices["SOB-ZH-01"], key=lambda item: item["from"])
        self.assertEqual(on_chat_date, Decimal(199000))
        self.assertNotEqual(on_chat_date, latest["minimum"])


if __name__ == "__main__":
    unittest.main()
