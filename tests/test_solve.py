import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import solve


class SolveTests(unittest.TestCase):
    def test_manager_aliases_ignore_employment_dates(self) -> None:
        aliases = solve.manager_names([
            {
                "Каноническое имя": "Марина Котова",
                "Варианты написания в CRM": "Котова М. | Марина | marina_k",
                "Дата выхода": "45810",
                "Дата увольнения": "46112",
            },
        ])
        self.assertEqual(aliases[solve.manager_key("marina_k")], "Марина Котова")
        self.assertNotIn(solve.manager_key("45810"), aliases)
        self.assertNotIn(solve.manager_key("46112"), aliases)

    def test_order_id_accepts_only_complete_supported_formats(self) -> None:
        self.assertEqual(solve.order_id("A-00075"), 75)
        self.assertEqual(solve.order_id("А-75"), 75)
        self.assertEqual(solve.order_id("a75"), 75)
        self.assertEqual(solve.order_id("75"), 75)
        self.assertIsNone(solve.order_id("2025/A-75"))
        self.assertIsNone(solve.order_id("order 75"))

    def test_invoice_number_is_not_order_number(self) -> None:
        existing = {828, 243}
        self.assertEqual(
            solve.extract_payment_order_ids("по счету №828 от 13.12.2025", existing),
            {828},
        )
        self.assertEqual(
            solve.extract_payment_order_ids("Счёт № 00828", existing),
            {828},
        )
        self.assertEqual(
            solve.extract_payment_order_ids("Оплата по счёту 828", existing),
            {828},
        )
        self.assertEqual(
            solve.extract_payment_order_ids("Предоплата по счету №243", existing),
            {243},
        )
        self.assertEqual(
            solve.extract_payment_order_ids("Оплата заказа A-00243", existing),
            {243},
        )

    def test_multiple_order_candidates_are_not_linked(self) -> None:
        existing = {243, 828}
        candidates = solve.extract_payment_order_ids(
            "Оплата заказа A-00243 и заказа 828", existing,
        )
        self.assertEqual(candidates, {243, 828})
        self.assertIsNone(solve.unique_payment_order_id(candidates))
        self.assertEqual(solve.unique_payment_order_id({243}), 243)

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

    def test_chat_events_use_timestamps_then_source_order(self) -> None:
        content = (
            "[01.01.2025 10:20] Менеджер: Стоимость модели — 50 000 ₽.\n"
            "[01.01.2025 10:10] Менеджер: Приглашаю на примерку.\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9998.txt"
            path.write_text(content, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertTrue(event["violation"])

        same_time = (
            "[01.01.2025 10:10] Менеджер: Приглашаю на примерку.\n"
            "[01.01.2025 10:10] Менеджер: Стоимость модели — 50 000 ₽.\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9997.txt"
            path.write_text(same_time, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertFalse(event["violation"])
        self.assertTrue(event["ambiguous"])
        self.assertIn("одинаковую временную метку", event["ambiguity_reason"])
        self.assertLess(event["first_invitation_line"], event["first_price_line"])

        reverse_order = (
            "[01.01.2025 10:10] Менеджер: Стоимость модели — 50 000 ₽.\n"
            "[01.01.2025 10:10] Менеджер: Приглашаю на примерку.\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9994.txt"
            path.write_text(reverse_order, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertFalse(event["violation"])
        self.assertTrue(event["ambiguous"])

    def test_same_line_chat_events_are_ambiguous_not_confirmed(self) -> None:
        content = (
            "[01.01.2025 10:10] Менеджер: Приглашаю на примерку, "
            "стоимость модели — 50 000 ₽.\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9996.txt"
            path.write_text(content, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertFalse(event["violation"])
        self.assertTrue(event["ambiguous"])

    def test_invitation_without_any_price_is_audited_as_before_price(self) -> None:
        content = "[01.01.2025 10:10] Менеджер: Приглашаю на примерку.\n"
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "C9995.txt"
            path.write_text(content, encoding="utf-8")
            event = solve.analyze_chat(path)
        self.assertTrue(event["violation"])
        self.assertFalse(event["ambiguous"])
        self.assertEqual(event["first_price_at"], "")

    def test_c0010_price_is_selected_for_chat_date(self) -> None:
        data_dir = Path(__file__).resolve().parents[1] / "Initial project data"
        sheets = solve.read_xlsx_sheets(data_dir / "03_catalog_prices_reference.xlsx")
        prices, _, _ = solve.workbook_price_data(sheets)
        on_chat_date = solve.find_minimum_price("SOB-ZH-01", date(2025, 9, 4), prices, {})
        latest = max(prices["SOB-ZH-01"], key=lambda item: item["from"])
        self.assertEqual(on_chat_date, Decimal(199000))
        self.assertNotEqual(on_chat_date, latest["minimum"])

    def test_c0010_reply_uses_recommended_price_without_internal_details(self) -> None:
        data_dir = Path(__file__).resolve().parents[1] / "Initial project data"
        sheets = solve.read_xlsx_sheets(data_dir / "03_catalog_prices_reference.xlsx")
        prices, _, _ = solve.workbook_price_data(sheets)
        message = solve.q5_client_message(prices, data_dir / "chats")
        self.assertIn("234 000 ₽", message)
        self.assertNotIn("199 000 ₽", message)
        self.assertNotIn("прайс на", message)
        self.assertNotIn("04.09.2025", message)


if __name__ == "__main__":
    unittest.main()
