#!/usr/bin/env python3
"""Rebuild the AK Furs assignment outputs from the supplied raw data."""

from __future__ import annotations

import argparse
import csv
import json
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
EXCEL_EPOCH = datetime(1899, 12, 30)
ZERO = Decimal("0")

MONTHS = {
    "янв": 1, "января": 1, "фев": 2, "февраля": 2, "мар": 3, "марта": 3,
    "апр": 4, "апреля": 4, "май": 5, "мая": 5, "июн": 6, "июня": 6,
    "июл": 7, "июля": 7, "авг": 8, "августа": 8, "сен": 9, "сент": 9,
    "сентября": 9, "окт": 10, "октября": 10, "ноя": 11, "ноября": 11,
    "дек": 12, "декабря": 12,
}

NUMBER_WORDS = {
    "ноль": 0, "один": 1, "одна": 1, "одно": 1, "одну": 1, "два": 2,
    "две": 2, "три": 3, "четыре": 4, "пять": 5, "шесть": 6, "семь": 7,
    "восемь": 8, "девять": 9, "десять": 10, "одиннадцать": 11,
    "двенадцать": 12, "тринадцать": 13, "четырнадцать": 14,
    "пятнадцать": 15, "шестнадцать": 16, "семнадцать": 17,
    "восемнадцать": 18, "девятнадцать": 19, "двадцать": 20,
    "тридцать": 30, "сорок": 40, "пятьдесят": 50, "шестьдесят": 60,
    "семьдесят": 70, "восемьдесят": 80, "девяносто": 90, "сто": 100,
    "двести": 200, "триста": 300, "четыреста": 400, "пятьсот": 500,
    "шестьсот": 600, "семьсот": 700, "восемьсот": 800, "девятьсот": 900,
}
SCALE_WORDS = {
    "тысяча": 1000, "тысячи": 1000, "тысяч": 1000,
    "миллион": 1_000_000, "миллиона": 1_000_000,
    "миллионов": 1_000_000,
}


def excel_date(value: Any) -> date | None:
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, (int, float, Decimal)) or re.fullmatch(r"\d+(?:\.\d+)?", str(value).strip()):
        serial = Decimal(str(value).strip())
        if serial < 1 or serial > 100000:
            return None
        return (EXCEL_EPOCH + timedelta(days=float(serial))).date()
    text = str(value).strip().lower().replace(",", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    match = re.fullmatch(r"(\d{1,2})\s+([а-яё]+)\s+(\d{4})", text)
    if match:
        month = MONTHS.get(match.group(2))
        if month:
            return date(int(match.group(3)), month, int(match.group(1)))
    return None


def excel_datetime(value: Any) -> datetime | None:
    if value is None or str(value).strip() == "":
        return None
    text = str(value).strip()
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        serial = Decimal(text)
        if serial < 1 or serial > 100000:
            return None
        return EXCEL_EPOCH + timedelta(days=float(serial))
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    parsed = excel_date(text)
    return datetime.combine(parsed, datetime.min.time()) if parsed else None


def parse_decimal(value: Any, *, field: str = "numeric value") -> Decimal:
    if value is None:
        raise ValueError(f"Missing {field}")
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        raise ValueError(f"Missing {field}")
    text = re.sub(r"(?i)(руб(?:\.|лей|ля|ль)?|rur|rub|₽)", "", text)
    text = text.replace("−", "-")
    multiplier = Decimal(1)
    suffix = re.search(r"(?i)(млн|миллион[аов]?|тыс\.?|тысяч[а-я]*)$", text)
    if suffix:
        word = suffix.group(1).lower()
        multiplier = Decimal(1_000_000 if word.startswith(("млн", "миллион")) else 1000)
        text = text[:suffix.start()]
    text = re.sub(r"[^\d,.\-+]", "", text)
    if not text or text in {"+", "-", ".", ","}:
        raise ValueError(f"Invalid {field}: {value!r}")
    comma, dot = text.rfind(","), text.rfind(".")
    sep = max(comma, dot)
    if comma >= 0 and dot >= 0:
        decimal_sep = text[sep] if len(text) - sep - 1 in (1, 2) else None
    elif sep >= 0:
        trailing = len(text) - sep - 1
        decimal_sep = text[sep] if trailing in (1, 2) else None
    else:
        decimal_sep = None
    if decimal_sep:
        integer_part, fraction = text.rsplit(decimal_sep, 1)
        integer_part = re.sub(r"[,.]", "", integer_part)
        normalized = integer_part + "." + fraction
    else:
        normalized = re.sub(r"[,.]", "", text)
    try:
        return Decimal(normalized) * multiplier
    except InvalidOperation as exc:
        raise ValueError(f"Cannot parse {field}: {value!r}") from exc


def order_id(value: Any) -> int | None:
    match = re.search(r"\d+", str(value or ""))
    return int(match.group()) if match else None


def status_name(value: str) -> str:
    key = value.strip().casefold().replace("ё", "е")
    if key in {"оплачен", "оплачено", "paid"}:
        return "Оплачен"
    if key in {"отменен", "cancelled", "canceled"}:
        return "Отменён"
    if key in {"возврат", "refund"}:
        return "Возврат"
    return "В работе"


def column_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref)
    if not letters:
        raise ValueError(f"Invalid Excel cell reference: {cell_ref}")
    number = 0
    for char in letters.group():
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def read_xlsx_sheets(path: Path) -> dict[str, list[list[str]]]:
    main = f"{{{NS_MAIN}}}"
    rel = f"{{{NS_REL}}}"
    pkg_rel = f"{{{NS_PKG_REL}}}"
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"]
                   for item in relationships.findall(f"{pkg_rel}Relationship")}
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(node.text or "" for node in item.iter(f"{main}t"))
                      for item in strings.findall(f"{main}si")]
        result: dict[str, list[list[str]]] = {}
        for sheet in workbook.findall(f".//{main}sheet"):
            target = targets[sheet.attrib[f"{rel}id"]]
            sheet_path = posixpath.normpath(posixpath.join("xl", target))
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            root = ET.fromstring(archive.read(sheet_path))
            rows: list[list[str]] = []
            for row in root.findall(f".//{main}sheetData/{main}row"):
                values: dict[int, str] = {}
                for cell in row.findall(f"{main}c"):
                    index = column_number(cell.attrib.get("r", "A1"))
                    value = cell.find(f"{main}v")
                    if cell.attrib.get("t") == "s" and value is not None:
                        text = shared[int(value.text or "0")]
                    elif cell.attrib.get("t") == "inlineStr":
                        text = "".join(node.text or "" for node in cell.iter(f"{main}t"))
                    else:
                        text = value.text if value is not None and value.text else ""
                    values[index] = text
                rows.append([values.get(i, "") for i in range(1, max(values, default=0) + 1)])
            result[sheet.attrib["name"]] = rows
    return result


def table_from_sheet(rows: list[list[str]], required_header: set[str]) -> list[dict[str, str]]:
    for index, row in enumerate(rows):
        headers = [item.strip() for item in row]
        if required_header.issubset(set(headers)):
            result = []
            for values in rows[index + 1:]:
                if not any(value.strip() for value in values):
                    continue
                result.append({headers[i]: values[i].strip() if i < len(values) else ""
                               for i in range(len(headers)) if headers[i]})
            return result
    raise ValueError(f"Could not find Excel table header: {sorted(required_header)}")


def read_csv_dicts(path: Path, encoding: str) -> list[dict[str, str]]:
    with path.open("r", encoding=encoding, newline="") as source:
        rows = list(csv.reader(source, delimiter=";"))
    header_index = next((i for i, row in enumerate(rows)
                         if {"Дата операции", "Сумма", "Тип", "Назначение платежа"}.issubset(row)),
                        None)
    if header_index is None:
        if not rows:
            return []
        return list(csv.DictReader(path.open("r", encoding=encoding, newline=""), delimiter=";"))
    header = rows[header_index]
    return [dict(zip(header, row)) for row in rows[header_index + 1:] if any(cell.strip() for cell in row)]


def read_crm(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, delimiter=";")
        rows = [dict(row) for row in reader]
    required = {"Номер заказа", "№ строки", "Дата создания", "Менеджер", "Артикул",
                "Кол-во", "Цена за ед.", "Скидка", "Статус", "Выгружено"}
    if not rows or not required.issubset(rows[0]):
        raise ValueError(f"CRM input is missing columns: {sorted(required - set(rows[0] if rows else []))}")
    return rows


def clean_crm(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    latest: dict[tuple[int, str], dict[str, Any]] = {}
    for row_index, row in enumerate(rows):
        number = order_id(row.get("Номер заказа"))
        if number is None:
            raise ValueError(f"CRM row {row_index + 2} has no order number")
        line = row.get("№ строки", "").strip()
        if not line:
            raise ValueError(f"CRM row {row_index + 2} has no line number")
        created = excel_date(row.get("Дата создания"))
        exported = excel_datetime(row.get("Выгружено"))
        if created is None or exported is None:
            raise ValueError(f"CRM row {row_index + 2} has an invalid date")
        quantity = parse_decimal(row.get("Кол-во"))
        if quantity <= ZERO:
            raise ValueError(f"CRM row {row_index + 2} has an invalid quantity")
        item = dict(row)
        item.update({
            "_order_id": number,
            "_line": line,
            "_created": created,
            "_exported": exported,
            "_quantity": quantity,
            "_unit_price": parse_decimal(row.get("Цена за ед.")),
            "_status": status_name(row.get("Статус", "")),
            "_row_index": row_index,
        })
        key = (number, line)
        previous = latest.get(key)
        if previous is None or (exported, row_index) >= (previous["_exported"], previous["_row_index"]):
            latest[key] = item
    return list(latest.values())


def manager_key(value: str) -> str:
    return re.sub(r"[\W_]+", "", value.casefold(), flags=re.UNICODE)


def manager_names(rows: list[dict[str, str]]) -> dict[str, str]:
    alias_to_canonical: dict[str, str] = {}
    for row in rows:
        values = list(row.values())
        nonempty = [value.strip() for value in values if value.strip()]
        if len(nonempty) < 2:
            continue
        canonical = nonempty[0]
        aliases = nonempty[1:]
        alias_to_canonical[manager_key(canonical)] = canonical
        for alias in aliases:
            for name in re.split(r"[;,|/]", alias):
                if name.strip():
                    alias_to_canonical[manager_key(name)] = canonical
    return alias_to_canonical


def extract_payment_order_ids(purpose: str, existing_orders: set[int]) -> set[int]:
    candidates: set[int] = set()
    for match in re.finditer(r"(?i)(?<![а-яёa-z])(?:a|а)\s*[-‐‑–—]?\s*0*(\d{1,4})(?!\d)", purpose):
        candidates.add(int(match.group(1)))
    contextual = re.compile(
        r"(?i)(?:заказ(?:а|у|ом|е)?|зак\.?)"
        r"\s*(?:№|#|N[º°.]?)?\s*[-–—]?\s*0*(\d{1,4})(?!\d)"
    )
    candidates.update(int(match.group(1)) for match in contextual.finditer(purpose))
    return candidates & existing_orders


def amount_after_discount(row: dict[str, Any]) -> tuple[Decimal, Decimal]:
    gross = row["_unit_price"] * row["_quantity"]
    discount_text = row.get("Скидка", "").strip()
    if not discount_text:
        return row["_unit_price"], gross
    if "%" in discount_text:
        percent = parse_decimal(discount_text.replace("%", ""), field="discount percentage")
        if percent < ZERO or percent > Decimal(100):
            raise ValueError(f"Discount percentage out of range: {discount_text!r}")
        discount_amount = gross * percent / Decimal(100)
    else:
        discount_amount = parse_decimal(discount_text, field="line discount")
    if discount_amount < ZERO or discount_amount > gross:
        raise ValueError(f"Discount exceeds line amount: {discount_text!r}")
    net = gross - discount_amount
    return net / row["_quantity"], net


def workbook_price_data(sheets: dict[str, list[list[str]]]) -> tuple[
        dict[str, list[dict[str, Any]]], dict[str, tuple[str, date]], dict[str, str]]:
    price_rows = table_from_sheet(
        sheets["Прайс_история"],
        {"Артикул", "Действует с", "Действует по", "Минимально допустимая цена, ₽"},
    )
    prices: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in price_rows:
        start = excel_date(row["Действует с"])
        end = excel_date(row["Действует по"])
        if start is None:
            raise ValueError(f"Invalid price-history start date for {row['Артикул']}")
        prices[row["Артикул"].strip()].append({
            "from": start,
            "to": end,
            "minimum": parse_decimal(row["Минимально допустимая цена, ₽"],
                                     field=f"minimum price for {row['Артикул']}"),
            "recommended": parse_decimal(row["Рекомендованная цена, ₽"],
                                         field=f"recommended price for {row['Артикул']}"),
        })
        if end is not None and end < start:
            raise ValueError(f"Price period ends before it starts for {row['Артикул']}")
        if prices[row["Артикул"].strip()][-1]["minimum"] <= ZERO:
            raise ValueError(f"Minimum price must be positive for {row['Артикул']}")
    for article, periods in prices.items():
        periods.sort(key=lambda period: period["from"])
        for previous, current in zip(periods, periods[1:]):
            if previous["to"] is None or previous["to"] >= current["from"]:
                raise ValueError(f"Overlapping price periods for {article}")
    rename_rows = table_from_sheet(
        sheets["Переименования"], {"Старый артикул", "Новый артикул", "Дата перехода"},
    )
    renames: dict[str, tuple[str, date]] = {}
    for row in rename_rows:
        old, new = row["Старый артикул"].strip(), row["Новый артикул"].strip()
        transition = excel_date(row["Дата перехода"])
        if transition is None:
            raise ValueError(f"Invalid article transition date: {old} -> {new}")
        if old in renames:
            raise ValueError(f"Duplicate article rename for {old}")
        renames[old] = (new, transition)
        if not any(period["from"] <= transition and
                   (period["to"] is None or transition <= period["to"])
                   for period in prices.get(new, [])):
            raise ValueError(f"No historical price for {new} on transition date {transition}")
    manager_rows = table_from_sheet(sheets["Менеджеры"], {"Каноническое имя"})
    manager_aliases = manager_names(manager_rows)
    return prices, renames, manager_aliases


def find_minimum_price(article: str, when: date, prices: dict[str, list[dict[str, Any]]],
                       renames: dict[str, tuple[str, date]]) -> Decimal | None:
    rename = renames.get(article)
    canonical = rename[0] if rename else article
    matches = [period for period in prices.get(canonical, [])
               if period["from"] <= when and (period["to"] is None or when <= period["to"])]
    if len(matches) > 1:
        raise ValueError(f"Ambiguous historical price for {article} on {when}")
    return matches[0]["minimum"] if matches else None


def parse_russian_number_words(text: str) -> tuple[Decimal | None, bool]:
    tokens = re.findall(r"[а-яё]+", text.casefold())
    total = 0
    current = 0
    saw_number = False
    saw_scale = False
    for token in tokens:
        if token in NUMBER_WORDS:
            current += NUMBER_WORDS[token]
            saw_number = True
        elif token in SCALE_WORDS:
            scale = SCALE_WORDS[token]
            total += max(current, 1) * scale
            current = 0
            saw_scale = True
        elif token in {"рубль", "рубля", "рублей", "руб"}:
            continue
        else:
            if saw_number:
                break
    if not saw_number:
        return None, saw_scale
    return Decimal(total + current), saw_scale


NUMERIC_AMOUNT_RE = re.compile(
    r"(?<![\w])(?:\d{1,3}(?:[ \u00a0.,]\d{3})+(?:[,.]\d{1,2})?|\d+(?:[,.]\d+)?)"
    r"\s*(?:млн|миллион[аов]?|тыс\.?|тысяч[а-я]*)?"
    r"\s*(?:₽|руб(?:\.|л(?:ь|я|ей)?)?|р\.)?",
    re.IGNORECASE,
)
PRICE_CUE_RE = re.compile(r"цена|стоимост|ценник|обойд[её]тся|цену|цене", re.IGNORECASE)
MONEY_MARKER_RE = re.compile(r"₽|руб|(?:млн|миллион|тыс|тысяч)", re.IGNORECASE)


def manager_mentioned_price(text: str) -> bool:
    if PRICE_CUE_RE.search(text):
        for match in NUMERIC_AMOUNT_RE.finditer(text):
            amount_text = match.group().strip()
            if re.search(r"\d", amount_text) and (MONEY_MARKER_RE.search(amount_text) or
                                                  len(re.findall(r"\d", amount_text)) >= 4):
                return True
        words_amount, has_scale = parse_russian_number_words(text)
        return words_amount is not None and (has_scale or re.search(r"руб", text, re.I) is not None)
    for match in NUMERIC_AMOUNT_RE.finditer(text):
        amount_text = match.group().strip()
        if re.search(r"\d", amount_text) and MONEY_MARKER_RE.search(amount_text):
            return True
    words_amount, has_scale = parse_russian_number_words(text)
    return words_amount is not None and has_scale and re.search(r"руб", text, re.I) is not None


INVITATION_RE = re.compile(
    r"приглаш|приезж|приех|подъ[её]х|подъ[её]зж|"
    r"примерк|встреч|запишу|записать|когда вам удобно|когда сможете|заезжайте|заехать|"
    r"жд[её]м вас|будем рады видеть|буду рада видеть|"
    r"приходите|встречаемся|визит в ателье",
    re.IGNORECASE,
)


def analyze_chat(path: Path) -> dict[str, str | bool]:
    chat_id = path.stem
    first_price: datetime | None = None
    first_invite: datetime | None = None
    price_text = ""
    invite_text = ""
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        match = re.match(r"^\[(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2})\]\s*([^:]+):\s*(.*)$", line)
        if not match:
            continue
        author = match.group(2).strip().casefold()
        if author not in {"менеджер", "клиент"}:
            continue
        timestamp = datetime.strptime(match.group(1), "%d.%m.%Y %H:%M")
        message = match.group(3)
        if author == "менеджер":
            if first_price is None and manager_mentioned_price(message):
                first_price = timestamp
                price_text = message
            if first_invite is None and INVITATION_RE.search(message):
                first_invite = timestamp
                invite_text = message
    violation = first_invite is not None and (first_price is None or first_invite < first_price)
    return {
        "chat_id": chat_id,
        "violation": violation,
        "first_invitation_at": first_invite.isoformat(sep=" ") if first_invite else "",
        "first_invitation_text": invite_text,
        "first_price_at": first_price.isoformat(sep=" ") if first_price else "",
        "first_price_text": price_text,
    }


def chat_first_timestamp(path: Path) -> datetime:
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        match = re.match(r"^\[(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2})\]", line)
        if match:
            return datetime.strptime(match.group(1), "%d.%m.%Y %H:%M")
    raise ValueError(f"Could not determine the date of chat {path.stem}")


def q5_client_message(prices: dict[str, list[dict[str, Any]]], chats_dir: Path) -> str:
    chat_path = chats_dir / "C0010.txt"
    chat_timestamp = chat_first_timestamp(chat_path)
    chat_date = chat_timestamp.date()
    period = next(
        (item for item in prices["SOB-ZH-01"]
         if item["from"] <= chat_date and (item["to"] is None or chat_date <= item["to"])),
        None,
    )
    if period is None:
        raise ValueError(f"No sable-vest price available on {chat_date}")
    low = period["minimum"]
    high = period["recommended"]
    return (
        "Марина: Понимаю Вас — ехать через весь город, не зная бюджета, действительно "
        "неудобно. Сразу сориентирую: жилеты из натурального соболя стоят "
        f"примерно от {format_rub(low)} до {format_rub(high)} ₽; точная стоимость зависит "
        f"от модели, длины и выбранной отделки. Это ориентир по прайсу на "
        f"{chat_date.strftime('%d.%m.%Y')}."
        " Если этот диапазон Вам подходит, с удовольствием заранее "
        "подготовлю подходящие варианты и согласую удобное время. За один визит Вы "
        "сможете примерить несколько моделей, сравнить мех и посадку, а также обсудить "
        "индивидуальные мерки — без обязательства принимать решение сразу."
    )


def format_rub(amount: Decimal | int | float) -> str:
    number = Decimal(amount).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return f"{number:,.0f}".replace(",", " ")


def json_number(value: Decimal) -> int | float:
    rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(rounded) if rounded == rounded.to_integral_value() else float(rounded)


def write_outputs(data_dir: Path, output_dir: Path, template_path: Path) -> dict[str, Any]:
    crm = read_crm(data_dir / "01_crm_orders_export.csv")
    clean_rows = clean_crm(crm)
    sheets = read_xlsx_sheets(data_dir / "03_catalog_prices_reference.xlsx")
    prices, renames, manager_aliases = workbook_price_data(sheets)

    order_ids = {row["_order_id"] for row in clean_rows}
    incoming: list[tuple[dict[str, str], set[int], Decimal]] = []
    paid_orders: set[int] = set()
    for payment in read_csv_dicts(data_dir / "02_bank_statement_rub.csv", "cp1251"):
        if payment.get("Тип", "").strip().casefold() != "поступление":
            continue
        amount = parse_decimal(payment.get("Сумма", ""), field="bank payment amount")
        if amount < ZERO:
            raise ValueError(f"Incoming payment cannot be negative: {payment!r}")
        linked = extract_payment_order_ids(payment.get("Назначение платежа", ""), order_ids)
        incoming.append((payment, linked, amount))
        paid_orders.update(linked)

    paid_crm_orders = {row["_order_id"] for row in clean_rows if row["_status"] == "Оплачен"}
    unidentified = [item for item in incoming if not item[1]]
    unidentified_total = sum((item[2] for item in unidentified), ZERO)

    violations: list[tuple[str, Decimal]] = []
    missing_price: Counter[str] = Counter()
    unknown_managers: Counter[str] = Counter()
    violations_by_manager: Counter[str] = Counter()
    violation_shortfall = ZERO
    for row in clean_rows:
        if row["_status"] == "Отменён":
            continue
        article = row["Артикул"].strip()
        minimum = find_minimum_price(article, row["_created"], prices, renames)
        if minimum is None:
            missing_price[article] += 1
            continue
        after_discount, _ = amount_after_discount(row)
        if after_discount < minimum - Decimal("0.0001"):
            manager_raw = row.get("Менеджер", "").strip()
            manager = manager_aliases.get(manager_key(manager_raw))
            if manager is None:
                unknown_managers[manager_raw] += 1
                manager = manager_raw or "Не указан"
            violations_by_manager[manager] += 1
            violation_shortfall += (minimum - after_discount) * row["_quantity"]
            violations.append((manager, minimum - after_discount))

    if missing_price:
        raise ValueError(f"Missing historical prices for CRM rows: {dict(missing_price)}")
    if unknown_managers:
        raise ValueError(f"Unmapped manager aliases in price violations: {dict(unknown_managers)}")

    chats_dir = data_dir / "chats"
    chat_files = sorted(chats_dir.glob("C*.txt"))
    if len(chat_files) != 150:
        raise ValueError(f"Expected 150 chat files in {chats_dir}, found {len(chat_files)}")
    chat_audit = [analyze_chat(path) for path in chat_files]
    q4_ids = [str(item["chat_id"]) for item in chat_audit if item["violation"]]

    template = json.loads(template_path.read_text(encoding="utf-8-sig"))
    required_keys = [
        "q1_unique_orders", "q1_unique_order_lines", "q2a_paid_orders_without_payment",
        "q2b_unidentified_payments_count", "q2b_unidentified_payments_sum_rub",
        "q3_price_violation_lines_total", "q3_price_violation_lines_by_manager",
        "q4_chats_meeting_before_price", "run_command", "notes",
    ]
    if list(template) != required_keys:
        raise ValueError("answers_template.json keys or key order have changed")
    result = dict(template)
    result.update({
        "q1_unique_orders": len({row["_order_id"] for row in clean_rows}),
        "q1_unique_order_lines": len(clean_rows),
        "q2a_paid_orders_without_payment": sorted(paid_crm_orders - paid_orders),
        "q2b_unidentified_payments_count": len(unidentified),
        "q2b_unidentified_payments_sum_rub": json_number(unidentified_total),
        "q3_price_violation_lines_total": sum(violations_by_manager.values()),
        "q3_price_violation_lines_by_manager": dict(sorted(violations_by_manager.items())),
        "q4_chats_meeting_before_price": q4_ids,
        "run_command": "python src\\solve.py",
        "notes": (
            "Дубли строк удалены по паре «номер заказа + № строки» с выбором наиболее поздней "
            "выгрузки; серийные Excel-даты преобразованы в календарные. Платёж считается "
            "связанным только при наличии в назначении явного номера заказа или кода A-...; "
            "номер счёта, дата и сумма сами по себе не считаются номером заказа. Старые артикулы "
            "сопоставлены с новыми по справочнику; дата перехода проверяется на наличие цены "
            "для нового артикула. Цены выбраны по включительным границам интервала. В Q4 "
            "учитываются только реплики менеджера, а первые найденные события сохраняются в "
            "output/q4_chat_audit.csv для проверки. Для ответа C0010 используется прайс на дату "
            "переписки, без данных о более позднем заказе. Скидка в процентах применяется к "
            "сумме строки, скидка без «%» вычитается из суммы всей строки."
        ),
    })
    if set(result) != set(required_keys):
        raise ValueError("Generated answers do not match the answer template")

    output_dir.mkdir(parents=True, exist_ok=True)
    answers_path = output_dir / "answers.json"
    answers_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    recommendation = max(violations_by_manager.items(), key=lambda item: item[1])[0] \
        if violations_by_manager else "менеджерами"
    memo = (
        "# Записка руководителю\n\n"
        f"1. После удаления повторных выгрузок осталось **{result['q1_unique_orders']} заказов** "
        f"и **{result['q1_unique_order_lines']} строк**.\n"
        f"2. В выписке не удалось связать с заказом **{len(unidentified)} поступлений** "
        f"на **{format_rub(unidentified_total)} ₽**. Такие платежи требуют ручной проверки "
        "по плательщику и дате.\n"
        f"3. Среди заказов, отмеченных в CRM как оплаченные, **{len(result['q2a_paid_orders_without_payment'])}** "
        "не имеют найденного поступления в этой выписке. Также выявлено "
        f"**{result['q3_price_violation_lines_total']} строк** с ценой ниже установленного минимума; "
        f"расчётная разница до минимальной цены — **{format_rub(violation_shortfall)} ₽**.\n\n"
        "**Рекомендации.** (1) Ежедневно сверять поступления и заказы, а неидентифицированные "
        "платежи направлять ответственному на разбор до закрытия смены. (2) Перед подтверждением "
        "скидки проверять итоговую цену каждой позиции по действующему минимуму; отдельно "
        f"разобрать случаи по менеджеру «{recommendation}».\n"
    )
    (output_dir / "management_memo.md").write_text(memo, encoding="utf-8")

    client_message = q5_client_message(prices, chats_dir)
    (output_dir / "client_message_C0010.txt").write_text(client_message + "\n", encoding="utf-8")
    with (output_dir / "q4_chat_audit.csv").open("w", encoding="utf-8-sig", newline="") as audit_file:
        writer = csv.DictWriter(
            audit_file,
            fieldnames=[
                "chat_id", "violation", "first_invitation_at", "first_invitation_text",
                "first_price_at", "first_price_text",
            ],
        )
        writer.writeheader()
        writer.writerows(chat_audit)

    validation_notes = (
        "# Проверка и допущения\n\n"
        "Банковские строки читаются в CP1251; строки с назначением вида «счёт №...», "
        "«по счёту №...» не связываются с заказом без отдельного номера заказа. Для Q4 "
        "в output/q4_chat_audit.csv записаны первые найденные менеджерские приглашение и "
        "озвучивание цены с временными метками и исходным текстом. Проверяйте эти реплики "
        "для пограничных формулировок: эвристический поиск по тексту не заменяет ручную "
        "проверку. Ответ для C0010 использует исторический диапазон прайса на дату чата "
        f"({chat_first_timestamp(chats_dir / 'C0010.txt').strftime('%d.%m.%Y')}); "
        "CRM-заказ, созданный позднее переписки, "
        "в сообщение не включается. Пустые и некорректные обязательные числовые поля приводят "
        "к ошибке вместо подстановки нуля. Исходные файлы не изменяются."
    )
    (output_dir / "ai_validation_notes.md").write_text(validation_notes + "\n", encoding="utf-8")

    if unknown_managers:
        print(f"Warning: unmapped manager aliases: {dict(unknown_managers)}", file=sys.stderr)
    if missing_price:
        print(f"Warning: no historical price match for SKUs: {dict(missing_price)}", file=sys.stderr)
    print(f"Created {answers_path}", file=sys.stderr)
    print(f"Created {output_dir / 'management_memo.md'}", file=sys.stderr)
    print(f"Created {output_dir / 'client_message_C0010.txt'}", file=sys.stderr)
    print(f"Created {output_dir / 'ai_validation_notes.md'}", file=sys.stderr)
    print(f"Created {output_dir / 'q4_chat_audit.csv'}", file=sys.stderr)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "data_dir", nargs="?", type=Path,
        help="directory containing the source files (default: ./Initial project data)",
    )
    parser.add_argument("--output", type=Path, help="directory for generated files")
    args = parser.parse_args()
    script_dir = Path(__file__).resolve().parent
    root = script_dir if (script_dir / "Initial project data").is_dir() else script_dir.parent
    data_dir = (args.data_dir or root / "Initial project data").resolve()
    output_dir = (args.output or root / "output").resolve()
    template_path = data_dir / "answers_template.json"
    try:
        result = write_outputs(data_dir, output_dir, template_path)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
