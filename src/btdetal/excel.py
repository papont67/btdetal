"""excel.py — чтение XLS/XLSX, проверка колонок и отбор доступных товаров."""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from io import BytesIO
from typing import Any

import openpyxl
import xlrd

HEADERS = ("код", "артикул", "номенклатура", "остаток", "1-я цена")


class PriceImportError(ValueError):
    """Файл не является корректным прайсом; действующая база не изменяется."""


@dataclass(frozen=True)
class Cell:
    """Значение ячейки с форматом для сохранения ведущих нулей."""

    value: Any
    number_format: str = "General"
    invalid: bool = False


def text(cell: Cell) -> str:
    """Вернуть текст, сохранив строковые и форматированные числовые идентификаторы."""
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        section = cell.number_format.split(";")[0]
        if re.fullmatch(r"0+\.0+", section):
            # В прайсе встречаются артикулы вроде 651017674.5, записанные числом.
            integer_part, decimals = section.split(".")
            return format(Decimal(str(value)), f".{len(decimals)}f").zfill(
                len(integer_part) + len(decimals) + 1
            )
        if not float(value).is_integer():
            return format(Decimal(str(value)), "f")
        if re.fullmatch("0+", section):
            return str(int(value)).zfill(len(section))
        return str(int(value))
    return str(value).strip()


def normalize(value: str) -> str:
    """Нормализовать регистр и пробелы, не удаляя значимые символы артикула."""
    return " ".join(value.casefold().split())


def number(cell: Cell) -> Decimal:
    """Прочитать конечное неотрицательное число с запятой или пробелами."""
    try:
        if cell.invalid or cell.value is None or isinstance(cell.value, bool):
            raise ValueError
        value = Decimal("".join(str(cell.value).split()).replace(",", "."))
        if not value.is_finite() or value < 0:
            raise ValueError
        return value
    except (ValueError, InvalidOperation):
        raise PriceImportError(f"Некорректное число: {cell.value!r}.") from None


def sheets(data: bytes) -> list[tuple[str, list[list[Cell]]]]:
    """Прочитать книгу по сигнатуре, используя сохранённые результаты формул."""
    result: list[tuple[str, list[list[Cell]]]] = []
    try:
        if data.startswith(b"PK"):
            book = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=True)
            try:
                for sheet in book:
                    result.append(
                        (
                            sheet.title,
                            [
                                [
                                    Cell(c.value, c.number_format or "General", c.data_type == "e")
                                    for c in row
                                ]
                                for row in sheet.iter_rows()
                            ],
                        )
                    )
            finally:
                book.close()
        elif data.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            book_xls = xlrd.open_workbook(file_contents=data, formatting_info=True)
            try:
                for sheet in book_xls.sheets():
                    rows = []
                    for row_index in range(sheet.nrows):
                        row = []
                        for c in sheet.row(row_index):
                            if c.xf_index is None:
                                raise PriceImportError("В ячейке XLS отсутствует формат.")
                            xf = book_xls.xf_list[c.xf_index]
                            fmt = book_xls.format_map[xf.format_key].format_str
                            row.append(Cell(c.value, fmt, c.ctype == xlrd.XL_CELL_ERROR))
                        rows.append(row)
                    result.append((sheet.name, rows))
            finally:
                book_xls.release_resources()
        else:
            raise PriceImportError("Ожидается файл XLS или XLSX.")
    except PriceImportError:
        raise
    except Exception as exc:
        raise PriceImportError(f"Не удалось прочитать Excel: {exc}") from exc
    return result


def parse(data: bytes, sheet_name: str | None) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Найти таблицу и проверить товары; неоднозначный выбор листа отклонить."""
    candidates = []
    for title, rows in sheets(data):
        if sheet_name is not None and title != sheet_name:
            continue
        for index, row in enumerate(rows[:100]):
            labels = [normalize(str(c.value or "")) for c in row]
            if all(label in labels for label in HEADERS):
                if any(labels.count(label) != 1 for label in HEADERS):
                    raise PriceImportError("Обязательные колонки повторяются.")
                candidates.append((title, rows, index, [labels.index(h) for h in HEADERS]))
                break
    if len(candidates) != 1:
        raise PriceImportError(
            "Нужен один лист с колонками Код, Артикул, Номенклатура, "
            "Остаток, 1-я цена. При нескольких листах укажите sheet_name."
        )
    title, rows, header, columns = candidates[0]
    products = []
    seen: set[str] = set()
    skipped_zero = 0
    skipped_sections = 0
    for row_number, row in enumerate(rows[header + 1 :], header + 2):
        cells = [row[i] if i < len(row) else Cell(None) for i in columns]
        try:
            code, article, name = (text(c) for c in cells[:3])
            if not code:
                # У заголовков групп нет кода и артикула, остаток пустой или нулевой.
                if article or (cells[3].value not in (None, "") and number(cells[3]) != 0):
                    raise PriceImportError("Товар без кода.")
                skipped_sections += 1
                continue
            if any(c.invalid for c in cells[:3]):
                raise PriceImportError("Ошибка Excel в идентификаторе или наименовании.")
            quantity = number(cells[3])
            if quantity == 0:
                skipped_zero += 1
                continue
            if not name:
                raise PriceImportError("Пустая номенклатура.")
            if code in seen:
                raise PriceImportError(f"Повтор кода {code}.")
            seen.add(code)
            price = (
                None if cells[4].value in (None, "") and not cells[4].invalid else number(cells[4])
            )
            products.append(
                {
                    "code": code,
                    "article": article,
                    "name": name,
                    "price": None
                    if price is None
                    else format(price.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), ".2f"),
                    "raw_price": None if price is None else str(price),
                    "quantity": str(quantity),
                }
            )
        except (PriceImportError, InvalidOperation) as exc:
            raise PriceImportError(f"Лист {title}, строка {row_number}: {exc}") from exc
    return products, {
        "sheet_name": title,
        "imported_count": len(products),
        "skipped_zero_stock": skipped_zero,
        "skipped_sections": skipped_sections,
        "missing_price_count": sum(p["price"] is None for p in products),
    }
