"""upload.py — пример передачи загруженного сайтом файла в библиотеку."""

from pathlib import Path
from typing import BinaryIO

from btdetal import BtdetalParser


def import_uploaded_file(file: BinaryIO, db_path: Path) -> dict:
    """Обработать бинарный поток из обработчика загрузки сайта."""
    parser = BtdetalParser(db_path)
    return parser.import_price(file)


def find_prices(db_path: Path, article: str) -> dict:
    """Выполнить поиск без повторного открытия Excel."""
    return BtdetalParser(db_path).search(article, max_results=10)
