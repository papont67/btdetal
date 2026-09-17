"""parser.py — атомарное обновление SQLite и совместимый поиск товаров БТДеталь."""

from __future__ import annotations

import hashlib
import json
import sqlite3
from collections.abc import Iterable
from contextlib import closing
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, BinaryIO

from .excel import PriceImportError, normalize, parse

MAX_FILE_BYTES = 50 * 1024 * 1024


class BtdetalError(RuntimeError):
    """База прайса недоступна или ещё не заполнена."""


def validate_limit(value: int) -> None:
    """Проверить положительный целочисленный лимит."""
    if type(value) is not int or value < 1:
        raise ValueError("Лимит должен быть положительным целым числом.")


class BtdetalParser:
    """Хранит путь к общей базе; каждый вызов использует отдельное соединение."""

    def __init__(self, db_path: str | Path = "btdetal.sqlite3") -> None:
        """Задать постоянный файл SQLite; база создаётся только при импорте."""
        if str(db_path) == ":memory:":
            raise ValueError("Укажите путь к постоянному файлу SQLite.")
        self.db_path = Path(db_path).resolve()

    def __enter__(self) -> BtdetalParser:
        """Поддержать интерфейс контекстного менеджера других парсеров."""
        return self

    def __exit__(self, *args: object) -> None:
        """Завершить контекст; постоянных соединений у объекта нет."""

    def _connect(self, *, write: bool = False) -> sqlite3.Connection:
        """Открыть SQLite; читатель не создаёт пустой файл при неверном пути."""
        try:
            connection = sqlite3.connect(
                self.db_path.as_uri() + ("?mode=rwc" if write else "?mode=ro"),
                uri=True,
                timeout=30,
            )
            connection.row_factory = sqlite3.Row
            return connection
        except sqlite3.Error as exc:
            raise BtdetalError(f"Не удалось открыть базу прайса: {exc}") from exc

    def import_price(
        self,
        source: str | Path | bytes | BinaryIO,
        *,
        sheet_name: str | None = None,
        price_date: str | None = None,
        allow_empty: bool = False,
    ) -> dict[str, Any]:
        """Принять путь, bytes или бинарный поток и целиком заменить текущий прайс.

        price_date — необязательная строка, сохраняемая без проверки даты или формата.
        Пустой прайс по умолчанию
        отклоняется. Поток читается с текущей позиции и не закрывается библиотекой.
        Ошибка чтения, проверки или записи сохраняет предыдущую версию базы.
        """
        source_name = None
        if isinstance(source, (str, Path)):
            source_name = Path(source).name
            with Path(source).open("rb") as stream:
                data = stream.read(MAX_FILE_BYTES + 1)
        elif isinstance(source, bytes):
            data = source
        else:
            data = source.read(MAX_FILE_BYTES + 1)
        if not isinstance(data, bytes) or len(data) > MAX_FILE_BYTES:
            raise PriceImportError("Нужен бинарный файл размером не более 50 МиБ.")
        products, report = parse(data, sheet_name)
        if not products and not allow_empty:
            raise PriceImportError("Нет товаров с положительным остатком; прежняя база сохранена.")
        imported_at = datetime.now(timezone.utc).isoformat()
        metadata = {
            **report,
            "source_name": source_name,
            "price_date": price_date,
            "imported_at": imported_at,
            "sha256": hashlib.sha256(data).hexdigest(),
        }
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with closing(self._connect(write=True)) as conn:
                _ = conn.execute("PRAGMA journal_mode=WAL")
                with conn:
                    _ = conn.execute("BEGIN IMMEDIATE")
                    version = conn.execute("PRAGMA user_version").fetchone()[0]
                    if version not in (0, 1):
                        raise BtdetalError("Неподдерживаемая версия базы.")
                    _ = conn.execute("CREATE TABLE IF NOT EXISTS metadata (value TEXT NOT NULL)")
                    _ = conn.execute("""CREATE TABLE IF NOT EXISTS products (
                        code TEXT PRIMARY KEY, article TEXT NOT NULL, name TEXT NOT NULL,
                        price TEXT, raw_price TEXT, quantity TEXT NOT NULL,
                        article_norm TEXT NOT NULL, name_norm TEXT NOT NULL)""")
                    _ = conn.execute("""CREATE VIRTUAL TABLE IF NOT EXISTS product_search
                        USING fts5(article_norm, name_norm, tokenize='trigram')""")
                    _ = conn.execute("DELETE FROM products")
                    _ = conn.execute("DELETE FROM product_search")
                    _ = conn.execute("DELETE FROM metadata")
                    _ = conn.executemany(
                        "INSERT INTO products VALUES (?,?,?,?,?,?,?,?)",
                        [
                            (
                                p["code"],
                                p["article"],
                                p["name"],
                                p["price"],
                                p["raw_price"],
                                p["quantity"],
                                normalize(p["article"]),
                                normalize(p["name"]),
                            )
                            for p in products
                        ],
                    )
                    _ = conn.execute("""INSERT INTO product_search(rowid, article_norm, name_norm)
                        SELECT rowid, article_norm, name_norm FROM products""")
                    _ = conn.execute(
                        "INSERT INTO metadata VALUES (?)",
                        (json.dumps(metadata, ensure_ascii=False),),
                    )
                    _ = conn.execute("PRAGMA user_version=1")
        except sqlite3.Error as exc:
            raise BtdetalError(
                f"Не удалось обновить прайс (нужен SQLite с FTS5/trigram): {exc}"
            ) from exc
        return metadata

    def _metadata(self, conn: sqlite3.Connection) -> dict[str, Any]:
        """Получить метаданные в той же транзакции, что и найденные товары."""
        try:
            row = conn.execute("SELECT value FROM metadata").fetchone()
        except sqlite3.Error as exc:
            raise BtdetalError("Сначала импортируйте прайс.") from exc
        if row is None:
            raise BtdetalError("Сначала импортируйте прайс.")
        return json.loads(row[0])

    def get_metadata(self) -> dict[str, Any]:
        """Вернуть дату прайса, время загрузки, хеш файла и счётчики импорта."""
        with closing(self._connect()) as conn:
            return self._metadata(conn)

    def _product(self, row: sqlite3.Row, metadata: dict[str, Any]) -> dict[str, Any]:
        """Преобразовать запись в контракт Omnia/Ziphol без вымышленных URL и скидок."""
        quantity = Decimal(row["quantity"])
        return {
            "id": row["code"],
            "code": row["code"],
            "article": row["article"],
            "name": row["name"],
            "url": None,
            "image_url": None,
            "price": row["price"],
            "raw_price": row["raw_price"],
            "currency": "RUB",
            "quantity_prices": [],
            "stock_quantity": int(quantity) if quantity == int(quantity) else float(quantity),
            "in_stock": True,
            "available_to_order": False,
            "price_is_current": row["price"] is not None,
            "status": "ok" if row["price"] is not None else "error",
            **(
                {"error": "В прайсе отсутствует 1-я цена или её сохранённое значение."}
                if row["price"] is None
                else {}
            ),
            "fetched_at": metadata["imported_at"],
            "price_date": metadata["price_date"],
        }

    def search(self, article: str, *, max_results: int = 3) -> dict[str, Any]:
        """Найти буквальную подстроку в артикуле ИЛИ наименовании без учёта регистра.

        Для запросов от трёх символов используется индекс триграмм; короткие
        запросы сканируют таблицу. Символы %, _, кавычки не являются операторами.
        При превышении лимита возвращается too_many_results с пустым products.
        """
        validate_limit(max_results)
        if not isinstance(article, str) or not article.strip() or "\x00" in article:
            raise ValueError("Запрос должен быть непустой строкой без NUL.")
        query = article.strip()
        normalized = normalize(query)
        with closing(self._connect()) as conn:
            _ = conn.execute("BEGIN")
            metadata = self._metadata(conn)
            params: list[Any] = []
            clause = ""
            if len(normalized) >= 3:
                clause = "p.rowid IN (SELECT rowid FROM product_search WHERE product_search MATCH ?) AND "
                params.append('"' + normalized.replace('"', '""') + '"')
            params.extend([normalized, normalized, normalized, max_results + 1])
            rows = conn.execute(
                "SELECT p.* FROM products p WHERE "
                + clause
                + "(instr(article_norm, ?) > 0 OR instr(name_norm, ?) > 0) "
                "ORDER BY (article_norm = ?) DESC, code LIMIT ?",
                params,
            ).fetchall()
        result = {
            "query": query,
            "search_query": query,
            "products": [],
            "fetched_at": metadata["imported_at"],
            "price_date": metadata["price_date"],
        }
        if len(rows) > max_results:
            return {
                **result,
                "status": "too_many_results",
                "count_at_least": max_results + 1,
                "message": "Слишком много вариантов. Уточните запрос или увеличьте max_results.",
            }
        products = []
        for row in rows:
            exact = row["article"].casefold() == query.casefold()
            products.append(
                {
                    **self._product(row, metadata),
                    "exact_match": exact,
                    "match_type": "exact" if exact else "search_result",
                    "search_url": None,
                }
            )
        return {
            **result,
            "status": "partial_error"
            if any(p["status"] == "error" for p in products)
            else "ok"
            if products
            else "not_found",
            "count": len(products),
            "products": products,
        }

    def get_product(self, product_id: str) -> dict[str, Any]:
        """Вернуть товар по строковому коду поставщика или вызвать KeyError."""
        if not isinstance(product_id, str) or not product_id.strip():
            raise ValueError("Код должен быть непустой строкой.")
        with closing(self._connect()) as conn:
            _ = conn.execute("BEGIN")
            metadata = self._metadata(conn)
            row = conn.execute("SELECT * FROM products WHERE code=?", (product_id,)).fetchone()
            if row is None:
                raise KeyError(product_id)
            return self._product(row, metadata)

    def search_many(
        self,
        articles: Iterable[str],
        *,
        max_results: int = 3,
        max_workers: int = 1,
    ) -> dict[str, Any]:
        """Вернуть results/summary в исходном порядке, сохранив повторы запросов.

        max_workers принимается для совместимости; локальный поиск последовательный.
        Каждый запрос видит целую версию прайса, но между запросами возможен импорт.
        """
        validate_limit(max_results)
        validate_limit(max_workers)
        if isinstance(articles, (str, bytes)) or not isinstance(articles, Iterable):
            raise ValueError("Передайте коллекцию строк.")
        queries = list(articles)
        if any(not isinstance(q, str) or not q.strip() or "\x00" in q for q in queries):
            raise ValueError("Каждый запрос должен быть непустой строкой без NUL.")
        results = [self.search(q, max_results=max_results) for q in queries]
        summary = dict(
            total=len(results),
            ok=0,
            not_found=0,
            too_many_results=0,
            partial_error=0,
            error=0,
            skipped=0,
        )
        for result in results:
            summary[result["status"]] += 1
        return {
            "status": "partial_error" if summary["partial_error"] else "ok",
            "results": results,
            "summary": summary,
        }
