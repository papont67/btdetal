"""cli.py — команды импорта, поиска и просмотра состояния прайса."""

import argparse
import json
import sys

from . import BtdetalError, BtdetalParser, PriceImportError


def main() -> int:
    """Выполнить команду и вывести JSON; ошибки завершить ненулевым кодом."""
    parser = argparse.ArgumentParser(description="Прайсы БТДеталь")
    _ = parser.add_argument("--db", default="btdetal.sqlite3", help="Постоянный файл SQLite")
    commands = parser.add_subparsers(dest="command", required=True)
    importer = commands.add_parser("import", help="Заменить прайс из XLS/XLSX")
    _ = importer.add_argument("file")
    _ = importer.add_argument("--sheet-name")
    _ = importer.add_argument("--price-date", help="Необязательная дата прайса, без проверки")
    _ = importer.add_argument("--allow-empty", action="store_true")
    search = commands.add_parser("search", help="Поиск по артикулу и наименованию")
    _ = search.add_argument("queries", nargs="+")
    _ = search.add_argument("--max-results", type=int, default=3)
    _ = commands.add_parser("info", help="Метаданные загруженного прайса")
    args = parser.parse_args()
    client = BtdetalParser(args.db)
    try:
        if args.command == "import":
            result = client.import_price(
                args.file,
                sheet_name=args.sheet_name,
                price_date=args.price_date,
                allow_empty=args.allow_empty,
            )
        elif args.command == "info":
            result = client.get_metadata()
        elif len(args.queries) == 1:
            result = client.search(args.queries[0], max_results=args.max_results)
        else:
            result = client.search_many(args.queries, max_results=args.max_results)
    except (BtdetalError, PriceImportError, OSError, ValueError) as exc:
        print(
            json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0
