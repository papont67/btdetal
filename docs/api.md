# API btdetal-prices 0.1

`BtdetalParser(db_path="btdetal.sqlite3")` не открывает Excel и не создаёт базу
до первого импорта. Поддерживается `with BtdetalParser(...) as parser`.

| Метод | Результат |
| --- | --- |
| `import_price(source, *, sheet_name=None, price_date=None, allow_empty=False)` | Метаданные нового прайса |
| `search(article, *, max_results=3)` | Одна поисковая выдача |
| `search_many(articles, *, max_results=3, max_workers=1)` | `status`, `results`, `summary` |
| `get_product(product_id)` | Карточка по точному строковому коду; `KeyError`, если отсутствует |
| `get_metadata()` | Дата, время импорта, хеш и счётчики |

`source` — путь (`str`/`Path`), `bytes` или бинарный поток.
`price_date` — необязательная строка, сохраняемая как передана, без проверки
формата, давности или сравнения с предыдущим прайсом. На импорт не влияет.
`max_workers` проверяется и принимается для совместимости с сетевыми
парсерами; быстрый локальный пакетный поиск выполняется последовательно.
Каждый запрос согласован внутри себя, но пакет может увидеть обновление
между запросами. Порядок и повторы входных запросов сохраняются.

Пример ответа (иллюстрация структуры):

```json
{
  "query": "198741",
  "search_query": "198741",
  "status": "ok",
  "count": 1,
  "fetched_at": "2026-09-17T13:00:00+00:00",
  "price_date": "2026-09-14",
  "products": [{
    "id": "00000021143",
    "code": "00000021143",
    "article": "198741",
    "name": "Втулка венчика блендера Gorenje, посадка одна лыська",
    "price": "62.00",
    "raw_price": "62",
    "currency": "RUB",
    "quantity_prices": [],
    "stock_quantity": 8,
    "in_stock": true,
    "available_to_order": false,
    "price_is_current": true,
    "status": "ok",
    "exact_match": true,
    "match_type": "exact",
    "url": null,
    "image_url": null,
    "search_url": null,
    "fetched_at": "2026-09-17T13:00:00+00:00",
    "price_date": "2026-09-14"
  }]
}
```

`exact_match` — равенство артикула запросу без учёта регистра, а не совпадение
с наименованием. При частичном совпадении `match_type="search_result"`.
Точные совпадения выводятся первыми, затем товары упорядочены по коду.

`raw_price` хранит исходное десятичное значение; `price` округляется до копеек
по ROUND_HALF_UP. Остаток возвращается JSON-числом. `fetched_at` показывает
время импорта, а не время запроса. Наличие и актуальность относятся к прайсу.

`too_many_results`: пустой `products`, `count_at_least=max_results+1`, `message`;
поля `count` нет, поскольку вся выдача не подсчитывается.
`not_found`: `count=0`, пустой `products`.

В `summary` сохраняются ключи Omnia/Ziphol: `total`, `ok`, `not_found`,
`too_many_results`, `partial_error`, `error`, `skipped`. При отсутствии цены
товар содержит `price=null`, `raw_price=null`, `status="error"`, пояснение `error`
и `price_is_current=false`; поисковый результат и пакет получают `partial_error`.
`error` и `skipped` в summary равны нулю: ошибки хранилища поднимаются исключениями.

Метаданные: `sheet_name`, `imported_count`, `skipped_zero_stock`,
`skipped_sections`, `missing_price_count`, `source_name` (null для bytes/потока), `price_date`,
`imported_at`, `sha256`. `price_date` хранится только как справочная информация;
`imported_at` — время загрузки файла, не дата его составления.

Исключения: `PriceImportError` (подкласс `ValueError`) — повреждённый файл,
неверная таблица, числа или дубли кодов; `BtdetalError` — недоступная или
неинициализированная база, ошибка атомарной записи; `ValueError` — неверный
запрос/лимит; `OSError` — ошибка чтения источника или создания директории.
CLI выводит ошибку JSON в stderr и завершает работу с кодом 1.
