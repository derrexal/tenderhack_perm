# TenderHack Price Justification API

Сервис генерирует документ-обоснование начальной цены СТЕ на основании входных данных.

**Что внутри**
- FastAPI приложение: `main.py`.
- Генерация документов `.docx` на основе шаблона `report_template.docx` в корне репозитория.
- Работа со стилями и таблицами через `python-docx`.

**Требования**
- Python 3.10+

**Установка**
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**Запуск**
```bash
uvicorn main:app --reload
```

## Эндпоинты

**1) Обоснование начальной цены СТЕ**
- `POST /api/v1/ste-price-justification/doc`
- Возвращает файл `ste_price_justification.docx` (DOCX со стилями)

**Быстрый старт (curl)**  
Использует готовый файл `ste_price_request_example.json` в корне репозитория.
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_price_request_example.json \
  --output ste_price_justification.docx
```

## Примечания

- Документы формируются на основе шаблона `report_template.docx` в корне репозитория.
- Если структура шаблона меняется, обновите заполнение в `build_ste_price_docx_from_template`.
- Параметр `docType` позволяет запросить `docx`, `doc` или `pdf` (по умолчанию `docx`).
- Каждая запись в `positions` создает отдельную таблицу и заголовок в документе.
- Для конвертации в `doc` и `pdf` требуется LibreOffice на машине, где работает API (команда `soffice`).
- Наличие Microsoft Word на клиентском компьютере не влияет на конвертацию — она происходит на сервере.

## Модели (контракты)

**StePriceTemplateRequest**
```json
{
  "contractName": "Наименование закупки",
  "summaryPrice": 5123,
  "positions": [
    {
      "positionName": "Наименование предмета закупки",
      "positionPrice": 12345,
      "items": [
        {
          "steId": 38499393,
          "steName": "Медеран пор. лиофил. д/приг. р-ра д/инъекц. и инф. 50 мг фл/в компл. с р-лем №1х1 Meone HealthCare Pvt.Ltd Индия",
          "contractId": "205604468",
          "contractSigningDate": "2025-01-13 00:00:00.000",
          "buyerInn": "7720093523",
          "buyerRegion": "Москва",
          "count": 100,
          "unit": "шт.",
          "unitPrice": "164.00000000000",
          "nds": "20%"
        }
      ]
    },
    {
      "positionName": "Наименование предмета закупки 2",
      "positionPrice": 24690,
      "items": [
        {
          "steId": 38499393,
          "steName": "Медеран пор. лиофил. д/приг. р-ра д/инъекц. и инф. 50 мг фл/в компл. с р-лем №1х1 Meone HealthCare Pvt.Ltd Индия",
          "contractId": "205604587",
          "contractSigningDate": "2025-01-13 00:00:00.000",
          "buyerInn": "7720093523",
          "buyerRegion": "Москва",
          "count": 150,
          "unit": "шт.",
          "unitPrice": "164.00000000000",
          "nds": "20%"
        }
      ]
    }
  ],
  "currency": "RUB",
  "docType": "docx"
}
```

## Примеры payload

**Готовый payload (СТЕ)**  
Файл: `ste_price_request_example.json`

Команда:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_price_request_example.json \
  --output ste_price_justification.docx
```
