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

Пример запроса (использует готовый файл `ste_payload_new.json` в корне репозитория):
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_payload_new.json \
  --output ste_price_justification.docx
```

## Примечания

- Документы формируются на основе шаблона `report_template.docx` в корне репозитория.
- Если структура шаблона меняется, обновите заполнение в `build_ste_price_docx_from_template`.

## Модели (контракты)

**StePriceTemplateRequest**
```json
{
  "contractName": "Наименование закупки",
  "summaryPrice": 5123,
  "position": {
    "positionName": "Наименование предмета закупки",
    "positionPrice": 12345,
    "items": [
      {
        "contractId": "205604468",
        "procurementMethod": "Контракт по итогам конкурентной процедуры",
        "initialContractValue": "49200.00000",
        "contractValueAfterSigning": "49200.00000",
        "reductionPercent": "0.00000000000000",
        "contractSigningDate": "2025-01-13 00:00:00.000",
        "buyerInn": "7720093523",
        "supplierInn": "7804710797",
        "supplierRegion": "Санкт-Петербург",
        "quantity": "1",
        "unit": "шт",
        "steId": 38499393,
        "steItemName": "Медеран пор. лиофил. д/приг. р-ра д/инъекц. и инф. 50 мг фл/в компл. с р-лем №1х1 Meone HealthCare Pvt.Ltd Индия",
        "unitPrice": "164.00000000000"
      }
    ]
  },
  "currency": "RUB"
}
```

## Примеры payload

**Готовый payload (СТЕ)**  
Файл: `ste_payload_new.json`

Команда:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_payload_new.json \
  --output ste_price_justification.docx
```
