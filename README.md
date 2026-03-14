# TenderHack Price Justification API

Сервис генерирует документ-обоснование начальной цены СТЕ на основании входных данных.

**Что внутри**
- FastAPI приложение: `main.py`.
- Встроенная обработка `.xlsx` без внешних библиотек для чтения.
- Генерация документов: `.docx` со стилями и табличным оформлением.

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

Пример запроса:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_payload.json \
  --output ste_price_justification.docx
```

## Примечания

- Генерация DOCX использует `python-docx` и формирует именованные стили, без локального форматирования.
- Документы формируются через `python-docx` с именованными стилями.
- Если таблица в DOCX должна быть уже или шире, параметры задаются в `build_ste_price_docx`.

## Модели (контракты)

**Ste**
```json
{
  "id": 35927039,
  "name": "Холодильник Haier MSR115 белый",
  "category": "Холодильники",
  "manufacturer": "Haier",
  "characteristics": "Объем:92;Цвет:белый"
}
```

**Contract**
```json
{
  "id": 204746787,
  "procurementName": "Поставка кухонного бытового оборудования",
  "procurementMethod": "Контракт по итогам котировочной сессии",
  "initialContractValue": "71091.90",
  "contractValueAfterSigning": "71091.90",
  "reductionPercent": "0.00",
  "vatRate": "20%",
  "contractSigningDate": "2025-12-01 14:32:09.307",
  "buyerInn": "9718159964",
  "buyerRegion": "Москва",
  "supplierInn": "7729101722",
  "supplierRegion": "Москва"
}
```

**ContractItemWithSte**
```json
{
  "id": 1,
  "steId": 35927039,
  "steItemName": "Холодильник Haier MSR115 белый",
  "quantity": "1.0",
  "unit": "шт",
  "unitPrice": "23760.00",
  "steName": "Холодильник Haier MSR115 белый",
  "steCategory": "Холодильники",
  "steManufacturer": "Haier",
  "steCharacteristics": "Объем:92;Цвет:белый"
}
```

**StePriceJustificationRow**
```json
{
  "contractId": "204746787",
  "procurementMethod": "Контракт по итогам котировочной сессии",
  "initialContractValue": "71091.90",
  "contractValueAfterSigning": "71091.90",
  "reductionPercent": "0.00",
  "contractSigningDate": "2025-12-01 14:32:09.307",
  "buyerInn": "9718159964",
  "supplierInn": "7729101722",
  "steId": 35927039,
  "steItemName": "Холодильник Haier MSR115 белый",
  "unitPrice": "23760.00"
}
```

## Примеры payload

**Готовый payload (СТЕ)**  
Файл: `ste_payload.json`

Команда:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_payload.json \
  --output ste_price_justification.docx
```
