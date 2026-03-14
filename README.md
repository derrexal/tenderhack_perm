# TenderHack Price Justification API

Сервис генерирует документы-обоснования НМЦК и начальной цены СТЕ на основании входных данных и агрегатов из `raw_data`.

**Что внутри**
- FastAPI приложение: `main.py`.
- Встроенная обработка `.xlsx` без внешних библиотек для чтения.
- Генерация документов: RTF внутри `.doc` и полноценный `.docx` со стилями.

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

Первый запрос может быть заметно дольше, так как идет загрузка и агрегация данных из `raw_data`.

## Эндпоинты

**1) Обоснование НМЦК по одному контракту**
- `POST /api/v1/price-justification/doc`
- Возвращает файл `justification.doc` (RTF)

Пример запроса:
```bash
curl -X POST http://localhost:8000/api/v1/price-justification/doc \
  -H "Content-Type: application/json" \
  --data ' {
    "contract": {
      "procurementName": "Поставка кухонного бытового оборудования",
      "procurementMethod": "Контракт по итогам котировочной сессии",
      "vatRate": "20%",
      "buyerInn": "9718159964",
      "buyerRegion": "Москва"
    },
    "items": [
      {
        "steId": 35927039,
        "steItemName": "Холодильник Haier MSR115 белый",
        "quantity": "1.0",
        "unit": "шт",
        "unitPrice": "23760"
      }
    ],
    "currency": "RUB",
    "maxSamples": 3
  }' \
  --output justification.doc
```

**2) Сводный отчет по нескольким контрактам**
- `POST /api/v1/price-justification/batch/doc`
- Возвращает файл `justification_batch.doc` (RTF)

Пример запроса:
```bash
curl -X POST http://localhost:8000/api/v1/price-justification/batch/doc \
  -H "Content-Type: application/json" \
  --data ' {
    "contracts": [
      {
        "contract": {
          "procurementName": "Поставка техники",
          "procurementMethod": "Контракт по итогам котировочной сессии"
        },
        "items": [
          {
            "steId": 38092313,
            "steItemName": "Печь микроволновая Samsung MS23K3513AS/BW 800 Вт",
            "quantity": "1.0",
            "unit": "шт",
            "unitPrice": "13860"
          }
        ]
      }
    ],
    "currency": "RUB",
    "maxSamples": 3,
    "reportTitle": "Сводный отчет по НМЦК"
  }' \
  --output justification_batch.doc
```

**3) Обоснование начальной цены СТЕ**
- `POST /api/v1/ste-price-justification/doc`
- Возвращает файл `ste_price_justification.docx` (DOCX со стилями)

Пример запроса:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data ' {
    "items": [
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
    ],
    "currency": "RUB",
    "reportTitle": "Обоснование начальной цены СТЕ",
    "steId": 35927039,
    "steItemName": "Холодильник Haier MSR115 белый",
    "signerName": "Иванов И.И.",
    "signerTitle": "Начальник отдела"
  }' \
  --output ste_price_justification.docx
```

## Данные

Сервис использует:
- `raw_data/TenderHack_Контракты_20260313.xlsx`
- `raw_data/TenderHack_СТЕ_20260313.xlsx`

## Примечания

- Генерация DOCX использует `python-docx` и формирует именованные стили, без локального форматирования.
- Для RTF `.doc` используется встроенный генератор на стороне сервиса.
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

## Примеры payload из `raw_data`

**Готовый payload (СТЕ)**
- Файл: `ste_payload.json`
- Команда:
```bash
curl -X POST http://localhost:8000/api/v1/ste-price-justification/doc \
  -H "Content-Type: application/json" \
  --data @ste_payload.json \
  --output ste_price_justification.docx
```

**Пример payload (НМЦК по одному контракту)**
```json
{
  "contract": {
    "procurementName": "Поставка кухонного бытового оборудования для организации работы катка",
    "procurementMethod": "Контракт по итогам котировочной сессии",
    "vatRate": "20%",
    "buyerInn": "9718159964",
    "buyerRegion": "Москва",
    "supplierInn": "7729101722",
    "supplierRegion": "Москва",
    "contractSigningDate": "2025-12-01 14:32:09.307"
  },
  "items": [
    {
      "steId": 35927039,
      "steItemName": "Холодильник Haier MSR115 белый",
      "quantity": "1.0",
      "unit": "шт",
      "unitPrice": "23760.00"
    },
    {
      "steId": 38092313,
      "steItemName": "Печь микроволновая Samsung MS23K3513AS/BW 800 Вт",
      "quantity": "1.0",
      "unit": "шт",
      "unitPrice": "13860.00"
    }
  ],
  "currency": "RUB",
  "maxSamples": 3
}
```

## Оптимизация первого запроса

- Первый вызов делает агрегацию данных из `raw_data`. Это нормально и может занять время.
- Чтобы прогреть сервис, сделайте разовый запрос к любому эндпоинту сразу после старта.
- Для стабильного времени отклика держите процесс прогретым и не перезапускайте его между запросами.
