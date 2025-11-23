# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Google Apps Script** проект для автоматизации обработки изображений товаров и интеграции с e-commerce. Система управляет изображениями из InSales, парсит данные от поставщиков и использует AI для SEO-оптимизации.

**Технологии:** Google Apps Script, InSales API, OpenAI GPT-4 Vision, Web scraping (UrlFetchApp)

**Runtime:** V8 | **Timezone:** Europe/Moscow

## Google Sheets Access

**Spreadsheet ID:** `1WbHHZmGErFJgxxIPBQ4My-gYJV6PBEqkP7TczzEGOac`

| Лист | Sheet ID | Назначение |
|------|----------|------------|
| Обработка изображений | 0 | Основные данные товаров (A-AB) |
| Справочник параметров | 549854820 | Маппинг характеристик |
| Выгрузка | 408711254 | Экспорт в InSales |

## Quick Reference

### Команды разработки (clasp)
```bash
clasp push          # Деплой в Google Apps Script
clasp pull          # Получить удалённые изменения
clasp open          # Открыть проект в браузере
```

### Точки входа меню (99_menu.js)
| Пункт меню | Функция | Модуль |
|-----------|----------|--------|
| 📥 Загрузить товары | `loadProductsFromInSalesMenu()` | 03_insales_api |
| 🔍 Спарсить у поставщиков | `showSupplierParsingDialog()` | 05_supplier_parsing |
| 🆕 Импорт товаров | `showFullProductImportDialog()` | 05_supplier_parsing |
| 🤖 Обработать изображения | `showImageSelectionForProcessing()` | 04_image_processing |
| 🔄 Управление параметрами | `showUnifiedParameterDialog()` | 06_specification_normalizer |
| 📤 Создать в InSales | `createProductsInInsalesMenu()` | 03_insales_api |

### Тестирование
- Запуск функций: Apps Script IDE → `Run`
- Проверка настроек: Меню `🖼️ Фото` → `⚙️ Проверить настройки API`
- Логи: Apps Script IDE → `View` → `Execution log`

## Architecture

### Модульная структура

Нумерация файлов определяет порядок загрузки и зависимости (выше → ниже):

```
00_config.gs.js              → Константы, настройки API
1_shared_utilities.js        → Логирование, обработка ошибок
02_data_manager.js           → CRUD операции Google Sheets
03_insales_api.js            → Интеграция с InSales
04_image_processing.js       → OpenAI GPT-4 Vision
05_supplier_parsing.gs.js    → Парсинг поставщиков (Veber, Sturman)
06_specification_normalizer.js → Нормализация характеристик
07_description_ai.js         → AI-рерайт описаний (3 ассистента)
08_product_matcher.js        → Поиск дубликатов
99_menu.js                   → UI меню и оркестрация
*.html                       → Диалоги (HtmlService)
data/                        → CSV бэкапы, JSON справочники
```

### Структура данных (Google Sheet)

**Основные колонки (A-L):**
| Колонка | Константа | Назначение |
|---------|-----------|------------|
| A | CHECKBOX | Выбор товаров |
| B | ARTICLE | Артикул (primary key) |
| C | INSALES_ID | ID в InSales |
| D | PRODUCT_NAME | Название |
| E | ORIGINAL_IMAGES | URL из InSales (\\n-separated) |
| F | SUPPLIER_IMAGES | URL от поставщиков |
| G | ADDITIONAL_IMAGES | Доп. фото |
| H | PROCESSED_IMAGES | Обработанные |
| I | ALT_TAGS | Alt-теги |
| J | SEO_FILENAMES | SEO имена файлов |
| K | PROCESSING_STATUS | Статус обработки |
| L | INSALES_STATUS | Статус отправки |

**Расширенные колонки для импорта (M-AB):**
| Колонка | Константа | Назначение |
|---------|-----------|------------|
| M-O | DESCRIPTION, DESCRIPTION_REWRITTEN, SHORT_DESCRIPTION | Описания |
| P-Q | SPECIFICATIONS_RAW, SPECIFICATIONS_NORMALIZED | Характеристики (JSON) |
| R-Y | PRICE, STOCK, CATEGORIES, BRAND, SERIES, WEIGHT, DIMENSIONS, PACKAGE_CONTENTS | Данные товара |
| Z-AB | MATCH_STATUS, MATCH_CONFIDENCE, IMPORT_STATUS | Статусы импорта |

Доступ через константы: `IMAGES_COLUMNS.ARTICLE`, `IMAGES_COLUMNS.ALT_TAGS` и т.д.

## Main Workflows

### 1. Обработка изображений
```
loadProductsFromInSalesMenu() → Загрузка товаров из InSales
    ↓
showSupplierParsingDialog() → Парсинг доп. изображений (опционально)
    ↓
showImageSelectionForProcessing() → AI генерирует alt-теги и SEO-имена
    ↓
sendProcessedImagesToInSales() → Загрузка обратно в InSales
```

### 2. Полный импорт товаров
```
showFullProductImportDialog() → Ввод артикулов, выбор поставщика
    ↓
parseVeberFullProduct() / parseSturmanFullProduct() → Парсинг карточки
    ↓
normalizeSpecifications() → Нормализация характеристик по справочнику
    ↓
generateProductDescription() → AI-рерайт описания (Copier → Editor)
    ↓
checkProductDuplicate() → Проверка дубликатов
    ↓
writeFullProductData() → Запись в таблицу (колонки M-AB)
    ↓
createProductsInInsalesMenu() → Создание товаров в InSales
```

### 3. Управление параметрами (UnifiedParameterDialog)
Диалог с двумя вкладками:
- **📦 Товар**: Маппинг параметров конкретного товара
- **📋 Все ненормализованные**: Массовая обработка ошибок enum

## Configuration

### API ключи (Script Properties)

```javascript
// Просмотр/установка: Apps Script IDE → Project Settings → Script Properties
const properties = PropertiesService.getScriptProperties();
properties.setProperty('insalesApiKey', 'your_key');
```

**Обязательные:**
- `insalesApiKey`, `insalesPassword`, `insalesShop` — InSales API
- `openaiApiKey` — OpenAI (для AI обработки)

**AI Assistants (hardcoded в 07_description_ai.js):**
- `AI_ASSISTANTS.COPIER` — Генерация описаний
- `AI_ASSISTANTS.EDITOR` — Редактирование

**Опциональные:** `replicateToken`, `tinypngKey`, `imgbbKey`, `telegramToken`

## Specification Normalization System

### Двухэтапный процесс (порядок критичен!)

```javascript
// 1️⃣ Кастомный нормализатор (если указан в справочнике)
if (normalizerName && NORMALIZER_FUNCTIONS[normalizerName]) {
  normalizedValue = NORMALIZER_FUNCTIONS[normalizerName](value);
  // "да" → "Влагозащищенный"
}

// 2️⃣ Enum matching (проверка допустимых значений)
if (fieldType === 'enum' && allowedValues) {
  normalizedValue = normalizeEnumValue(normalizedValue, allowedValues);
}
```

**Важно:** Если поменять порядок — сырые значения не пройдут через нормализатор.

### Справочник параметров

Лист `SHEET_NAMES.SPEC_REFERENCE` ("Справочник параметров"):

| Колонка | Назначение |
|---------|------------|
| B | Параметр (каноническое имя) |
| C | Тип (enum/число/текст) |
| D | Допустимые значения (для enum) |
| F-G | Синонимы Veber / Sturman |
| H | Функция нормализации |

### UNNORMALIZED_VALUES

Хранит значения, не прошедшие enum-валидацию:
```javascript
UNNORMALIZED_VALUES.add(paramName, rawValue, allowedValues, supplierField, article);
UNNORMALIZED_VALUES.getAll();  // Для диалога "Ненормализованные значения"
```

### Ре-нормализация

После изменения справочника вызовите `renormalizeProduct(article)` — перечитает `SPECIFICATIONS_RAW` (колонка P) и пересчитает `SPECIFICATIONS_NORMALIZED` (колонка Q).

### Troubleshooting нормализации

| Симптом | Причина | Решение |
|---------|---------|---------|
| "Не найдено совпадение для enum" | Пустой normalizer в колонке H | Заполнить имя функции |
| Изменения не применяются | Нужна ре-нормализация | `renormalizeProduct(article)` |
| "JSON НЕВАЛИДЕН" | Колонка P не JSON | Переимпортировать товар |
| Артикул не найден | Тип данных (string vs number) | Проверить логи |

## Important Implementation Details

### SKU из вариантов
```javascript
// ПРАВИЛЬНО: артикул из варианта
const article = variant.sku;

// НЕПРАВИЛЬНО: product.sku часто пустой
const article = product.sku;
```

### Многострочные данные
```javascript
// Запись
const stored = imageUrls.join('\n');

// Чтение
const urls = cellValue.split('\n').filter(url => url.trim());
```

### Rate Limiting
```javascript
Utilities.sleep(1000);  // 1 сек между запросами парсинга

// Для API — экспоненциальный backoff при 429
if (responseCode === 429) {
  Utilities.sleep(Math.pow(2, attempt) * 1000);
}
```

### Лимиты
- **Batch**: max 50 товаров (6-минутный таймаут Apps Script)
- **Full import**: max 30 товаров (AI обработка ~2 мин/товар)
- **URLFetch**: 20,000 запросов/день

## Adding New Features

### Новый парсер поставщика
1. Добавить в `SUPPLIERS_CONFIG` (05_supplier_parsing.gs.js)
2. Реализовать `parseNewSupplierImages(url)`
3. Добавить case в `executeSupplierParser()`

### Новая колонка
1. Добавить в `IMAGES_COLUMNS` (00_config.gs.js)
2. Обновить `setupHeaders()` (02_data_manager.js)
3. Обновить read/write функции

### Новый нормализатор
```javascript
// 06_specification_normalizer.js
NORMALIZER_FUNCTIONS.myNormalizer = (value) => transformedValue;
```
Затем указать `myNormalizer` в колонке H справочника.

### Новый HTML диалог
```html
<!-- MyDialog.html -->
<!DOCTYPE html>
<html>
<head><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-gray-900 text-gray-100 p-6">
  <div id="loading">Loading...</div>
  <div id="main" class="hidden"><!-- UI --></div>
  <script>
    google.script.run
      .withSuccessHandler(data => {
        document.getElementById('loading').classList.add('hidden');
        document.getElementById('main').classList.remove('hidden');
      })
      .yourServerFunction();
  </script>
</body>
</html>
```

```javascript
// 99_menu.js
function showMyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MyDialog')
    .setWidth(800).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Title');
}
```

**UI темы:** Dark (bg-gray-900) для большинства диалогов, Light (bg-white) для UnifiedParameterDialog.

## Code Style

- **Язык:** Русские комментарии, английские API
- **Логи:** Emoji-префиксы (✅ ❌ ⚠️ 📝)
- **Константы:** UPPERCASE_SNAKE_CASE
- **Функции:** camelCase
- **Статусы:** Использовать `STATUS_VALUES` объект

## Files Reference

| Файл | Назначение |
|------|------------|
| `.clasp.json` | Конфигурация clasp (scriptId) |
| `appsscript.json` | Манифест Apps Script (runtime V8, timezone) |
| `data/JSON-справочник.json` | Справочник параметров (JSON) |
| `data/*.csv` | CSV бэкапы таблиц |

## Resources

- [InSales API](https://www.insales.ru/collection/api)
- [Google Apps Script](https://developers.google.com/apps-script/reference)
- [OpenAI Vision API](https://platform.openai.com/docs/guides/vision)
- [clasp](https://github.com/google/clasp)
