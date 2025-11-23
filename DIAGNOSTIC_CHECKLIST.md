# 🔍 Чек-лист диагностики проблем

Используйте этот чек-лист для систематической диагностики любых проблем в проекте.

---

## 📸 Фаза 1: Сбор информации

### Скриншоты и логи
- [ ] Скриншот ошибки из Google Apps Script IDE
- [ ] Execution log (View → Execution log)
- [ ] Stackdriver Logging (View → Stackdriver Logging)
- [ ] Скриншот из Google Sheets (если проблема с данными)
- [ ] Скриншот из InSales (если проблема с интеграцией)

### Описание проблемы
- [ ] **Что происходит?** Конкретные симптомы
- [ ] **Что ожидалось?** Желаемое поведение
- [ ] **Когда началось?** Работало ли раньше?
- [ ] **Условия воспроизведения:** Артикул товара, функция, меню
- [ ] **Частота:** Всегда / Иногда / Один раз

---

## 🔎 Фаза 2: Категоризация проблемы

### Определите тип проблемы:

#### 🌐 API Integration
- [ ] InSales API (загрузка товаров, изображений)
- [ ] OpenAI API (GPT-4 Vision, Assistants)
- [ ] Другие внешние сервисы

**Действия:**
1. ✅ **use context7** - проверить актуальность документации API
2. Читать код через `mcp__filesystem__read_text_file`
3. Сохранить находки в `api-notes.md`

#### 🕷️ Web Parsing
- [ ] Veber.ru парсинг не работает
- [ ] Sturman.ru парсинг не работает
- [ ] Селекторы устарели

**Действия:**
1. ✅ **mcp__playwright__browser_navigate** - открыть страницу
2. ✅ **mcp__playwright__browser_snapshot** - получить структуру
3. ✅ **mcp__playwright__browser_evaluate** - извлечь данные
4. Обновить `parsing-guides.md`

#### 📊 Data Processing
- [ ] Нормализация характеристик
- [ ] Проблемы с UNNORMALIZED_VALUES
- [ ] Ре-нормализация не работает
- [ ] Invalid JSON в SPECIFICATIONS_RAW

**Действия:**
1. Читать код: `06_specification_normalizer.js`
2. ✅ **mcp__google-sheets__sheets_get_values** - проверить данные
3. Проверить справочник параметров (колонка H - normalizer function)
4. Сохранить в `issues.md`

#### 🖼️ Image Processing
- [ ] AI не генерирует alt-теги
- [ ] SEO filenames некорректные
- [ ] Проблемы с загрузкой в InSales

**Действия:**
1. Проверить OpenAI API key в Script Properties
2. Читать код: `04_image_processing.js`
3. use context7 - OpenAI Vision API актуальные параметры

#### 🎨 UI/UX
- [ ] Диалог не открывается
- [ ] Данные не сохраняются
- [ ] Autocomplete не работает

**Действия:**
1. Читать HTML файл: `UnifiedParameterDialog.html` или другой
2. Проверить server-side функции в соответствующем модуле
3. Browser DevTools (если можно воспроизвести в браузере)

---

## 🛠️ Фаза 3: Диагностические команды

### Чтение кода (MCP Filesystem)
```bash
# Читать весь файл
mcp__filesystem__read_text_file(path: "06_specification_normalizer.js")

# Только начало (экономия токенов)
mcp__filesystem__read_text_file(path: "06_specification_normalizer.js", head: 100)

# Поиск по паттерну
mcp__filesystem__search_files(path: ".", pattern: "*normalizer*")
```

### Проверка данных (Google Sheets MCP)
```bash
# Читать конкретный диапазон
mcp__google-sheets__sheets_get_values(
  spreadsheetId: "your_id",
  range: "Обработка изображений!B:D"
)

# Справочник параметров
mcp__google-sheets__sheets_get_values(
  range: "Справочник параметров!A:H"
)

# Batch чтение
mcp__google-sheets__sheets_batch_get_values(
  ranges: ["A:A", "P:Q"] # Article + Specs
)
```

### Парсинг сайтов (Playwright)
```bash
# Открыть страницу
mcp__playwright__browser_navigate(url: "https://veber.ru/product/26175")

# Получить структуру (accessibility tree)
mcp__playwright__browser_snapshot()

# Извлечь данные через JavaScript
mcp__playwright__browser_evaluate(
  function: "() => {
    return {
      images: Array.from(document.querySelectorAll('img')).map(i => i.src),
      title: document.querySelector('h1').textContent
    }
  }"
)

# Скриншот (только если нужно показать пользователю)
mcp__playwright__browser_take_screenshot(filename: "debug.png")
```

### Проверка API документации (Context7)
```bash
# Найти библиотеку
mcp__context7__resolve-library-id(libraryName: "InSales API")

# Получить документацию
mcp__context7__get-library-docs(
  context7CompatibleLibraryID: "/insales/api",
  topic: "image upload",
  page: 1
)
```

---

## 📝 Фаза 4: Документирование

### Обязательно сохранить в Memory Bank:

#### issues.md
```markdown
## [2025-11-17] Проблема: Нормализация водозащиты не работает

**Статус:** 🟢 Решена

**Симптомы:**
- Параметр "Защита от влаги" со значением "да" не нормализуется
- Остается в UNNORMALIZED_VALUES

**Причина:**
- Пустое поле normalizer function в колонке H справочника
- Этап 1 нормализации (custom normalizer) пропускается
- Этап 2 (enum matching) ищет "да" вместо "Влагозащищенный"

**Решение:**
- Добавлено значение "normalizeWaterproofing" в колонку H:15
- Запущена ре-нормализация через UnifiedParameterDialog

**Файлы:** 
- [06_specification_normalizer.js:450](06_specification_normalizer.js#L450)
- [Справочник параметров H:15](...)

**Тэги:** #normalization #enum #waterproofing
```

#### api-notes.md (если использовали Context7)
```markdown
## InSales API - Image Upload (проверено: 2025-11-17)

**Endpoint:** `POST /admin/products/{id}/images.json`

**Ключевые изменения:**
- С 2024 требуется Content-Type: application/json
- Поддержка WebP формата
- Лимит размера: 10MB

**Файл проекта:** [03_insales_api.js:uploadImageToProduct()](03_insales_api.js#L234)
```

#### parsing-guides.md (если использовали Playwright)
```markdown
## Veber.ru - Карточка товара (проверено: 2025-11-17)

**URL Pattern:** `https://veber.ru/product/{id}/`

**Селекторы:**
- Изображения: `.product-gallery img[data-src]`
- Цена: `.price-current`
- Характеристики: `.specs-table tr > td`

**Особенности:**
- Lazy loading (data-src вместо src)
- JSON-LD schema в <script type="application/ld+json">

**Playwright код:**
\`\`\`javascript
const images = await page.evaluate(() => 
  Array.from(document.querySelectorAll('.product-gallery img'))
    .map(img => img.dataset.src || img.src)
);
\`\`\`

**Файл проекта:** [05_supplier_parsing.gs.js:parseVeberImages()](05_supplier_parsing.gs.js#L123)
```

---

## ✅ Фаза 5: Решение и верификация

### Перед исправлением:
- [ ] Понятна root cause проблемы
- [ ] Есть план решения (для сложных → Sequential Thinking)
- [ ] Определены файлы для изменения
- [ ] Создана backup (git commit или CSV export)

### Во время исправления:
- [ ] Использую `mcp__filesystem__edit_file` для изменений
- [ ] Комментарии в коде объясняют "почему", не "что"
- [ ] Обновляю CLAUDE.md если изменилась архитектура

### После исправления:
- [ ] Протестировано в Google Apps Script IDE
- [ ] Обновлен Memory Bank (issues.md с пометкой 🟢 Решена)
- [ ] Сделан git commit с описанием
- [ ] Обновлен current-task.md

---

## 🚨 Быстрая диагностика типичных проблем

### "Нормализация не работает"
1. ✅ Проверить колонку H в справочнике (normalizer function заполнен?)
2. ✅ Проверить колонку P (SPECIFICATIONS_RAW - валидный JSON?)
3. ✅ Читать UNNORMALIZED_VALUES через ScriptProperties
4. Использовать диагностическую функцию из кода

### "API возвращает ошибку"
1. ✅ **use context7** - актуальная документация API
2. Проверить Script Properties (ключи установлены?)
3. Проверить лимиты и квоты
4. Exponential backoff на 429 работает?

### "Парсинг не находит элементы"
1. ✅ **Playwright browser_navigate** - открыть страницу
2. ✅ **browser_snapshot** - проверить структуру
3. Сайт использует JavaScript рендеринг? (UrlFetchApp не подходит!)
4. Обновить селекторы в parsing-guides.md

### "Execution timeout (6 минут)"
1. Уменьшить batch size (50 → 30 товаров)
2. Добавить `SpreadsheetApp.flush()` периодически
3. Разбить на несколько функций
4. Использовать time-based triggers

### "Invalid JSON в SPECIFICATIONS_RAW"
1. Проверить товар через `mcp__google-sheets__sheets_get_values`
2. Если JSON невалиден → переимпорт товара
3. Ре-нормализация НЕ работает с невалидным JSON

---

## 📊 Шаблон отчета для пользователя

После диагностики предоставить пользователю:

```markdown
## 🔍 Диагностика завершена

### Проблема
[Краткое описание проблемы]

### Root Cause
[Найденная причина с ссылками на код]

### Решение
[План исправления]

### Изменения
- [file.js:123](file.js#L123) - описание изменения

### Как проверить
1. Шаг 1
2. Шаг 2
3. Ожидаемый результат

### Документация
- Сохранено в Memory Bank: issues.md
- [Дополнительная информация]
```

---

**Этот чек-лист поможет систематически подходить к любой проблеме в проекте!**
