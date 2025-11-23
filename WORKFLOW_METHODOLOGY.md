# 🚀 Методология эффективной работы с проектом через MCP серверы

## 📊 Принципы работы

### 1. Правило "Документация сначала"
- **ВСЕГДА** используйте Context7 перед работой с API (InSales, OpenAI, Google Apps Script)
- Парсинг сайтов через Playwright вместо догадок о структуре
- Чтение актуальной документации экономит часы отладки

### 2. Правило "Сохраняй контекст"
- Memory Bank - ваша долговременная память проекта
- Сохраняйте ВСЕ важные находки: архитектуру, баги, решения
- Перед завершением сессии - обязательная запись прогресса

### 3. Правило "Делим и властвуем"
- Сложная задача → Sequential Thinking для планирования
- План → разбивка на подзадачи с отдельными чатами
- Каждый чат решает ОДНУ конкретную проблему

### 4. Правило "Экономим токены"
- MCP Filesystem вместо встроенного Read/Write
- Playwright snapshot вместо screenshot (структурированные данные)
- Memory Bank для передачи контекста между сессиями

---

## 🔍 ФАЗА 1: Диагностика проблем

### Шаг 1.1: Сбор информации о проблеме

**Инструменты:**
```bash
# Вы предоставляете:
- Скриншоты ошибок
- Описание симптомов
- Ожидаемое поведение

# Я использую:
- mcp__filesystem__read_text_file - читаю код
- mcp__google-sheets__sheets_get_values - проверяю данные
- Playwright (если проблема в веб-интеграции)
```

**Пример команды:**
> "Вот скриншот ошибки при нормализации. Проверь код в 06_specification_normalizer.js и таблицу 'Справочник параметров'"

### Шаг 1.2: Анализ с использованием актуальной документации

**Триггеры для Context7:**
- Ошибки API (InSales, OpenAI)
- Deprecated методы
- Неожиданное поведение библиотек

**Пример:**
```javascript
// ❌ Ошибка: "InSales API returns 400"
// ✅ Действие:
use context7 - InSales API upload images endpoint parameters 2025
// → Узнаем, что поменялись обязательные поля
```

### Шаг 1.3: Документирование проблемы в Memory Bank

**Структура файла `issues.md`:**
```markdown
## [Дата] Проблема: Короткое описание

**Симптомы:**
- Что происходит
- Условия воспроизведения

**Причина:**
- Найденная root cause

**Решение:**
- Что было сделано

**Файлы:** [file.js:123](file.js#L123)
```

**Команда сохранения:**
```javascript
mcp__memory-bank__memory_bank_write(
  projectName: "Фото", // Имя рабочей директории
  fileName: "issues.md",
  content: "..."
)
```

---

## 📋 ФАЗА 2: Планирование решения

### Шаг 2.1: Sequential Thinking для сложных задач

**Когда использовать:**
- Задача требует >3 этапов
- Есть зависимости между шагами
- Нужно выбрать архитектурный подход

**Пример использования:**
```
Используй Sequential Thinking:

Задача: Реализовать систему кэширования InSales API запросов

Требования:
- Минимизировать количество запросов
- Учесть rate limits
- Инвалидация при изменении данных
- Google Apps Script ограничения (6 мин, нет файловой системы)

Предложи архитектуру и разбей на подзадачи.
```

**Выход:**
- Подробный план с этапами
- Список подзадач для отдельных чатов
- Оценка рисков и зависимостей

### Шаг 2.2: Разбивка на подзадачи

**Критерии выделения подзадачи:**
- Касается одного модуля/файла
- Может быть решена независимо
- Имеет четкий критерий готовности
- Занимает <30 минут работы

**Формат подзадачи для Memory Bank:**
```markdown
## Подзадача 1: Создать Cache менеджер

**Цель:** Базовый класс для кэширования через ScriptProperties

**Входные данные:**
- Структура данных для кэша
- TTL настройки

**Выходные данные:**
- CacheManager в 1_shared_utilities.js
- Методы: get, set, invalidate, clear

**Зависимости:** Нет

**Критерий готовности:**
- Unit тесты проходят
- Документация в коде

**Отдельный чат:** ДА (создать новую сессию)
```

### Шаг 2.3: Сохранение плана

**Memory Bank структура:**
```
projectName: "Фото"

Файлы:
- architecture.md - общая архитектура и принципы
- current-task.md - текущая задача и прогресс
- issues.md - найденные проблемы и решения
- api-notes.md - важные особенности API из Context7
- parsing-guides.md - селекторы и структура сайтов (Playwright)
```

---

## 🛠️ ФАЗА 3: Реализация с использованием MCP

### Шаг 3.1: Работа с файлами через MCP Filesystem

**ВСЕГДА используйте MCP инструменты:**

```javascript
// ✅ ПРАВИЛЬНО:
mcp__filesystem__read_text_file(path: "00_config.gs.js")
mcp__filesystem__edit_file(
  path: "06_specification_normalizer.js",
  edits: [{
    oldText: "// старый код",
    newText: "// новый код"
  }]
)

// ❌ НЕПРАВИЛЬНО:
Read(file_path: "00_config.gs.js")
Edit(file_path: "06_specification_normalizer.js", ...)
```

**Преимущества:**
- Единый интерфейс для всех проектов
- Меньше токенов на операцию
- Поддержка batch операций

### Шаг 3.2: Парсинг сайтов через Playwright

**Когда использовать:**
- Нужно обновить селекторы в 05_supplier_parsing.gs.js
- Отладка парсинга (почему не находит элементы)
- Анализ новых поставщиков
- Тестирование InSales магазина

**Рабочий процесс:**
```bash
# 1. Открыть страницу
mcp__playwright__browser_navigate(url: "https://veber.ru/product/26175")

# 2. Получить структуру
mcp__playwright__browser_snapshot()
# → Accessibility tree (не нужно анализировать пиксели!)

# 3. Извлечь данные
mcp__playwright__browser_evaluate(
  function: "() => {
    return {
      images: Array.from(document.querySelectorAll('.product-image img')).map(img => img.src),
      price: document.querySelector('.price').textContent
    }
  }"
)

# 4. Сохранить селекторы в Memory Bank
```

**Сохранение результатов:**
```markdown
# parsing-guides.md

## Veber.ru - Структура карточки товара

**URL pattern:** `https://veber.ru/product/{id}/`

**Селекторы (проверено 2025-11-17):**
- Изображения: `.product-gallery img[data-src]`
- Цена: `.price-current`
- Характеристики: `.specs-table tr`
  - Название: `td:first-child`
  - Значение: `td:last-child`

**Особенности:**
- Используют lazy loading (data-src вместо src)
- Цена в формате "1 234,56 ₽"
- JSON-LD schema в <script type="application/ld+json">
```

### Шаг 3.3: Работа с Google Sheets через MCP

**Быстрая проверка данных без открытия браузера:**

```javascript
// Проверить сколько товаров со статусом "Ошибка"
mcp__google-sheets__sheets_get_values(
  spreadsheetId: "ваш_id",
  range: "Обработка изображений!K:K", // Колонка статусов
  valueRenderOption: "FORMATTED_VALUE"
)

// Найти товар по артикулу
mcp__google-sheets__sheets_get_values(
  range: "Обработка изображений!B:D" // Article, InSales ID, Name
)

// Обновить статус
mcp__google-sheets__sheets_update_values(
  range: "Обработка изображений!K2",
  values: [["Исправлено ✅"]]
)
```

**Когда НЕ использовать:**
- Сложная логика обработки → пишите Apps Script код
- Batch операции >100 строк → используйте встроенные функции проекта
- Нужна транзакционность → только Apps Script

### Шаг 3.4: Получение актуальной документации

**Context7 workflow:**

```bash
# 1. Resolve library ID
mcp__context7__resolve-library-id(libraryName: "InSales API")

# 2. Get documentation
mcp__context7__get-library-docs(
  context7CompatibleLibraryID: "/insales/api",
  topic: "image upload",
  page: 1
)

# 3. Сохранить важное в Memory Bank
```

**Сохранение API notes:**
```markdown
# api-notes.md

## InSales API - Image Upload (проверено через Context7: 2025-11-17)

**Endpoint:** `POST /admin/products/{id}/images.json`

**Обязательные поля:**
- `image[src]` - URL или base64
- `image[filename]` - с расширением!
- `image[position]` - порядок в галерее

**Изменения с 2024:**
- Теперь требуется `Content-Type: application/json`
- Поддержка WebP формата
- Лимит размера: 10MB

**Пример:**
\`\`\`javascript
// см. 03_insales_api.js:uploadImageToProduct()
\`\`\`
```

---

## 🔄 ФАЗА 4: Управление контекстом и чатами

### Стратегия: Одна задача = Один чат

**Главный чат (этот):**
- Планирование и координация
- Код-ревью результатов подзадач
- Интеграция изменений
- Обновление Memory Bank

**Вспомогательные чаты:**
- Решение конкретной подзадачи
- Фокус на одном модуле
- Закрывается после завершения
- Результаты сохраняются в Memory Bank

### Передача контекста между чатами

**Экспорт контекста (в конце вспомогательного чата):**
```markdown
Сохрани в Memory Bank:

Файл: subtask-1-cache-manager.md

---
## Подзадача 1: CacheManager - Выполнено ✅

**Дата:** 2025-11-17

**Что сделано:**
- Создан CacheManager класс в 1_shared_utilities.js:450-520
- Реализованы методы: get, set, invalidate, clear
- TTL через timestamp сравнение
- Unit тесты добавлены

**Изменения в коде:**
- [1_shared_utilities.js:450-520](1_shared_utilities.js#L450-L520)

**Использование:**
\`\`\`javascript
const cache = new CacheManager('insales_products', 3600); // 1 час
cache.set('product_123', productData);
const data = cache.get('product_123'); // null если истек TTL
\`\`\`

**Следующие шаги:**
- Интегрировать в 03_insales_api.js
- Добавить invalidation при обновлении товаров

**Проблемы/Ограничения:**
- ScriptProperties лимит 9KB на значение
- Нужна логика очистки старых ключей
---
```

**Импорт контекста (в начале следующего чата):**
```
Новая подзадача:

Прочитай из Memory Bank контекст:
- current-task.md - общий план
- subtask-1-cache-manager.md - готовый CacheManager

Задача: Интегрировать CacheManager в 03_insales_api.js для кэширования loadProductsFromCategory()
```

### Очистка контекста для экономии токенов

**Правила:**
1. **Не загружай весь код** - используй `mcp__filesystem__search_files` для поиска нужных участков
2. **Читай только измененные файлы** - остальное в Memory Bank
3. **Используй Playwright snapshot** вместо screenshot (меньше токенов)
4. **Context7 → Memory Bank** - прочитал документацию → сохранил ключевые моменты → не читай повторно

**Пример экономии:**
```bash
# ❌ Неэффективно (50000 токенов):
Read(file_path: "05_supplier_parsing.gs.js") # Весь файл 2000 строк
Read(file_path: "06_specification_normalizer.js") # Еще 1500 строк
# + скриншот страницы (10000 токенов)

# ✅ Эффективно (5000 токенов):
mcp__memory-bank__memory_bank_read(
  projectName: "Фото",
  fileName: "parsing-guides.md" # Только селекторы
)
mcp__filesystem__search_files(
  path: ".",
  pattern: "*normalizer*" # Найти нужный файл
)
mcp__filesystem__read_text_file(
  path: "06_specification_normalizer.js",
  head: 100 # Только первые 100 строк
)
mcp__playwright__browser_snapshot() # Структурированные данные
```

---

## 📝 ФАЗА 5: Документирование и коммит

### Обновление Memory Bank после решения

**Обязательные обновления:**
```javascript
// 1. Архитектурные изменения
mcp__memory-bank__memory_bank_update(
  projectName: "Фото",
  fileName: "architecture.md",
  content: "... добавить описание CacheManager ..."
)

// 2. Решенные проблемы
mcp__memory-bank__memory_bank_update(
  projectName: "Фото",
  fileName: "issues.md",
  content: "... закрыть issue с пометкой РЕШЕНО ..."
)

// 3. Прогресс задачи
mcp__memory-bank__memory_bank_update(
  projectName: "Фото",
  fileName: "current-task.md",
  content: "... обновить статусы подзадач ..."
)
```

### Git коммит

**После завершения задачи:**
```bash
git add .
git commit -m "feat: Добавлена система кэширования InSales API

- CacheManager в 1_shared_utilities.js
- Интеграция в loadProductsFromCategory()
- TTL 1 час, автоочистка старых записей

🤖 Generated with Claude Code
Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## 🎯 Практические примеры workflow

### Пример 1: Отладка нормализации параметров

**Проблема:** Параметр "Защита от влаги" не нормализуется

**Workflow:**

```bash
# ШАГ 1: Диагностика
Покажи скриншот ошибки

Я использую:
- mcp__filesystem__read_text_file("06_specification_normalizer.js")
- mcp__google-sheets__sheets_get_values(range: "Справочник параметров!A:H")

Нахожу: normalizer function пустой в колонке H

# ШАГ 2: Проверка документации
use context7 - Google Apps Script PropertiesService limits
→ Узнаю, что UNNORMALIZED_VALUES может переполниться

# ШАГ 3: Сохранение проблемы
mcp__memory-bank__memory_bank_write(
  fileName: "issues.md",
  content: "## 2025-11-17: Нормализация водозащиты
  
  **Причина:** Пустое поле normalizer в H колонке
  **Решение:** fixWaterproofingNormalizerInReference()
  **Код:** [06_specification_normalizer.js:450](06_specification_normalizer.js#L450)"
)

# ШАГ 4: Исправление
mcp__filesystem__edit_file(...)

# ШАГ 5: Проверка в таблице
mcp__google-sheets__sheets_update_values(
  range: "Справочник параметров!H15",
  values: [["normalizeWaterproofing"]]
)

# ШАГ 6: Коммит
```

### Пример 2: Добавление нового поставщика

**Задача:** Добавить парсинг Sturman.ru

**Workflow:**

```bash
# ШАГ 1: Sequential Thinking
Используй Sequential Thinking:

Задача: Добавить парсинг Sturman.ru
- Изучить структуру сайта
- Написать parser функцию
- Добавить в SUPPLIERS_CONFIG
- Протестировать на 3 товарах

# → План с 5 этапами

# ШАГ 2: Подзадачи
Подзадача 1: Исследование структуры Sturman.ru (отдельный чат)
Подзадача 2: Реализация парсера (этот чат)
Подзадача 3: Интеграция и тесты (этот чат)

# ШАГ 3: Подзадача 1 - Новый чат
Используй Playwright для анализа https://sturman.ru/product/12345

mcp__playwright__browser_navigate(...)
mcp__playwright__browser_snapshot()
mcp__playwright__browser_evaluate(
  function: "() => {
    return {
      images: [...],
      specs: [...],
      price: ...
    }
  }"
)

Сохрани селекторы в Memory Bank (parsing-guides.md)

# ШАГ 4: Подзадача 2 - Этот чат
Прочитай из Memory Bank селекторы

mcp__memory-bank__memory_bank_read(fileName: "parsing-guides.md")

Реализуй parseSturmanImages() в 05_supplier_parsing.gs.js

mcp__filesystem__edit_file(...)

# ШАГ 5: Тестирование через Playwright
Открой Sturman, запусти функцию, сравни результат

# ШАГ 6: Документирование
mcp__memory-bank__memory_bank_update(
  fileName: "architecture.md",
  content: "... добавить Sturman в список поставщиков ..."
)
```

### Пример 3: Обновление InSales API интеграции

**Задача:** InSales поменял формат ответа API

**Workflow:**

```bash
# ШАГ 1: Context7 - актуальная документация
use context7 - InSales API products endpoint response format 2025

# → Нахожу: поле `variants` теперь `product_variants`

# ШАГ 2: Сохранение в Memory Bank
mcp__memory-bank__memory_bank_write(
  fileName: "api-notes.md",
  content: "## InSales API Breaking Change (2025-11)
  
  **Изменение:** variants → product_variants
  **Файлы:** 03_insales_api.js:loadProductVariants()
  **Migration:** обратная совместимость через fallback"
)

# ШАГ 3: Обновление кода
mcp__filesystem__edit_file(
  path: "03_insales_api.js",
  edits: [{
    oldText: "product.variants || []",
    newText: "product.product_variants || product.variants || []"
  }]
)

# ШАГ 4: Тестирование через Google Sheets
Menu: "📥 Загрузить товары"
→ Проверяем, что SKU извлекаются

# ШАГ 5: Коммит с пометкой breaking change
```

---

## ✅ Чек-лист эффективной работы

### Перед началом задачи:
- [ ] Прочитал Memory Bank: current-task.md, issues.md
- [ ] Проверил актуальность API через Context7
- [ ] Составил план через Sequential Thinking (для сложных задач)
- [ ] Разбил на подзадачи с четкими критериями готовности

### Во время работы:
- [ ] Использую mcp__filesystem__* вместо встроенных Read/Write/Edit
- [ ] Playwright для парсинга вместо UrlFetchApp (если нужен JS рендеринг)
- [ ] Google Sheets MCP для быстрой проверки данных
- [ ] Сохраняю важные находки в Memory Bank по ходу

### После завершения:
- [ ] Обновил Memory Bank (architecture.md, issues.md, current-task.md)
- [ ] Закрыл вспомогательные чаты, сохранив результаты
- [ ] Сделал git commit с подробным описанием
- [ ] Проверил, что не осталось TODO в коде

### Экономия токенов:
- [ ] Не читаю весь код - только нужные участки
- [ ] Использую search_files для поиска
- [ ] Playwright snapshot вместо screenshot
- [ ] Context7 → Memory Bank (не перечитываю документацию)
- [ ] Передаю контекст через Memory Bank между чатами

---

## 🚨 Частые ошибки и как их избежать

### ❌ Ошибка 1: Работа без плана
**Симптом:** Прыгаешь между файлами, забываешь что делал

**Решение:**
- Всегда начинай с Sequential Thinking для задач >3 шагов
- Сохраняй план в current-task.md
- Используй TodoWrite для отслеживания прогресса

### ❌ Ошибка 2: Использование устаревшей документации
**Симптом:** Код работал раньше, теперь API возвращает ошибки

**Решение:**
- ВСЕГДА используй Context7 перед работой с внешними API
- Сохраняй дату проверки в api-notes.md
- При ошибках API - первым делом Context7

### ❌ Ошибка 3: Потеря контекста между сессиями
**Симптом:** Не помнишь, что делал вчера, приходится перечитывать код

**Решение:**
- Обязательное обновление Memory Bank в конце сессии
- Формат: что сделано, что осталось, следующие шаги
- Читай current-task.md в начале новой сессии

### ❌ Ошибка 4: Перерасход токенов
**Симптом:** Быстро достигаешь лимита, приходится начинать новый чат

**Решение:**
- Используй head/tail параметры при чтении файлов
- Playwright snapshot (не screenshot) для веб-страниц
- Memory Bank вместо повторного чтения документации
- Вспомогательные чаты для подзадач

### ❌ Ошибка 5: Не используешь Playwright
**Симптом:** Парсинг не работает, селекторы устарели, страница рендерится через JS

**Решение:**
- Playwright для ЛЮБОГО парсинга сайтов
- Сохраняй селекторы в parsing-guides.md с датой проверки
- Используй browser_snapshot для структурированных данных

---

## 🎓 Шпаргалка команд

### Memory Bank
```javascript
// Создать/прочитать/обновить
mcp__memory-bank__memory_bank_write(projectName: "Фото", fileName: "...", content: "...")
mcp__memory-bank__memory_bank_read(projectName: "Фото", fileName: "...")
mcp__memory-bank__memory_bank_update(projectName: "Фото", fileName: "...", content: "...")

// Список файлов
mcp__memory-bank__list_project_files(projectName: "Фото")
```

### Filesystem
```javascript
// Чтение/запись
mcp__filesystem__read_text_file(path: "...", head: 100) // Только первые 100 строк
mcp__filesystem__write_file(path: "...", content: "...")
mcp__filesystem__edit_file(path: "...", edits: [{oldText: "...", newText: "..."}])

// Поиск
mcp__filesystem__search_files(path: ".", pattern: "*normalizer*")
```

### Playwright
```javascript
// Навигация и анализ
mcp__playwright__browser_navigate(url: "...")
mcp__playwright__browser_snapshot() // Структурированные данные
mcp__playwright__browser_evaluate(function: "() => { return {...} }")

// Скриншот только когда нужно показать пользователю
mcp__playwright__browser_take_screenshot(filename: "debug.png")
```

### Google Sheets
```javascript
// Чтение
mcp__google-sheets__sheets_get_values(spreadsheetId: "...", range: "A:C")
mcp__google-sheets__sheets_batch_get_values(spreadsheetId: "...", ranges: ["A:A", "K:K"])

// Запись
mcp__google-sheets__sheets_update_values(spreadsheetId: "...", range: "K2", values: [["Готово ✅"]])
```

### Context7
```javascript
// Поиск и чтение документации
mcp__context7__resolve-library-id(libraryName: "InSales API")
mcp__context7__get-library-docs(context7CompatibleLibraryID: "/insales/api", topic: "...", page: 1)
```

### Sequential Thinking
```javascript
// Для сложных задач с планированием
mcp__sequential-thinking__sequentialthinking(
  thought: "Шаг 1: Анализирую структуру...",
  thoughtNumber: 1,
  totalThoughts: 5,
  nextThoughtNeeded: true
)
```

---

## 📚 Структура Memory Bank для проекта "Фото"

```
Фото/
├── architecture.md          # Общая архитектура, модули, зависимости
├── current-task.md          # Текущая задача, план, прогресс
├── issues.md                # История проблем и решений
├── api-notes.md             # Особенности API из Context7
├── parsing-guides.md        # Селекторы и структура сайтов
├── conventions.md           # Стиль кода, паттерны
└── subtasks/
    ├── subtask-1-*.md       # Выполненные подзадачи
    └── subtask-2-*.md
```

---

## 🎯 Готовы начать?

**Следующие шаги:**

1. **Загрузите скриншоты проблем** - я проанализирую через Playwright/Filesystem
2. **Опишите задачу** - я составлю план через Sequential Thinking
3. **Я создам Memory Bank** - сохраню текущее состояние проекта
4. **Разобьем на подзадачи** - каждая в своем чате с передачей контекста

**Готов работать эффективно! 🚀**
