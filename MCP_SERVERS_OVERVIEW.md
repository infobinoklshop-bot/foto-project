# Установленные MCP серверы - Обзор

## Текущая конфигурация

**ВАЖНО:** Для Claude Code в VSCode используется файл `.mcp.json` в корне проекта!

Файл: `/Users/evgenijermakov/Documents/Фото/.mcp.json`

```json
{
  "mcpServers": {
    "memory-bank": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "@ellpepper/memory-bank-mcp@latest"
      ],
      "env": {
        "MEMORY_BANK_ROOT": "/Users/evgenijermakov/memory-bank"
      }
    },
    "filesystem": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-filesystem",
        "/Users/evgenijermakov/Documents/Фото"
      ]
    },
    "context7": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "@upstash/context7-mcp@latest"
      ]
    },
    "playwright": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "@playwright/mcp@latest"
      ]
    },
    "sequential-thinking": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-sequential-thinking@latest"
      ]
    },
    "google-sheets": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "mcp-gsheets@latest"
      ],
      "env": {
        "GOOGLE_PROJECT_ID": "gen-lang-client-0597231611",
        "GOOGLE_APPLICATION_CREDENTIALS": "/Users/evgenijermakov/Documents/Фото/.credentials/google-sheets-key.json"
      }
    }
  }
}
```

## Установленные серверы

### 1. Memory Bank 🗄️
**Назначение:** Долгосрочная память между сессиями

**Инструменты:**
- `mcp__memory-bank__list_projects` - список проектов
- `mcp__memory-bank__list_project_files` - файлы проекта
- `mcp__memory-bank__memory_bank_read` - чтение памяти
- `mcp__memory-bank__memory_bank_write` - запись памяти
- `mcp__memory-bank__memory_bank_update` - обновление памяти

**Когда использовать:**
- Сохранение архитектурных решений
- Прогресс по задачам
- Найденные проблемы и решения
- Контекст для продолжения работы

**Расположение:** `/Users/evgenijermakov/Documents/my_memory_bank/`

---

### 2. Context7 📚
**Назначение:** Актуальная документация библиотек и API

**Как использовать:**
Добавьте `use context7` в промпт:
```
use context7 - покажи актуальные методы Google Apps Script API

use context7 - как использовать OpenAI Assistants API в 2025?

use context7 - актуальная документация InSales API
```

**Когда использовать:**
- Нужна свежая документация
- Проверка breaking changes в API
- Примеры кода из официальных источников
- Новые возможности библиотек

**Провайдер:** Upstash

---

### 3. Playwright 🎭
**Назначение:** Автоматизация браузера и веб-скрейпинг

**Инструменты:** (доступны через MCP)
- Навигация по веб-страницам
- Выполнение JavaScript на странице
- Клики, ввод текста, скролл
- Получение структурированных данных
- Скриншоты и PDF

**Когда использовать:**
- Динамический контент (React, Vue, AJAX)
- Парсинг сложных сайтов поставщиков
- Автоматизация действий на веб-страницах
- Обход защиты от ботов
- Тестирование веб-интеграций

**Преимущества над UrlFetchApp:**
- ✅ Выполняет JavaScript
- ✅ Видит динамический контент
- ✅ Автоматизация действий
- ✅ Эмуляция реального браузера

**Провайдер:** Microsoft

---

### 4. Sequential Thinking 🧠
**Назначение:** Пошаговое решение сложных задач

**Когда использовать:**
- Многоэтапные задачи с зависимыми шагами
- Архитектурное планирование
- Отладка сложных проблем
- Рефакторинг кода
- Принятие технических решений

**Примеры использования:**
```
Спланируй архитектуру для нового модуля парсинга с учетом:
- Rate limiting
- Error handling
- Расширяемость для новых поставщиков

Найди и исправь проблему с дубликатами товаров пошагово
```

**Провайдер:** Anthropic (официальный)

---

### 5. Google Sheets 📊
**Назначение:** Прямой доступ к Google таблицам

**Инструменты:** (через MCP)
- Чтение данных из таблиц
- Запись и обновление данных
- Поиск и фильтрация
- Создание новых таблиц

**Когда использовать:**
- Быстрая проверка данных без открытия браузера
- Анализ статусов обработки товаров
- Поиск товаров по параметрам
- Обновление статусов из Claude Code

**Примеры использования:**
```
Покажи последние 10 товаров из таблицы "Фото товаров - Обработка изображений"

Найди все товары со статусом "Ошибка"

Обнови статус товара 26175 на "Обработано"

Сколько товаров без изображений поставщиков?
```

**Безопасность:**
- Credentials в `.credentials/google-sheets-key.json` (добавлено в .gitignore)
- Service Account: `claude-code-sheets@gen-lang-client-0597231611.iam.gserviceaccount.com`
- Доступ только к разрешенным таблицам

**Пакет:** mcp-gsheets@1.5.3  
**Документация:** [GOOGLE_SHEETS_MCP_SETUP.md](GOOGLE_SHEETS_MCP_SETUP.md)

---

### 6. Filesystem 📁
**Назначение:** Работа с файловой системой (встроенный)

**Инструменты:**
- `mcp__filesystem__read_text_file` - чтение текстовых файлов
- `mcp__filesystem__write_file` - запись файлов
- `mcp__filesystem__edit_file` - редактирование файлов
- `mcp__filesystem__list_directory` - список файлов
- `mcp__filesystem__search_files` - поиск файлов
- `mcp__filesystem__directory_tree` - дерево директорий
- `mcp__filesystem__read_multiple_files` - чтение нескольких файлов
- `mcp__filesystem__create_directory` - создание папок
- `mcp__filesystem__move_file` - перемещение/переименование
- `mcp__filesystem__get_file_info` - метаданные файла

**Приоритет:** Используйте вместо стандартных Read/Write/Edit

---

## Глобальные инструкции

Все инструкции записаны в User memory: `~/.claude/CLAUDE.md`

Применяются автоматически ко всем проектам.

---

## Активация

**ВАЖНО:** После любых изменений в `claude_desktop_config.json` нужно перезапустить Claude Code:

1. Полностью закройте приложение (⌘Q)
2. Откройте заново
3. MCP серверы загрузятся автоматически

---

## Проверка установки

После перезапуска проверьте доступность инструментов в Claude Code.

### Проверка Memory Bank
```
Покажи список проектов в memory bank
```

### Проверка Context7
```
use context7 - актуальная документация React
```

### Проверка Playwright
```
Используй Playwright для открытия https://google.com и получения заголовка
```

### Проверка Filesystem
```
Прочитай файл CLAUDE.md через mcp filesystem
```

---

## Системные требования

- ✅ Node.js v24.10.0 (требуется >= v18)
- ✅ Claude Desktop / Claude Code
- ✅ macOS (Darwin 24.6.0)

---

## Дополнительная информация

**Детальная документация:**
- [CONTEXT7_SETUP.md](CONTEXT7_SETUP.md) - Context7
- [PLAYWRIGHT_SETUP.md](PLAYWRIGHT_SETUP.md) - Playwright

**Применение к проекту:**
Все MCP инструменты доступны для улучшения текущего проекта Google Apps Script:
- Memory Bank - сохранение прогресса и решений
- Context7 - актуальная документация API (InSales, OpenAI, Google)
- Playwright - улучшенный парсинг сайтов поставщиков
- Filesystem - эффективная работа с файлами проекта

---

## Troubleshooting

### MCP сервер не загружается
1. Проверьте синтаксис JSON в конфигурации
2. Убедитесь, что Node.js установлен: `node --version`
3. Проверьте, что пакет существует: `npm view @package/name version`
4. Перезапустите Claude Code полностью

### Инструменты не появляются
1. Дождитесь полной загрузки Claude Code (10-30 секунд)
2. Проверьте логи в консоли разработчика
3. Попробуйте переустановить пакет: `npx -y @package/name@latest`

### Ошибки при выполнении
1. Проверьте права доступа к файлам/директориям
2. Для Playwright - убедитесь, что браузеры установлены
3. Для Memory Bank - проверьте путь к директории

---

## История установки

**Дата начала:** 2025-01-16  
**Последнее обновление:** 2025-01-17

**Установлено:**
1. ✅ Memory Bank (@ellpepper/memory-bank-mcp)
2. ✅ Context7 (@upstash/context7-mcp@1.0.27)
3. ✅ Playwright (@playwright/mcp@0.0.47)
4. ✅ Sequential Thinking (@modelcontextprotocol/server-sequential-thinking@2025.7.1)
5. ✅ Google Sheets (mcp-gsheets@1.5.3) - **НОВОЕ!**

**Не добавлено:**
- ❌ Fetch MCP - не нужен (есть Playwright который делает больше)
- ❌ Brave Search - не удалось получить API ключ
