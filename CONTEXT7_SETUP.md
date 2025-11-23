# Context7 MCP Server - Установка завершена ✓

## Что установлено

Context7 MCP сервер добавлен в конфигурацию Claude Desktop для получения актуальной документации библиотек и фреймворков.

## Конфигурация

Файл: `/Users/evgenijermakov/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "context7": {
      "command": "npx",
      "args": ["-y", "@upstash/context7-mcp@latest"]
    }
  }
}
```

## Как использовать

### В промптах Claude Code

Добавьте фразу `use context7` в ваш запрос:

**Примеры:**
```
use context7 - как использовать React hooks в 2025?

use context7 - покажи актуальный синтаксис TypeScript 5.x

use context7 - какие методы есть в современной Google Apps Script API для работы с листами?

use context7 - покажи примеры использования Tailwind CSS v4
```

### Когда использовать

- ✅ Нужна актуальная документация (последняя версия библиотеки)
- ✅ Примеры кода из официальных источников
- ✅ Новые API методы, появившиеся после 2024 года
- ✅ Миграция между версиями фреймворков

### Когда НЕ нужно

- ❌ Базовые вопросы по языку (JavaScript, Python основы)
- ❌ Работа с локальным кодом проекта
- ❌ Общие концепции программирования

## Активация

**ВАЖНО:** Для активации Context7 нужно перезапустить Claude Code/Claude Desktop!

1. Закройте Claude Code полностью
2. Откройте заново
3. Context7 будет доступен автоматически

## Проверка установки

После перезапуска можно проверить командой:
```
use context7 - покажи документацию для Google Apps Script SpreadsheetApp
```

## Дополнительные возможности

### API ключ (опционально)

Для повышенных лимитов и приватных репозиториев можно добавить API ключ:

1. Получить ключ: https://context7.com/dashboard
2. Добавить в конфигурацию:

```json
{
  "mcpServers": {
    "context7": {
      "command": "npx",
      "args": [
        "-y",
        "@upstash/context7-mcp@latest",
        "--api-key",
        "YOUR_API_KEY"
      ]
    }
  }
}
```

## Технические детали

- **Пакет**: `@upstash/context7-mcp@1.0.27`
- **Node.js**: v24.10.0 ✓ (требуется >= v18)
- **Провайдер**: Upstash
- **GitHub**: https://github.com/upstash/context7

## Глобальные инструкции

Инструкции по использованию Context7 добавлены в User memory:
`~/.claude/CLAUDE.md` - применяется ко всем проектам автоматически
