# 🐛 Руководство по отладке Google Apps Script через MCP

## 🎯 Три способа работы с логами

---

## 📺 Способ 1: clasp logs (Рекомендуется для быстрой проверки)

### Преимущества
- ✅ Самый быстрый способ
- ✅ Не требует открытия браузера
- ✅ Показывает логи сразу в терминале
- ✅ Можно фильтровать и следить в реальном времени

### Использование через Bash MCP

```bash
# Показать последние логи
clasp logs

# Упрощенный формат
clasp logs --simplified

# Следить за логами в реальном времени
clasp logs --watch

# Открыть Stackdriver Logging в браузере
clasp logs --open
```

### Пример в Claude Code

```
Запусти через Bash MCP:
clasp logs --simplified

Покажи последние 20 строк логов проекта Apps Script
```

### Что показывает
- `Logger.log()` выводы
- `console.log()` выводы
- `console.error()` ошибки
- Время выполнения
- Имя функции

### Ограничения
- Только последние логи (не исторические)
- Нет визуального интерфейса
- Не показывает Execution transcript

---

## 🌐 Способ 2: Playwright (Для визуальной отладки)

### Преимущества
- ✅ Полный доступ к Script Editor UI
- ✅ Можно запускать функции
- ✅ Видно Execution log с деталями
- ✅ Доступ к Stackdriver Logging
- ✅ Можно взаимодействовать с UI (хотя код лучше редактировать локально)

### Workflow через Playwright MCP

#### Шаг 1: Получить URL Script Editor

```bash
# Через Bash MCP
clasp open --webapp

# Или прочитать из .clasp.json
cat .clasp.json | grep scriptId
# → URL: https://script.google.com/d/{scriptId}/edit
```

#### Шаг 2: Открыть в Playwright

```javascript
// Навигация
mcp__playwright__browser_navigate(
  url: "https://script.google.com/d/YOUR_SCRIPT_ID/edit"
)

// Подождать загрузки (нужна авторизация!)
mcp__playwright__browser_wait_for(time: 3)

// Сделать скриншот для проверки
mcp__playwright__browser_take_screenshot(
  filename: "script-editor.png"
)
```

#### Шаг 3: Запустить функцию

```javascript
// Клик на функцию в dropdown
mcp__playwright__browser_click(
  element: "Select function dropdown",
  ref: "combobox[aria-label='Select function']"
)

// Выбрать функцию
mcp__playwright__browser_select_option(
  element: "Function selector",
  ref: "select",
  values: ["testFunction"]
)

// Клик Run
mcp__playwright__browser_click(
  element: "Run button",
  ref: "button[aria-label='Run']"
)

// Подождать выполнения
mcp__playwright__browser_wait_for(time: 5)
```

#### Шаг 4: Получить логи из Execution log

```javascript
// Открыть Execution log
mcp__playwright__browser_click(
  element: "View > Execution log",
  ref: "menuitem[aria-label='Execution log']"
)

// Извлечь текст логов
mcp__playwright__browser_evaluate(
  function: "() => {
    const logPanel = document.querySelector('.execution-log-panel');
    return logPanel ? logPanel.innerText : 'No logs found';
  }"
)
```

### Важно: Авторизация

При первом запуске Playwright откроет браузер **БЕЗ** авторизации Google. Нужно:

**Вариант A: Использовать существующую сессию Chrome**
```javascript
// В конфигурации Playwright указать userDataDir
mcp__playwright__browser_navigate(
  url: "https://script.google.com/...",
  // Используется профиль браузера с авторизацией
)
```

**Вариант B: Авторизоваться один раз**
```javascript
// 1. Открыть браузер
mcp__playwright__browser_navigate(url: "https://script.google.com")

// 2. Сделать скриншот - увидите форму входа
mcp__playwright__browser_take_screenshot()

// 3. Вручную войти в браузер Playwright (окно не закрывается)

// 4. После авторизации продолжить автоматизацию
```

### Пример в Claude Code

```
Используй Playwright для отладки:

1. Открой Script Editor моего проекта (URL из clasp open)
2. Запусти функцию testNormalization()
3. Покажи логи из Execution log
4. Сделай скриншот результата
```

### Ограничения
- Требует авторизации Google
- Медленнее чем clasp logs
- Playwright может не распознать все элементы UI (сложный интерфейс Apps Script)

---

## 🔧 Способ 3: Диагностические функции в коде (Для production)

### Преимущества
- ✅ Детальное логирование в нужных местах
- ✅ Persistence через ScriptProperties
- ✅ Не зависит от внешних инструментов
- ✅ Можно включать/выключать через DEBUG_MODE

### Реализация

#### Шаг 1: Добавить DEBUG_MODE в config

```javascript
// 00_config.gs.js

const DEBUG_CONFIG = {
  ENABLED: false, // Включить через Script Properties
  VERBOSE: false, // Подробные логи
  SAVE_TO_PROPERTIES: true, // Сохранять в ScriptProperties
  MAX_LOG_ENTRIES: 100 // Максимум записей
};

// Проверка через Script Properties (приоритет)
function isDebugMode() {
  const props = PropertiesService.getScriptProperties();
  const debugMode = props.getProperty('DEBUG_MODE');
  return debugMode === 'true' || DEBUG_CONFIG.ENABLED;
}
```

#### Шаг 2: Создать debug logger

```javascript
// 1_shared_utilities.js

function debugLog(context, message, data = null) {
  if (!isDebugMode()) return;
  
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    context,
    message,
    data: data ? JSON.stringify(data) : null
  };
  
  // Console log
  console.log(`[DEBUG] ${context}: ${message}`);
  if (data) console.log('Data:', data);
  
  // Сохранение в ScriptProperties
  if (DEBUG_CONFIG.SAVE_TO_PROPERTIES) {
    saveDebugLog(logEntry);
  }
}

function saveDebugLog(logEntry) {
  try {
    const props = PropertiesService.getScriptProperties();
    const logsJson = props.getProperty('DEBUG_LOGS') || '[]';
    const logs = JSON.parse(logsJson);
    
    // Добавить новую запись
    logs.push(logEntry);
    
    // Ограничить размер
    if (logs.length > DEBUG_CONFIG.MAX_LOG_ENTRIES) {
      logs.shift(); // Удалить самую старую
    }
    
    props.setProperty('DEBUG_LOGS', JSON.stringify(logs));
  } catch (error) {
    console.error('Failed to save debug log:', error);
  }
}

function getDebugLogs(limit = 50) {
  const props = PropertiesService.getScriptProperties();
  const logsJson = props.getProperty('DEBUG_LOGS') || '[]';
  const logs = JSON.parse(logsJson);
  return logs.slice(-limit); // Последние N записей
}

function clearDebugLogs() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('DEBUG_LOGS');
  console.log('Debug logs cleared');
}
```

#### Шаг 3: Использовать в коде

```javascript
// 06_specification_normalizer.js

function normalizeSpecifications(rawSpecs, supplier, article) {
  debugLog('normalizeSpecifications', 'Started', {
    supplier,
    article,
    rawSpecsKeys: Object.keys(rawSpecs)
  });
  
  // ... ваш код ...
  
  if (normalizerFunction) {
    debugLog('normalizeSpecifications', 'Applying normalizer', {
      parameter: canonicalName,
      normalizer: normalizerName,
      rawValue: value
    });
    
    value = NORMALIZER_FUNCTIONS[normalizerName](value);
    
    debugLog('normalizeSpecifications', 'Normalizer result', {
      parameter: canonicalName,
      normalizedValue: value
    });
  }
  
  // ... остальной код ...
  
  debugLog('normalizeSpecifications', 'Completed', {
    article,
    normalizedCount: Object.keys(normalized).length
  });
  
  return normalized;
}
```

#### Шаг 4: Меню для управления DEBUG_MODE

```javascript
// 99_menu.js

function showDebugMenu() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentMode = props.getProperty('DEBUG_MODE') === 'true';
  
  const response = ui.alert(
    '🐛 Debug Mode',
    `Текущий статус: ${currentMode ? 'Включен' : 'Выключен'}\n\nИзменить?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    props.setProperty('DEBUG_MODE', (!currentMode).toString());
    ui.alert(`Debug mode ${!currentMode ? 'включен' : 'выключен'}`);
  }
}

function showDebugLogs() {
  const logs = getDebugLogs(50);
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: monospace; font-size: 12px; }
        .log-entry { border-bottom: 1px solid #ccc; padding: 10px; }
        .timestamp { color: #666; }
        .context { font-weight: bold; color: #0066cc; }
        .data { background: #f5f5f5; padding: 5px; margin-top: 5px; }
      </style>
    </head>
    <body>
      <h2>🐛 Debug Logs (Last 50)</h2>
      ${logs.map(log => `
        <div class="log-entry">
          <div class="timestamp">${log.timestamp}</div>
          <div class="context">${log.context}</div>
          <div>${log.message}</div>
          ${log.data ? `<div class="data"><pre>${log.data}</pre></div>` : ''}
        </div>
      `).join('')}
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Debug Logs');
}

// Добавить в меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🖼️ Фото')
    // ... существующие пункты ...
    .addSeparator()
    .addSubMenu(ui.createMenu('🐛 Debug')
      .addItem('Toggle Debug Mode', 'showDebugMenu')
      .addItem('View Debug Logs', 'showDebugLogs')
      .addItem('Clear Debug Logs', 'clearDebugLogs'))
    .addToUi();
}
```

### Использование через Claude Code

```
Включи DEBUG_MODE в проекте:

1. Добавь диагностические функции в 1_shared_utilities.js
2. Добавь debugLog() вызовы в normalizeSpecifications()
3. Создай меню Debug в 99_menu.js

После этого можно:
- Включить debug mode через Script Properties
- Запустить функцию в Apps Script
- Прочитать логи через showDebugLogs()
```

---

## 📊 Сравнение способов

| Критерий | clasp logs | Playwright | Debug Functions |
|----------|-----------|-----------|-----------------|
| **Скорость** | ⚡ Мгновенно | 🐌 Медленно | ⚡ Мгновенно |
| **Удобство** | ✅ Просто | ⚠️ Сложно | ✅ Просто |
| **Авторизация** | ✅ Через clasp | ❌ Нужна вручную | ✅ Не требуется |
| **Детальность** | ⭐⭐ Базовая | ⭐⭐⭐ Полная | ⭐⭐⭐ Настраиваемая |
| **История** | ❌ Только последние | ✅ Полная в UI | ✅ Persistence |
| **Production** | ⚠️ Не подходит | ❌ Не подходит | ✅ Идеально |
| **Визуальная отладка** | ❌ Нет | ✅ Да | ❌ Нет |

---

## 🎯 Рекомендации по использованию

### Для быстрой проверки
```bash
clasp logs --simplified
```
**Когда:** Проверить что функция отработала, посмотреть простые логи

### Для отладки UI/диалогов
```javascript
mcp__playwright__browser_navigate(...)
mcp__playwright__browser_take_screenshot(...)
```
**Когда:** Проблемы с HTML диалогами, нужно видеть визуально что происходит

### Для production мониторинга
```javascript
debugLog('context', 'message', data)
showDebugLogs() // Через меню в Sheets
```
**Когда:** Отслеживание проблем в production, детальное логирование процессов

### Для отладки парсинга
```javascript
// Комбинация Playwright + debug logs
mcp__playwright__browser_navigate(url: "https://veber.ru/...")
mcp__playwright__browser_evaluate(...) // Извлечь данные
debugLog('parsing', 'Veber result', extractedData)
```
**Когда:** Проверить что парсится с сайта, отладить селекторы

---

## 💡 Лучшие практики

### 1. Начинайте с clasp logs
Самый быстрый способ проверить базовую работу функций.

### 2. Добавьте debug functions для сложных мест
Критические участки кода (нормализация, парсинг) → подробное логирование.

### 3. Используйте Playwright для визуальной отладки
Только когда нужно увидеть UI или взаимодействовать с браузером.

### 4. Всегда выключайте DEBUG_MODE в production
```javascript
// В конце отладки:
PropertiesService.getScriptProperties().deleteProperty('DEBUG_MODE');
```

### 5. Логируйте контекст, не только значения
```javascript
// ❌ Плохо
debugLog('norm', value);

// ✅ Хорошо
debugLog('normalizeSpecifications', 'Enum matching failed', {
  parameter: canonicalName,
  value: value,
  allowedValues: allowedValues,
  article: article
});
```

---

## 🚀 Быстрый старт: Добавить отладку в проект

### Команда для Claude Code:

```
Добавь debug функционал в проект:

1. Прочитай 1_shared_utilities.js
2. Добавь DEBUG_MODE конфиг и функции debugLog, getDebugLogs, clearDebugLogs
3. Обнови 99_menu.js - добавь Debug submenu
4. Добавь debugLog() в критические места:
   - normalizeSpecifications() в 06_specification_normalizer.js
   - parseVeberFullProduct() в 05_supplier_parsing.gs.js
   - uploadImageToProduct() в 03_insales_api.js

Используй mcp__filesystem__edit_file для изменений.
```

---

**Теперь у вас есть полный арсенал для отладки Google Apps Script! 🎯**
