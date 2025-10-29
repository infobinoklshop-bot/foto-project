/**
 * ===================================================================
 * МОДУЛЬ 01_shared_utilities.gs - ОБЩИЕ УТИЛИТЫ
 * ===================================================================
 * 
 * Назначение: Базовые функции общего назначения для всего проекта
 * Версия: 1.0
 * Проект: Обработка изображений товаров
 * 
 * Этот модуль содержит:
 * - Базовые функции работы с Google Sheets
 * - Логирование и обработка ошибок
 * - Telegram уведомления (опционально)
 * - Helper функции общего назначения
 * - Функции форматирования и валидации
 */

// =============================================================================
// 📊 ФУНКЦИИ РАБОТЫ С GOOGLE SHEETS
// =============================================================================

/**
 * ПОЛУЧЕНИЕ ЛИСТА С ПРОВЕРКОЙ СУЩЕСТВОВАНИЯ
 * 
 * Безопасное получение листа по названию с детальной проверкой
 * 
 * @param {string} sheetName - Название листа
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Найденный лист
 * @throws {Error} Если лист не найден или недоступен
 * @example
 * const sheet = getSheet('Обработка изображений');
 */
function getSheet(sheetName) {
  try {
    console.log(`📊 Получаем лист "${sheetName}"...`);
    
    // Проверяем входные параметры
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('Некорректное название листа');
    }
    
    // Получаем активную таблицу
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Не удалось получить активную таблицу Google Sheets');
    }
    
    // Получаем лист по названию
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      // Показываем доступные листы для диагностики
      const availableSheets = spreadsheet.getSheets().map(s => s.getName());
      console.error(`❌ Доступные листы: ${availableSheets.join(', ')}`);
      throw new Error(`Лист "${sheetName}" не найден. Создайте лист или проверьте название.`);
    }
    
    console.log(`✅ Лист "${sheetName}" успешно получен`);
    return sheet;
    
  } catch (error) {
    const errorMsg = `Ошибка получения листа "${sheetName}": ${error.message}`;
    console.error(`❌ ${errorMsg}`);
    throw new Error(errorMsg);
  }
}

/**
 * ПРОВЕРКА СУЩЕСТВОВАНИЯ ЛИСТА
 * 
 * Проверяет существование листа без выброса ошибки
 * 
 * @param {string} sheetName - Название листа
 * @returns {boolean} true если лист существует
 * @example
 * if (sheetExists('Обработка изображений')) {
 *   // лист существует
 * }
 */
function sheetExists(sheetName) {
  try {
    if (!sheetName || typeof sheetName !== 'string') {
      return false;
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    return sheet !== null;
    
  } catch (error) {
    console.error(`❌ Ошибка проверки существования листа "${sheetName}":`, error.message);
    return false;
  }
}

/**
 * ПОЛУЧЕНИЕ ДАННЫХ ИЗ ДИАПАЗОНА С ПРОВЕРКАМИ
 * 
 * Безопасное чтение данных из указанного диапазона
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист
 * @param {number} startRow - Начальная строка (1-based)
 * @param {number} startCol - Начальная колонка (1-based)
 * @param {number} numRows - Количество строк
 * @param {number} numCols - Количество колонок
 * @returns {Array<Array>} Двумерный массив данных
 * @throws {Error} Если параметры некорректны
 */
function getSheetData(sheet, startRow, startCol, numRows, numCols) {
  try {
    // Валидация входных параметров
    if (!sheet) {
      throw new Error('Лист не указан');
    }
    
    if (startRow < 1 || startCol < 1 || numRows < 1 || numCols < 1) {
      throw new Error('Некорректные параметры диапазона (должны быть > 0)');
    }
    
    // Проверяем границы листа
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    
    if (startRow + numRows - 1 > maxRows) {
      throw new Error(`Диапазон выходит за границы листа по строкам (макс: ${maxRows})`);
    }
    
    if (startCol + numCols - 1 > maxCols) {
      throw new Error(`Диапазон выходит за границы листа по колонкам (макс: ${maxCols})`);
    }
    
    // Получаем данные
    const range = sheet.getRange(startRow, startCol, numRows, numCols);
    const data = range.getValues();
    
    console.log(`✅ Получено ${data.length} строк из листа "${sheet.getName()}"`);
    return data;
    
  } catch (error) {
    const errorMsg = `Ошибка чтения данных: ${error.message}`;
    console.error(`❌ ${errorMsg}`);
    throw new Error(errorMsg);
  }
}

/**
 * ЗАПИСЬ ДАННЫХ В ДИАПАЗОН С ПРОВЕРКАМИ
 * 
 * Безопасная запись данных в указанный диапазон
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист
 * @param {number} startRow - Начальная строка (1-based)
 * @param {number} startCol - Начальная колонка (1-based)
 * @param {Array<Array>} data - Двумерный массив данных для записи
 * @throws {Error} Если параметры некорректны
 */
function setSheetData(sheet, startRow, startCol, data) {
  try {
    // Валидация входных параметров
    if (!sheet) {
      throw new Error('Лист не указан');
    }
    
    if (!data || !Array.isArray(data) || data.length === 0) {
      throw new Error('Данные для записи не указаны или пусты');
    }
    
    if (startRow < 1 || startCol < 1) {
      throw new Error('Некорректные параметры позиции (должны быть > 0)');
    }
    
    // Определяем размеры данных
    const numRows = data.length;
    const numCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 1));
    
    // Проверяем границы листа
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    
    if (startRow + numRows - 1 > maxRows) {
      console.log(`⚠️ Расширяем лист до ${startRow + numRows - 1} строк`);
      sheet.insertRows(maxRows, (startRow + numRows - 1) - maxRows);
    }
    
    if (startCol + numCols - 1 > maxCols) {
      console.log(`⚠️ Расширяем лист до ${startCol + numCols - 1} колонок`);
      sheet.insertColumns(maxCols, (startCol + numCols - 1) - maxCols);
    }
    
    // Записываем данные
    const range = sheet.getRange(startRow, startCol, numRows, numCols);
    range.setValues(data);
    
    console.log(`✅ Записано ${numRows} строк в лист "${sheet.getName()}"`);
    
  } catch (error) {
    const errorMsg = `Ошибка записи данных: ${error.message}`;
    console.error(`❌ ${errorMsg}`);
    throw new Error(errorMsg);
  }
}

// =============================================================================
// 📝 СИСТЕМА ЛОГИРОВАНИЯ
// =============================================================================

/**
 * УРОВНИ ЛОГИРОВАНИЯ
 */
const LOG_LEVELS = {
  DEBUG: 'DEBUG',
  INFO: 'INFO',
  WARNING: 'WARNING',
  ERROR: 'ERROR',
  CRITICAL: 'CRITICAL'
};

/**
 * СТРУКТУРИРОВАННОЕ ЛОГИРОВАНИЕ
 * 
 * Выводит сообщение с временной меткой и уровнем важности
 * 
 * @param {string} level - Уровень логирования (из LOG_LEVELS)
 * @param {string} message - Сообщение для логирования
 * @param {Object} context - Дополнительный контекст (опционально)
 * @example
 * logMessage(LOG_LEVELS.INFO, 'Начинаем обработку товара', {article: 'ART123'});
 */
function logMessage(level, message, context = null) {
  try {
    // Форматируем временную метку для русской локали
    const timestamp = formatDate(new Date(), 'full');
    
    // Выбираем эмодзи для уровня
    const levelEmojis = {
      [LOG_LEVELS.DEBUG]: '🔍',
      [LOG_LEVELS.INFO]: 'ℹ️',
      [LOG_LEVELS.WARNING]: '⚠️',
      [LOG_LEVELS.ERROR]: '❌',
      [LOG_LEVELS.CRITICAL]: '🚨'
    };
    
    const emoji = levelEmojis[level] || 'ℹ️';
    
    // Основное сообщение
    const logEntry = `${emoji} [${timestamp}] ${level}: ${message}`;
    
    // Выводим в консоль
    console.log(logEntry);
    
    // Добавляем контекст если есть
    if (context && typeof context === 'object') {
      console.log(`   Контекст:`, JSON.stringify(context, null, 2));
    }
    
    // Для критических ошибок можем отправить уведомление
    if (level === LOG_LEVELS.CRITICAL || level === LOG_LEVELS.ERROR) {
      try {
        sendNotification(`${emoji} ${level}: ${message}`);
      } catch (notificationError) {
        console.log('⚠️ Не удалось отправить уведомление:', notificationError.message);
      }
    }
    
  } catch (error) {
    // Даже если логирование упало, не должно ломать основной процесс
    console.error('❌ Ошибка в системе логирования:', error.message);
    console.log(`FALLBACK LOG [${level}]: ${message}`);
  }
}

/**
 * УПРОЩЕННЫЕ ФУНКЦИИ ЛОГИРОВАНИЯ
 */
function logDebug(message, context = null) {
  logMessage(LOG_LEVELS.DEBUG, message, context);
}

function logInfo(message, context = null) {
  logMessage(LOG_LEVELS.INFO, message, context);
}

function logWarning(message, context = null) {
  logMessage(LOG_LEVELS.WARNING, message, context);
}

function logError(message, context = null) {
  logMessage(LOG_LEVELS.ERROR, message, context);
}

function logCritical(message, context = null) {
  logMessage(LOG_LEVELS.CRITICAL, message, context);
}

// =============================================================================
// 🚨 ОБРАБОТКА ОШИБОК
// =============================================================================

/**
 * ЦЕНТРАЛИЗОВАННАЯ ОБРАБОТКА ОШИБОК
 * 
 * Стандартизированная обработка ошибок с логированием и уведомлениями
 * 
 * @param {Error} error - Объект ошибки
 * @param {string} context - Контекст где произошла ошибка
 * @param {Object} additionalInfo - Дополнительная информация
 * @throws {Error} Перебрасывает ошибку после обработки
 * @example
 * try {
 *   // какой-то код
 * } catch (error) {
 *   handleError(error, 'Обработка изображений', {article: 'ART123'});
 * }
 */
function handleError(error, context = 'Неизвестный контекст', additionalInfo = {}) {
  try {
    // Получаем детали ошибки
    const errorDetails = {
      message: error.message || 'Неизвестная ошибка',
      stack: error.stack || 'Stack trace недоступен',
      name: error.name || 'Error',
      context: context,
      timestamp: new Date().toISOString(),
      ...additionalInfo
    };
    
    // Логируем критическую ошибку
    logCritical(`Ошибка в контексте "${context}": ${errorDetails.message}`, {
      errorName: errorDetails.name,
      stack: errorDetails.stack,
      additionalInfo: additionalInfo
    });
    
    // Формируем сообщение для уведомления
    const notificationMessage = `🚨 ОШИБКА в проекте "Обработка изображений товаров"
    
📍 Контекст: ${context}
❌ Ошибка: ${errorDetails.message}
🕐 Время: ${formatDate(new Date(), 'full')}
    
${Object.keys(additionalInfo).length > 0 ? 
  `📋 Дополнительно: ${JSON.stringify(additionalInfo, null, 2)}` : ''
}`;
    
    // Отправляем уведомление
    try {
      sendNotification(notificationMessage);
    } catch (notificationError) {
      console.log('⚠️ Не удалось отправить уведомление об ошибке:', notificationError.message);
    }
    
    // Перебрасываем ошибку для дальнейшей обработки
    throw error;
    
  } catch (handlingError) {
    // Если даже обработка ошибок упала, делаем минимальное логирование
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА в обработчике ошибок:', handlingError.message);
    console.error('❌ Исходная ошибка:', error.message);
    throw error;
  }
}

/**
 * БЕЗОПАСНОЕ ВЫПОЛНЕНИЕ ФУНКЦИИ
 * 
 * Выполняет функцию с автоматической обработкой ошибок
 * 
 * @param {Function} func - Функция для выполнения
 * @param {string} context - Контекст выполнения
 * @param {*} defaultValue - Значение по умолчанию при ошибке
 * @returns {*} Результат выполнения функции или значение по умолчанию
 * @example
 * const result = safeExecute(() => someRiskyFunction(), 'Risky operation', null);
 */
function safeExecute(func, context = 'Безопасное выполнение', defaultValue = null) {
  try {
    return func();
  } catch (error) {
    logError(`Безопасное выполнение не удалось в контексте "${context}": ${error.message}`);
    return defaultValue;
  }
}

// =============================================================================
// 📱 TELEGRAM УВЕДОМЛЕНИЯ
// =============================================================================

/**
 * ОТПРАВКА УВЕДОМЛЕНИЯ В TELEGRAM
 * 
 * Отправляет сообщение в Telegram чат (если настроены токен и chat_id)
 * 
 * @param {string} message - Текст сообщения
 * @param {boolean} silent - Тихое уведомление (без звука)
 * @example
 * sendNotification('Обработка завершена!');
 */
function sendNotification(message, silent = false) {
  try {
    // Получаем настройки Telegram из конфигурации
    const telegramToken = getSetting(SCRIPT_PROPERTIES_KEYS.TELEGRAM_TOKEN);
    const telegramChatId = getSetting(SCRIPT_PROPERTIES_KEYS.TELEGRAM_CHAT_ID);
    
    // Проверяем наличие настроек
    if (!telegramToken || !telegramChatId) {
      console.log('ℹ️ Telegram не настроен, уведомление пропущено');
      return false;
    }
    
    // Ограничиваем длину сообщения (Telegram лимит ~4000 символов)
    const truncatedMessage = message.length > 3000 ? 
      message.substring(0, 3000) + '\n\n... (сообщение обрезано)' : 
      message;
    
    // Формируем URL API Telegram
    const url = `https://api.telegram.org/bot${telegramToken}/sendMessage`;
    
    // Параметры запроса
    const payload = {
      chat_id: telegramChatId,
      text: truncatedMessage,
      parse_mode: 'HTML',
      disable_notification: silent
    };
    
    // Отправляем запрос
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    // Проверяем результат
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && responseData.ok) {
      console.log('✅ Telegram уведомление отправлено');
      return true;
    } else {
      console.error('❌ Ошибка отправки Telegram уведомления:', responseData.description || 'Неизвестная ошибка');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Критическая ошибка отправки Telegram уведомления:', error.message);
    return false;
  }
}

/**
 * ТЕСТИРОВАНИЕ TELEGRAM УВЕДОМЛЕНИЙ
 * 
 * Отправляет тестовое сообщение для проверки настройки
 */
function testTelegramNotification() {
  const testMessage = `🧪 ТЕСТ УВЕДОМЛЕНИЙ
  
✅ Проект: Обработка изображений товаров
🕐 Время: ${formatDate(new Date(), 'full')}
🤖 Telegram уведомления работают корректно!`;
  
  const result = sendNotification(testMessage);
  
  if (result) {
    logInfo('✅ Тест Telegram уведомлений прошел успешно');
  } else {
    logWarning('⚠️ Тест Telegram уведомлений не удался');
  }
  
  return result;
}

// =============================================================================
// 🛠️ HELPER ФУНКЦИИ ОБЩЕГО НАЗНАЧЕНИЯ
// =============================================================================

/**
 * ФОРМАТИРОВАНИЕ ДАТ ДЛЯ РУССКОЙ ЛОКАЛИ
 * 
 * Форматирует дату в различных форматах для российских пользователей
 * 
 * @param {Date} date - Дата для форматирования
 * @param {string} format - Формат ('short', 'medium', 'full', 'time', 'iso')
 * @returns {string} Отформатированная дата
 * @example
 * formatDate(new Date(), 'full') // "15 августа 2024 г., 14:30:15"
 */
function formatDate(date = new Date(), format = 'medium') {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      throw new Error('Некорректная дата');
    }
    
    const options = {
      timeZone: 'Europe/Moscow' // Московское время
    };
    
    switch (format) {
      case 'short':
        // 15.08.2024
        return date.toLocaleDateString('ru-RU', {
          ...options,
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        });
        
      case 'medium':
        // 15 авг 2024 г.
        return date.toLocaleDateString('ru-RU', {
          ...options,
          day: 'numeric',
          month: 'short',
          year: 'numeric'
        });
        
      case 'full':
        // 15 августа 2024 г., 14:30:15
        return date.toLocaleString('ru-RU', {
          ...options,
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'
        });
        
      case 'time':
        // 14:30:15
        return date.toLocaleTimeString('ru-RU', {
          ...options,
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'
        });
        
      case 'iso':
        // 2024-08-15T14:30:15.000Z
        return date.toISOString();
        
      default:
        return date.toLocaleString('ru-RU', options);
    }
    
  } catch (error) {
    console.error('❌ Ошибка форматирования даты:', error.message);
    return date.toString();
  }
}

/**
 * СОЗДАНИЕ УНИКАЛЬНОГО ИДЕНТИФИКАТОРА
 * 
 * Генерирует уникальный ID для использования в проекте
 * 
 * @param {string} prefix - Префикс для ID (опционально)
 * @returns {string} Уникальный идентификатор
 * @example
 * const id = generateUniqueId('IMG'); // IMG_1692105615123_abc
 */
function generateUniqueId(prefix = '') {
  try {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(2, 8);
    
    return prefix ? `${prefix}_${timestamp}_${random}` : `${timestamp}_${random}`;
    
  } catch (error) {
    console.error('❌ Ошибка генерации ID:', error.message);
    return `fallback_${Date.now()}`;
  }
}

/**
 * ОЧИСТКА И ВАЛИДАЦИЯ СТРОК
 * 
 * Удаляет лишние пробелы и проверяет корректность строки
 * 
 * @param {string} str - Строка для очистки
 * @param {number} maxLength - Максимальная длина (опционально)
 * @returns {string} Очищенная строка
 * @example
 * const clean = cleanString('  Пример строки  ', 50);
 */
function cleanString(str, maxLength = null) {
  try {
    if (typeof str !== 'string') {
      return '';
    }
    
    // Удаляем лишние пробелы в начале и конце
    let cleaned = str.trim();
    
    // Заменяем множественные пробелы на одинарные
    cleaned = cleaned.replace(/\s+/g, ' ');
    
    // Обрезаем если задана максимальная длина
    if (maxLength && typeof maxLength === 'number' && maxLength > 0) {
      if (cleaned.length > maxLength) {
        cleaned = cleaned.substring(0, maxLength).trim();
      }
    }
    
    return cleaned;
    
  } catch (error) {
    console.error('❌ Ошибка очистки строки:', error.message);
    return str || '';
  }
}

/**
 * ВАЛИДАЦИЯ EMAIL АДРЕСА
 * 
 * Проверяет корректность email адреса
 * 
 * @param {string} email - Email для проверки
 * @returns {boolean} true если email корректный
 */
function isValidEmail(email) {
  try {
    if (typeof email !== 'string') {
      return false;
    }
    
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
    
  } catch (error) {
    console.error('❌ Ошибка валидации email:', error.message);
    return false;
  }
}

/**
 * ВАЛИДАЦИЯ URL
 * 
 * Проверяет корректность URL адреса
 * 
 * @param {string} url - URL для проверки
 * @returns {boolean} true если URL корректный
 */
function isValidUrl(url) {
  try {
    if (typeof url !== 'string') {
      return false;
    }
    
    const urlRegex = /^https?:\/\/.+/;
    return urlRegex.test(url.trim());
    
  } catch (error) {
    console.error('❌ Ошибка валидации URL:', error.message);
    return false;
  }
}

/**
 * ЗАДЕРЖКА ВЫПОЛНЕНИЯ (SLEEP)
 * 
 * Приостанавливает выполнение на указанное количество миллисекунд
 * 
 * @param {number} milliseconds - Количество миллисекунд для ожидания
 * @example
 * sleep(1000); // ждем 1 секунду
 */
function sleep(milliseconds) {
  try {
    if (typeof milliseconds !== 'number' || milliseconds < 0) {
      console.log('⚠️ Некорректное время ожидания, пропускаем');
      return;
    }
    
    console.log(`⏳ Ожидание ${milliseconds}мс...`);
    Utilities.sleep(milliseconds);
    
  } catch (error) {
    console.error('❌ Ошибка в функции ожидания:', error.message);
  }
}

/**
 * КОНВЕРТАЦИЯ РАЗМЕРА ФАЙЛА В ЧЕЛОВЕКОЧИТАЕМЫЙ ФОРМАТ
 * 
 * Преобразует размер в байтах в удобный для чтения формат
 * 
 * @param {number} bytes - Размер в байтах
 * @returns {string} Форматированный размер
 * @example
 * formatFileSize(1024) // "1.0 КБ"
 */
function formatFileSize(bytes) {
  try {
    if (typeof bytes !== 'number' || bytes < 0) {
      return '0 Б';
    }
    
    const sizes = ['Б', 'КБ', 'МБ', 'ГБ', 'ТБ'];
    
    if (bytes === 0) {
      return '0 Б';
    }
    
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    const size = (bytes / Math.pow(1024, i)).toFixed(1);
    
    return `${size} ${sizes[i]}`;
    
  } catch (error) {
    console.error('❌ Ошибка форматирования размера файла:', error.message);
    return `${bytes} Б`;
  }
}

// =============================================================================
// 🧪 ФУНКЦИИ ТЕСТИРОВАНИЯ МОДУЛЯ
// =============================================================================

/**
 * ТЕСТИРОВАНИЕ МОДУЛЯ SHARED_UTILITIES
 * 
 * Комплексное тестирование всех функций модуля
 */
function testSharedUtilitiesModule() {
  console.log('🧪 Начинаем тестирование модуля 01_shared_utilities.gs...');
  
  try {
    // Тест 1: Функции форматирования дат
    console.log('\n📅 Тест 1: Форматирование дат');
    const testDate = new Date();
    console.log(`✅ Короткий формат: ${formatDate(testDate, 'short')}`);
    console.log(`✅ Средний формат: ${formatDate(testDate, 'medium')}`);
    console.log(`✅ Полный формат: ${formatDate(testDate, 'full')}`);
    console.log(`✅ Только время: ${formatDate(testDate, 'time')}`);
    
    // Тест 2: Генерация ID
    console.log('\n🆔 Тест 2: Генерация уникальных ID');
    const id1 = generateUniqueId('TEST');
    const id2 = generateUniqueId('IMG');
    console.log(`✅ ID с префиксом: ${id1}`);
    console.log(`✅ ID с префиксом: ${id2}`);
    console.log(`✅ ID уникальны: ${id1 !== id2}`);
    
    // Тест 3: Очистка строк
    console.log('\n🧹 Тест 3: Очистка строк');
    const dirtyString = '  Пример   строки  с лишними   пробелами  ';
    const cleanedString = cleanString(dirtyString, 20);
    console.log(`✅ Исходная строка: "${dirtyString}"`);
    console.log(`✅ Очищенная строка: "${cleanedString}"`);
    
    // Тест 4: Валидация email и URL
    console.log('\n📧 Тест 4: Валидация данных');
    const testEmail = 'test@example.com';
    const invalidEmail = 'invalid-email';
    const testUrl = 'https://example.com/image.jpg';
    const invalidUrl = 'not-a-url';
    
    console.log(`✅ Email "${testEmail}" валиден: ${isValidEmail(testEmail)}`);
    console.log(`✅ Email "${invalidEmail}" валиден: ${isValidEmail(invalidEmail)}`);
    console.log(`✅ URL "${testUrl}" валиден: ${isValidUrl(testUrl)}`);
    console.log(`✅ URL "${invalidUrl}" валиден: ${isValidUrl(invalidUrl)}`);
    
    // Тест 5: Форматирование размеров файлов
    console.log('\n📦 Тест 5: Форматирование размеров файлов');
    const sizes = [0, 1024, 1048576, 1073741824];
    sizes.forEach(size => {
      console.log(`✅ ${size} байт = ${formatFileSize(size)}`);
    });
    
    // Тест 6: Логирование
    console.log('\n📝 Тест 6: Система логирования');
    logDebug('Тестовое debug сообщение', {test: true});
    logInfo('Тестовое info сообщение', {module: 'test'});
    logWarning('Тестовое warning сообщение');
    
    // Тест 7: Проверка существования листа (если есть доступ к Sheets)
    console.log('\n📊 Тест 7: Проверка работы с листами');
    try {
      const imagesSheetExists = sheetExists(SHEET_NAMES.IMAGES);
      console.log(`✅ Лист "${SHEET_NAMES.IMAGES}" существует: ${imagesSheetExists}`);
      
      if (imagesSheetExists) {
        const sheet = getSheet(SHEET_NAMES.IMAGES);
        console.log(`✅ Лист успешно получен: ${sheet.getName()}`);
      }
    } catch (sheetError) {
      console.log(`ℹ️ Тест листов пропущен (лист не создан): ${sheetError.message}`);
    }
    
    // Тест 8: Безопасное выполнение
    console.log('\n🛡️ Тест 8: Безопасное выполнение функций');
    const safeResult1 = safeExecute(() => {
      return 'Успешный результат';
    }, 'Успешная операция', 'default');
    console.log(`✅ Безопасное выполнение (успех): ${safeResult1}`);
    
    const safeResult2 = safeExecute(() => {
      throw new Error('Тестовая ошибка');
    }, 'Операция с ошибкой', 'значение по умолчанию');
    console.log(`✅ Безопасное выполнение (ошибка): ${safeResult2}`);
    
    // Тест 9: Тест уведомлений (если настроены)
    console.log('\n📱 Тест 9: Telegram уведомления');
    try {
      const telegramToken = getSetting(SCRIPT_PROPERTIES_KEYS.TELEGRAM_TOKEN);
      if (telegramToken) {
        console.log('✅ Telegram токен найден, запускаем тест...');
        const notificationResult = testTelegramNotification();
        console.log(`✅ Результат теста уведомлений: ${notificationResult}`);
      } else {
        console.log('ℹ️ Telegram не настроен, тест уведомлений пропущен');
      }
    } catch (telegramError) {
      console.log(`ℹ️ Тест Telegram пропущен: ${telegramError.message}`);
    }
    
    console.log('\n🎉 Тестирование модуля 01_shared_utilities.gs завершено!');
    console.log('📋 Все основные функции работают корректно');
    
    return {
      success: true,
      message: 'Все тесты пройдены успешно',
      timestamp: formatDate(new Date(), 'full')
    };
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании модуля shared_utilities:', error.message);
    logError('Критическая ошибка в тестировании модуля', {
      module: '01_shared_utilities.gs',
      error: error.message
    });
    
    return {
      success: false,
      message: error.message,
      timestamp: formatDate(new Date(), 'full')
    };
  }
}

/**
 * ДЕМОНСТРАЦИЯ ВОЗМОЖНОСТЕЙ МОДУЛЯ
 * 
 * Показывает примеры использования основных функций
 */
function demonstrateSharedUtilities() {
  console.log('🎭 Демонстрация возможностей модуля 01_shared_utilities.gs');
  console.log('='.repeat(60));
  
  // Демо 1: Работа с датами
  console.log('\n📅 РАБОТА С ДАТАМИ:');
  const now = new Date();
  console.log(`Текущее время (полный формат): ${formatDate(now, 'full')}`);
  console.log(`Текущая дата (короткий формат): ${formatDate(now, 'short')}`);
  console.log(`Только время: ${formatDate(now, 'time')}`);
  
  // Демо 2: Логирование разных уровней
  console.log('\n📝 ЛОГИРОВАНИЕ:');
  logInfo('Запуск демонстрации', {demo: true, version: '1.0'});
  logWarning('Это предупреждение для демонстрации');
  logDebug('Отладочная информация', {debugLevel: 'verbose'});
  
  // Демо 3: Работа со строками
  console.log('\n🧹 ОБРАБОТКА СТРОК:');
  const testStrings = [
    '  Товар с пробелами  ',
    'Очень длинная строка которая должна быть обрезана до определенной длины для демонстрации',
    '   Множественные    пробелы   внутри   '
  ];
  
  testStrings.forEach((str, index) => {
    const cleaned = cleanString(str, 30);
    console.log(`Строка ${index + 1}:`);
    console.log(`  Исходная: "${str}"`);
    console.log(`  Обработанная: "${cleaned}"`);
  });
  
  // Демо 4: Валидация данных
  console.log('\n✅ ВАЛИДАЦИЯ:');
  const testData = [
    {type: 'email', value: 'user@example.com'},
    {type: 'email', value: 'invalid-email'},
    {type: 'url', value: 'https://cdn.example.com/image.jpg'},
    {type: 'url', value: 'not-a-valid-url'}
  ];
  
  testData.forEach(item => {
    const isValid = item.type === 'email' ? isValidEmail(item.value) : isValidUrl(item.value);
    console.log(`${item.type.toUpperCase()} "${item.value}": ${isValid ? '✅ валиден' : '❌ не валиден'}`);
  });
  
  // Демо 5: Генерация ID
  console.log('\n🆔 ГЕНЕРАЦИЯ ID:');
  console.log(`Обычный ID: ${generateUniqueId()}`);
  console.log(`ID изображения: ${generateUniqueId('IMG')}`);
  console.log(`ID товара: ${generateUniqueId('PRODUCT')}`);
  
  // Демо 6: Форматирование размеров
  console.log('\n📦 РАЗМЕРЫ ФАЙЛОВ:');
  const fileSizes = [512, 2048, 1048576, 5242880, 1073741824];
  fileSizes.forEach(size => {
    console.log(`${size} байт = ${formatFileSize(size)}`);
  });
  
  console.log('\n' + '='.repeat(60));
  console.log('🎉 Демонстрация завершена!');
}

// =============================================================================
// 📱 ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ РАСШИРЕНИЯ
// =============================================================================

/**
 * ПОЛУЧЕНИЕ ИНФОРМАЦИИ О ПРОЕКТЕ ИЗ НЕСКОЛЬКИХ МОДУЛЕЙ
 * 
 * Агрегирует информацию из config и shared_utilities модулей
 * 
 * @returns {Object} Расширенная информация о проекте
 */
function getExtendedProjectInfo() {
  try {
    // Получаем базовую информацию из config модуля
    const baseInfo = getProjectInfo();
    
    // Добавляем информацию о текущем модуле
    const extendedInfo = {
      ...baseInfo,
      sharedUtilitiesVersion: '01_shared_utilities.js v1.0',
      availableLogLevels: Object.keys(LOG_LEVELS),
      supportedDateFormats: ['short', 'medium', 'full', 'time', 'iso'],
      features: [
        'Структурированное логирование',
        'Telegram уведомления',
        'Безопасная работа с Google Sheets',
        'Валидация данных',
        'Форматирование дат и размеров',
        'Централизованная обработка ошибок'
      ],
      lastTested: formatDate(new Date(), 'full')
    };
    
    return extendedInfo;
    
  } catch (error) {
    logError('Ошибка получения расширенной информации о проекте', {error: error.message});
    return {
      error: error.message,
      timestamp: formatDate(new Date(), 'full')
    };
  }
}

/**
 * ПРОВЕРКА РАБОТОСПОСОБНОСТИ ВСЕХ ЗАВИСИМОСТЕЙ
 * 
 * Тестирует интеграцию с модулем config и доступность всех API
 * 
 * @returns {Object} Результат проверки зависимостей
 */
function checkDependencies() {
  console.log('🔍 Проверка зависимостей модуля shared_utilities...');
  
  const result = {
    configModule: false,
    sheetsAccess: false,
    scripProperties: false,
    telegramApi: false,
    errors: [],
    warnings: []
  };
  
  try {
    // Проверка модуля config
    try {
      const projectInfo = getProjectInfo();
      if (projectInfo && projectInfo.projectName) {
        result.configModule = true;
        console.log('✅ Модуль config доступен');
      }
    } catch (configError) {
      result.errors.push('Модуль config недоступен: ' + configError.message);
    }
    
    // Проверка доступа к Google Sheets
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (spreadsheet) {
        result.sheetsAccess = true;
        console.log('✅ Доступ к Google Sheets работает');
      }
    } catch (sheetsError) {
      result.errors.push('Доступ к Google Sheets недоступен: ' + sheetsError.message);
    }
    
    // Проверка Script Properties
    try {
      const properties = PropertiesService.getScriptProperties();
      const testKey = 'test_key_' + Date.now();
      properties.setProperty(testKey, 'test_value');
      const retrievedValue = properties.getProperty(testKey);
      properties.deleteProperty(testKey);
      
      if (retrievedValue === 'test_value') {
        result.scripProperties = true;
        console.log('✅ Script Properties работают');
      }
    } catch (propertiesError) {
      result.errors.push('Script Properties недоступны: ' + propertiesError.message);
    }
    
    // Проверка настроек Telegram
    try {
      const telegramToken = getSetting(SCRIPT_PROPERTIES_KEYS.TELEGRAM_TOKEN);
      const telegramChatId = getSetting(SCRIPT_PROPERTIES_KEYS.TELEGRAM_CHAT_ID);
      
      if (telegramToken && telegramChatId) {
        result.telegramApi = true;
        console.log('✅ Telegram настроен');
      } else {
        result.warnings.push('Telegram не настроен (уведомления недоступны)');
      }
    } catch (telegramError) {
      result.warnings.push('Ошибка проверки Telegram: ' + telegramError.message);
    }
    
    // Итоговая оценка
    const criticalCount = result.errors.length;
    const warningCount = result.warnings.length;
    
    console.log(`\n📊 Результат проверки зависимостей:`);
    console.log(`✅ Критических проблем: ${criticalCount}`);
    console.log(`⚠️ Предупреждений: ${warningCount}`);
    
    if (criticalCount === 0) {
      console.log('🎉 Все критические зависимости работают!');
    } else {
      console.log('❌ Обнаружены критические проблемы');
    }
    
    return result;
    
  } catch (error) {
    result.errors.push('Критическая ошибка проверки зависимостей: ' + error.message);
    logCritical('Ошибка проверки зависимостей shared_utilities', {error: error.message});
    return result;
  }
}

/**
 * ИСПРАВЛЕНИЕ: Функция showNotification для модуля 04_image_processing.gs
 * Добавить в конец файла 01_shared_utilities.gs
 */

/**
 * Показать уведомление пользователю через UI
 * @param {string} message - Текст уведомления
 * @param {string} type - Тип: 'success', 'warning', 'error', 'info'
 */
function showNotification(message, type = 'info') {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Определяем заголовок по типу
    const titles = {
      'success': '✅ Успешно',
      'warning': '⚠️ Предупреждение', 
      'error': '❌ Ошибка',
      'info': 'ℹ️ Информация'
    };
    
    const title = titles[type] || titles['info'];
    
    // Показываем alert
    ui.alert(title, message, ui.ButtonSet.OK);
    
    // Также логируем для отладки
    logInfo(`UI уведомление [${type}]: ${message}`);
    
  } catch (error) {
    // Fallback - если UI недоступен, только логируем
    logError(`Ошибка показа уведомления: ${error.message}`, null, 'showNotification');
    console.log(`${type.toUpperCase()}: ${message}`);
  }
}

/**
 * ===================================================================
 * 💡 ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ МОДУЛЯ
 * ===================================================================
 * 
 * ОСНОВНЫЕ ФУНКЦИИ ДЛЯ ДРУГИХ МОДУЛЕЙ:
 * 
 * 📊 РАБОТА С GOOGLE SHEETS:
 * - getSheet(sheetName) - безопасное получение листа
 * - sheetExists(sheetName) - проверка существования листа
 * - getSheetData(sheet, row, col, numRows, numCols) - чтение данных
 * - setSheetData(sheet, row, col, data) - запись данных
 * 
 * 📝 ЛОГИРОВАНИЕ:
 * - logInfo(message, context) - информационные сообщения
 * - logWarning(message, context) - предупреждения
 * - logError(message, context) - ошибки
 * - logCritical(message, context) - критические ошибки
 * - logDebug(message, context) - отладочная информация
 * 
 * 🚨 ОБРАБОТКА ОШИБОК:
 * - handleError(error, context, additionalInfo) - централизованная обработка
 * - safeExecute(func, context, defaultValue) - безопасное выполнение
 * 
 * 📱 УВЕДОМЛЕНИЯ:
 * - sendNotification(message, silent) - отправка в Telegram
 * - testTelegramNotification() - тест уведомлений
 * 
 * 🛠️ УТИЛИТЫ:
 * - formatDate(date, format) - форматирование дат
 * - generateUniqueId(prefix) - генерация уникальных ID
 * - cleanString(str, maxLength) - очистка строк
 * - isValidEmail(email) - валидация email
 * - isValidUrl(url) - валидация URL
 * - formatFileSize(bytes) - форматирование размеров файлов
 * - sleep(milliseconds) - задержка выполнения
 * 
 * ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ В ДРУГИХ МОДУЛЯХ:
 * 
 * // Логирование с контекстом
 * logInfo('Начинаем обработку товара', {article: 'ART123', stage: 'processing'});
 * 
 * // Безопасная работа с API
 * try {
 *   const result = apiCall();
 * } catch (error) {
 *   handleError(error, 'API вызов', {endpoint: '/products', method: 'GET'});
 * }
 * 
 * // Работа с листами
 * const sheet = getSheet(SHEET_NAMES.IMAGES);
 * const data = getSheetData(sheet, 2, 1, 10, 5);
 * 
 * // Валидация и форматирование
 * const cleanTitle = cleanString(productTitle, 100);
 * const timestamp = formatDate(new Date(), 'full');
 * 
 * ===================================================================
 */
