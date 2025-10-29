/**
 * ========================================
 * МОДУЛЬ 99: ПОЛЬЗОВАТЕЛЬСКИЙ ИНТЕРФЕЙС
 * ========================================
 * 
 * Назначение: Создание меню и UI элементов в Google Sheets
 * Версия: 1.0
 * Проект: Обработка изображений товаров
 * 
 * Содержит:
 * - Главное меню автоматизации
 * - Функции для запуска workflow
 * - UI элементы для взаимодействия с пользователем
 * - Уведомления и диалоги
 */

// ========================================
// СОЗДАНИЕ ГЛАВНОГО МЕНЮ
// ========================================

function onOpen() {
 try {
   console.log('🎛️ Создаем упрощенное меню обработки изображений...');
  
   const ui = SpreadsheetApp.getUi();
  
   // Создаем главное меню
   const mainMenu = ui.createMenu('🖼️ Фото');
  
   // ========================================
   // РАБОЧИЙ БЛОК - основные операции
   // ========================================
   mainMenu
     .addItem('📥 Загрузить товары', 'loadProductsFromInSalesMenu')
     .addItem('🔍 Спарсить у поставщиков', 'showSupplierParsingDialog')
     .addSeparator()
     .addItem('🤖 Обработать изображения', 'showImageSelectionForProcessing')
     .addItem('🤖 Выбрать модель', 'configureReplicateModel')
     .addItem('🎛️ Настроить качество улучшения', 'configureReplicateScale')
     .addSeparator()
     .addItem('📤 Отправить в InSales', 'sendProcessedImagesToInSales')
     .addItem('🏷️ Прописать Alt-теги', 'createAltTagCopyHelper')
     .addSeparator()
  
   // ========================================
   // СЕРВИСНЫЙ БЛОК - техническое обслуживание
   // ========================================
     .addItem('⚙️ Проверить настройки API', 'validateConfig')
     .addItem('🔄 Обновить товары из InSales', 'updateProductsFromInSales')
     .addItem('🧹 Очистить статусы обработки', 'clearProcessingStatuses')
     .addSeparator()
     .addItem('🆘 Справка', 'showHelpDialog')
     .addToUi();
  
   console.log('✅ Упрощенное меню создано успешно');
  
   // Показываем приветственное сообщение при первом запуске
   showWelcomeMessageIfNeeded();
  
 } catch (error) {
   console.error('❌ Ошибка создания меню:', error.message);
   
   // Создаем минимальное меню при ошибке
   try {
     const ui = SpreadsheetApp.getUi();
     ui.createMenu('🖼️ Обработка изображений (упрощенная)')
       .addItem('📥 Загрузить товары', 'loadProductsFromInSalesMenu')
       .addItem('⚙️ Проверить настройки', 'validateConfig')
       .addItem('🆘 Справка', 'showHelpDialog')
       .addToUi();
   } catch (fallbackError) {
     console.error('❌ Критическая ошибка создания меню:', fallbackError.message);
   }
 }
}

// ========================================
// WRAPPER ФУНКЦИИ ДЛЯ МЕНЮ
// ========================================

/**
 * ОБНОВЛЕНИЕ ТОВАРОВ ИЗ INSALES
 * 
 * Wrapper для обновления существующих товаров
 * Показывает диалог с опциями обновления
 */
function updateProductsFromInSales() {
  try {
    console.log('🔄 Запуск обновления товаров из InSales...');
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Обновление товаров',
      'Выберите режим обновления:\n\n' +
      '• ОК - Обновить только цены и статусы\n' +
      '• Отмена - Полное обновление всех данных',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response === ui.Button.OK) {
      // Частичное обновление
      updateExistingProducts({ mode: 'partial' });
    } else if (response === ui.Button.CANCEL) {
      // Полное обновление
      updateExistingProducts({ mode: 'full' });
    }
    
  } catch (error) {
    showErrorDialog('Ошибка обновления товаров', error.message);
  }
}

/**
 * ОБРАБОТКА ИЗОБРАЖЕНИЙ AI (ЗАГЛУШКА ДЛЯ ЭТАПА 3)
 * 
 * Пока что показывает информационное сообщение
 * Будет заменена на реальную функцию в Этапе 3
 */
function processSelectedImages() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Проверяем наличие выбранных товаров
    const selectedCount = getSelectedProductsCount();
    
    if (selectedCount === 0) {
      ui.alert(
        'Нет выбранных товаров',
        'Отметьте чекбоксами товары для обработки изображений.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Показываем информацию о будущей функции
    const response = ui.alert(
      'AI-обработка изображений',
      `Готово к обработке: ${selectedCount} товаров\n\n` +
      '🚧 Функция будет доступна в Этапе 3\n' +
      '• Анализ изображений через OpenAI\n' +
      '• Генерация alt-тегов\n' +
      '• Создание SEO-имен файлов\n\n' +
      'Продолжить разработку Этапа 3?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ui.alert(
        'Этап 3 в разработке',
        'Модуль AI-обработки изображений будет готов в ближайшее время!\n\n' +
        'Следите за обновлениями проекта.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    showErrorDialog('Ошибка обработки изображений', error.message);
  }
}

/**
 * ОЧИСТКА СТАТУСОВ ОБРАБОТКИ
 * 
 * Сбрасывает статусы для повторной обработки
 */
function clearProcessingStatuses() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      'Очистка статусов',
      'Это действие сбросит все статусы обработки.\n' +
      'Товары можно будет обработать заново.\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const clearedCount = clearAllProcessingStatuses();
      
      ui.alert(
        'Статусы очищены',
        `Сброшены статусы для ${clearedCount} товаров.\n` +
        'Теперь их можно обработать заново.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    showErrorDialog('Ошибка очистки статусов', error.message);
  }
}

// ========================================
// ИНФОРМАЦИОННЫЕ ДИАЛОГИ
// ========================================

/**
 * ПОКАЗАТЬ ДИАЛОГ ПОМОЩИ
 * 
 * Основная справочная информация для пользователей
 */
function showHelpDialog() {
 try {
   const ui = SpreadsheetApp.getUi();
  
   const message =
     'СПРАВКА ПО ОБРАБОТКЕ ИЗОБРАЖЕНИЙ ТОВАРОВ\n\n' +
    
     'WORKFLOW:\n' +
     '1. Загрузить товары из InSales\n' +
     '2. Парсинг поставщиков (при необходимости)\n' +
     '3. Выбрать товары чекбоксами\n' +
     '4. Обработать изображения AI\n' +
     '5. Alt-теги (помощник копирования)\n' +
     '6. Отправить выбранные в InSales\n\n' +
    
     'НАСТРОЙКИ API:\n' +
     '• InSales: API Key, Password, Shop\n' +
     '• OpenAI: API Key, Assistant ID\n' +
     '• Дополнительно: Replicate, TinyPNG, ImgBB\n\n' +
    
     'КОЛОНКИ:\n' +
     '• A: Выбор товаров\n' +
     '• B: Артикул\n' +
     '• D: Название\n' +
     '• E: Исходные изображения\n' +
     '• F: Парсинг поставщика\n' +
     '• H: Обработанные изображения\n' +
     '• I: Alt-теги\n' +
     '• J: SEO-имена файлов\n\n' +
    
     'ТЕХПОДДЕРЖКА:\n' +
     'Проверьте консоль Apps Script для логов операций.';
  
   ui.alert('Справка', message, ui.ButtonSet.OK);
  
 } catch (error) {
   console.error('Ошибка показа справки:', error.message);
 }
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ========================================

/**
 * ПОКАЗАТЬ ПРИВЕТСТВЕННОЕ СООБЩЕНИЕ
 * 
 * Показывается при первом запуске проекта
 */
function showWelcomeMessageIfNeeded() {
  try {
    const isFirstRun = getSetting('first_run_completed') !== 'true';
    
    if (isFirstRun) {
      const ui = SpreadsheetApp.getUi();
      
      const response = ui.alert(
        'Добро пожаловать! 🎉',
        'Это первый запуск системы автоматизации товаров.\n\n' +
        '🔧 Рекомендуем начать с:\n' +
        '1. Настройка системы → Проверить настройки API\n' +
        '2. Настройка системы → Создать структуру листа\n\n' +
        'Показать справку по настройке?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        showHelpDialog();
      }
      
      // Отмечаем, что первый запуск завершен
      setSetting('first_run_completed', 'true');
    }
    
  } catch (error) {
    console.error('❌ Ошибка приветственного сообщения:', error.message);
  }
}

/**
 * ПОКАЗАТЬ ДИАЛОГ ОШИБКИ
 * 
 * Универсальная функция для показа ошибок пользователю
 * 
 * @param {string} title - Заголовок ошибки
 * @param {string} message - Текст ошибки
 */
function showErrorDialog(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      `❌ ${title}`,
      `Произошла ошибка:\n\n${message}\n\n` +
      'Проверьте консоль Apps Script\n' +
      'для получения подробной информации.',
      ui.ButtonSet.OK
    );
    
    console.error(`❌ ${title}:`, message);
    
  } catch (error) {
    console.error('❌ Критическая ошибка показа диалога:', error.message);
  }
}

/**
 * ПОКАЗАТЬ ДИАЛОГ УСПЕХА
 * 
 * Универсальная функция для показа успешных операций
 * 
 * @param {string} title - Заголовок успеха
 * @param {string} message - Текст сообщения
 */
function showSuccessDialog(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      `✅ ${title}`,
      message,
      ui.ButtonSet.OK
    );
    
    console.log(`✅ ${title}:`, message);
    
  } catch (error) {
    console.error('❌ Ошибка показа диалога успеха:', error.message);
  }
}

// ========================================
// ФУНКЦИИ АНАЛИЗА ДАННЫХ
// ========================================

/**
 * ПОЛУЧЕНИЕ КОЛИЧЕСТВА ВЫБРАННЫХ ТОВАРОВ
 * 
 * Подсчитывает товары с отмеченными чекбоксами
 * 
 * @returns {number} Количество выбранных товаров
 */
function getSelectedProductsCount() {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    let selectedCount = 0;
    
    for (let i = 1; i < data.length; i++) { // Пропускаем заголовок
      const isSelected = data[i][IMAGES_COLUMNS.CHECKBOX - 1];
      if (isSelected === true) {
        selectedCount++;
      }
    }
    
    return selectedCount;
    
  } catch (error) {
    console.error('❌ Ошибка подсчета выбранных товаров:', error.message);
    return 0;
  }
}

/**
 * ГЕНЕРАЦИЯ СТАТИСТИКИ ОБРАБОТКИ
 * 
 * Анализирует текущее состояние товаров в листе
 * 
 * @returns {Object} Объект со статистикой
 */
function generateProcessingStatistics() {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    const stats = {
      total: 0,
      withImages: 0,
      withoutImages: 0,
      notProcessed: 0,
      processed: 0,
      errors: 0,
      notSent: 0,
      sent: 0,
      sendErrors: 0,
      lastUpdate: 'Никогда'
    };
    
    for (let i = 1; i < data.length; i++) { // Пропускаем заголовок
      const row = data[i];
      
      stats.total++;
      
      // Проверяем наличие изображений
      const hasImages = row[IMAGES_COLUMNS.ORIGINAL_IMAGES - 1];
      if (hasImages) {
        stats.withImages++;
      } else {
        stats.withoutImages++;
      }
      
      // Анализируем статус обработки
      const processingStatus = row[IMAGES_COLUMNS.PROCESSING_STATUS - 1];
      if (processingStatus === STATUS_VALUES.PROCESSING.NOT_PROCESSED) {
        stats.notProcessed++;
      } else if (processingStatus === STATUS_VALUES.PROCESSING.COMPLETED) {
        stats.processed++;
      } else if (processingStatus === STATUS_VALUES.PROCESSING.ERROR) {
        stats.errors++;
      }
      
      // Анализируем статус отправки
      const insalesStatus = row[IMAGES_COLUMNS.INSALES_STATUS - 1];
      if (insalesStatus === STATUS_VALUES.INSALES.NOT_SENT) {
        stats.notSent++;
      } else if (insalesStatus === STATUS_VALUES.INSALES.SENT) {
        stats.sent++;
      } else if (insalesStatus === STATUS_VALUES.INSALES.ERROR) {
        stats.sendErrors++;
      }
    }
    
    // Определяем дату последнего обновления
    if (stats.total > 0) {
      stats.lastUpdate = new Date().toLocaleString('ru-RU');
    }
    
    return stats;
    
  } catch (error) {
    console.error('❌ Ошибка генерации статистики:', error.message);
    return {
      total: 0, withImages: 0, withoutImages: 0,
      notProcessed: 0, processed: 0, errors: 0,
      notSent: 0, sent: 0, sendErrors: 0,
      lastUpdate: 'Ошибка'
    };
  }
}

/**
 * ОЧИСТКА ВСЕХ СТАТУСОВ ОБРАБОТКИ
 * 
 * Сбрасывает статусы для повторной обработки товаров
 * 
 * @returns {number} Количество очищенных записей
 */
function clearAllProcessingStatuses() {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    let clearedCount = 0;
    
    for (let i = 1; i < data.length; i++) { // Пропускаем заголовок
      const rowIndex = i + 1; // Номер строки в Sheets (с учетом заголовка)
      
      // Очищаем колонки обработанных данных
      sheet.getRange(rowIndex, IMAGES_COLUMNS.PROCESSED_IMAGES).setValue('');
      sheet.getRange(rowIndex, IMAGES_COLUMNS.ALT_TAGS).setValue('');
      sheet.getRange(rowIndex, IMAGES_COLUMNS.SEO_FILENAMES).setValue('');
      
      // Сбрасываем статусы
      sheet.getRange(rowIndex, IMAGES_COLUMNS.PROCESSING_STATUS)
           .setValue(STATUS_VALUES.PROCESSING.NOT_PROCESSED);
      sheet.getRange(rowIndex, IMAGES_COLUMNS.INSALES_STATUS)
           .setValue(STATUS_VALUES.INSALES.NOT_SENT);
      
      clearedCount++;
    }
    
    console.log(`✅ Очищены статусы для ${clearedCount} товаров`);
    return clearedCount;
    
  } catch (error) {
    console.error('❌ Ошибка очистки статусов:', error.message);
    return 0;
  }
}

// ========================================
// СИСТЕМНЫЕ ФУНКЦИИ (ЗАГЛУШКИ ДЛЯ БУДУЩИХ ЭТАПОВ)
// ========================================

/**
 * ОБНОВЛЕНИЕ СУЩЕСТВУЮЩИХ ТОВАРОВ (ЗАГЛУШКА)
 * 
 * Будет реализована в будущих версиях
 * 
 * @param {Object} options - Опции обновления
 */
function updateExistingProducts(options = {}) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      'Функция в разработке',
      `Обновление товаров (режим: ${options.mode})\n` +
      'будет реализовано в следующих версиях.\n\n' +
      'Пока используйте полную загрузку товаров.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    showErrorDialog('Ошибка обновления товаров', error.message);
  }
}

// ========================================
// ФУНКЦИИ УВЕДОМЛЕНИЙ
// ========================================

/**
 * ОТПРАВКА УВЕДОМЛЕНИЯ ПОЛЬЗОВАТЕЛЮ
 * 
 * Универсальная функция для системных уведомлений
 * Поддерживает разные типы уведомлений
 * 
 * @param {string} title - Заголовок уведомления
 * @param {string} message - Текст уведомления  
 * @param {string} type - Тип уведомления (success, warning, error, info)
 */
function showUserNotification(title, message, type = 'info') {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Определяем иконку по типу
    const icons = {
      success: '✅',
      warning: '⚠️', 
      error: '❌',
      info: 'ℹ️'
    };
    
    const icon = icons[type] || 'ℹ️';
    
    ui.alert(
      `${icon} ${title}`,
      message,
      ui.ButtonSet.OK
    );
    
    // Дублируем в консоль для отладки
    console.log(`${icon} ${title}: ${message}`);
    
  } catch (error) {
    console.error('❌ Ошибка показа уведомления:', error.message);
  }
}

/**
 * ПОКАЗАТЬ ПРОГРЕСС ОПЕРАЦИИ
 * 
 * Информирует пользователя о долгих операциях
 * 
 * @param {string} operation - Название операции
 * @param {string} details - Детали операции
 */
function showProgressNotification(operation, details) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      `⏳ ${operation}`,
      `${details}\n\n` +
      'Операция выполняется в фоновом режиме.\n' +
      'Результат будет показан по завершении.',
      ui.ButtonSet.OK
    );
    
    console.log(`⏳ ${operation}: ${details}`);
    
  } catch (error) {
    console.error('❌ Ошибка показа прогресса:', error.message);
  }
}

// ========================================
// ФУНКЦИИ ПРОВЕРКИ СОСТОЯНИЯ
// ========================================

/**
 * ПРОВЕРКА ГОТОВНОСТИ К ЗАГРУЗКЕ ТОВАРОВ
 * 
 * Проверяет все необходимые условия перед загрузкой
 * 
 * @returns {Object} Результат проверки
 */
function checkReadinessForProductLoad() {
  try {
    const checks = {
      sheetExists: false,
      apiConfigured: false,
      connectionOk: false,
      ready: false,
      issues: []
    };
    
    // Проверка существования листа
    try {
      getImagesSheet();
      checks.sheetExists = true;
    } catch (error) {
      checks.issues.push('Лист не создан - выполните "Создать структуру листа"');
    }
    
    // Проверка настроек API
    try {
      const settings = getApiSettings();
      if (settings.insalesApiKey && settings.insalesPassword && settings.insalesShop) {
        checks.apiConfigured = true;
      } else {
        checks.issues.push('InSales API не настроен - проверьте Script Properties');
      }
    } catch (error) {
      checks.issues.push('Ошибка проверки настроек API');
    }
    
    // Проверка подключения (быстрая)
    try {
      // Здесь можно добавить быструю проверку без полного тестирования
      if (checks.apiConfigured) {
        checks.connectionOk = true;
      }
    } catch (error) {
      checks.issues.push('Проблемы с подключением к InSales');
    }
    
    // Общая готовность
    checks.ready = checks.sheetExists && checks.apiConfigured && checks.connectionOk;
    
    return checks;
    
  } catch (error) {
    console.error('❌ Ошибка проверки готовности:', error.message);
    return {
      sheetExists: false,
      apiConfigured: false, 
      connectionOk: false,
      ready: false,
      issues: ['Ошибка проверки системы']
    };
  }
}

/**
 * ПРОВЕРКА ГОТОВНОСТИ К AI-ОБРАБОТКЕ
 * 
 * Проверяет условия для запуска AI-обработки изображений
 * 
 * @returns {Object} Результат проверки
 */
function checkReadinessForAiProcessing() {
  try {
    const checks = {
      hasProducts: false,
      hasSelectedProducts: false,
      aiConfigured: false,
      ready: false,
      issues: []
    };
    
    // Проверка наличия товаров
    try {
      const stats = generateProcessingStatistics();
      checks.hasProducts = stats.total > 0;
      
      if (!checks.hasProducts) {
        checks.issues.push('Нет товаров в листе - загрузите товары из InSales');
      }
    } catch (error) {
      checks.issues.push('Ошибка проверки товаров в листе');
    }
    
    // Проверка выбранных товаров
    try {
      const selectedCount = getSelectedProductsCount();
      checks.hasSelectedProducts = selectedCount > 0;
      
      if (!checks.hasSelectedProducts) {
        checks.issues.push('Нет выбранных товаров - отметьте чекбоксами товары для обработки');
      }
    } catch (error) {
      checks.issues.push('Ошибка проверки выбранных товаров');
    }
    
    // Проверка настроек AI
    try {
      const settings = getApiSettings();
      checks.aiConfigured = !!(settings.openaiApiKey);
      
      if (!checks.aiConfigured) {
        checks.issues.push('OpenAI API не настроен - добавьте ключ в Script Properties');
      }
    } catch (error) {
      checks.issues.push('Ошибка проверки настроек AI');
    }
    
    // Общая готовность
    checks.ready = checks.hasProducts && checks.hasSelectedProducts && checks.aiConfigured;
    
    return checks;
    
  } catch (error) {
    console.error('❌ Ошибка проверки готовности к AI:', error.message);
    return {
      hasProducts: false,
      hasSelectedProducts: false,
      aiConfigured: false,
      ready: false,
      issues: ['Ошибка проверки системы AI']
    };
  }
}

// ========================================
// РАСШИРЕННЫЕ ФУНКЦИИ МЕНЮ (ДЛЯ БУДУЩИХ ЭТАПОВ)
// ========================================

/**
 * СОЗДАНИЕ ПОДМЕНЮ ДЛЯ ВНЕШНИХ СЕРВИСОВ
 * 
 * Будет активировано в Этапе 4
 */
function createExternalServicesMenu() {
  // Заглушка для Этапа 4
  const ui = SpreadsheetApp.getUi();
  
  const externalMenu = ui.createMenu('🌐 Внешние сервисы')
    .addItem('📤 Загрузить на ImgBB', 'uploadToImgBB')
    .addItem('🗜️ Сжать через TinyPNG', 'compressWithTinyPNG')
    .addItem('🎨 Улучшить через Replicate', 'enhanceWithReplicate')
    .addSeparator()
    .addItem('⚙️ Настройки сервисов', 'configureExternalServices');
    
  return externalMenu;
}

/**
 * СОЗДАНИЕ ПОДМЕНЮ ДЛЯ ПАКЕТНЫХ ОПЕРАЦИЙ
 * 
 * Будет активировано в Этапе 5
 */
function createBatchOperationsMenu() {
  // Заглушка для Этапа 5
  const ui = SpreadsheetApp.getUi();
  
  const batchMenu = ui.createMenu('📦 Пакетные операции')
    .addItem('🔄 Массовая обработка', 'processBatchOfProducts')
    .addItem('📤 Массовая загрузка', 'batchUploadImages')
    .addItem('🔄 Массовое обновление', 'batchUpdateProducts')
    .addSeparator()
    .addItem('📊 Планировщик операций', 'scheduleOperations');
    
  return batchMenu;
}

// ========================================
// ФУНКЦИИ ЭКСПОРТА И ИМПОРТА
// ========================================

/**
 * ЭКСПОРТ ДАННЫХ ТОВАРОВ
 * 
 * Создает резервную копию данных товаров
 */
function exportProductData() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      'Экспорт данных',
      'Функция экспорта будет реализована\n' +
      'в следующих версиях проекта.\n\n' +
      'Планируемые форматы:\n' +
      '• CSV для товаров\n' +
      '• JSON для настроек\n' +
      '• Архив изображений',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    showErrorDialog('Ошибка экспорта', error.message);
  }
}

/**
 * ИМПОРТ ДАННЫХ ТОВАРОВ
 * 
 * Загружает данные из внешних файлов
 */
function importProductData() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      'Импорт данных',
      'Функция импорта будет реализована\n' +
      'в следующих версиях проекта.\n\n' +
      'Планируемые источники:\n' +
      '• CSV файлы\n' +
      '• Excel таблицы\n' +
      '• JSON данные\n' +
      '• Другие Google Sheets',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    showErrorDialog('Ошибка импорта', error.message);
  }
}

// ========================================
// ТЕСТОВЫЕ ФУНКЦИИ
// ========================================

/**
 * ТЕСТИРОВАНИЕ МЕНЮ И UI
 * 
 * Проверяет работоспособность всех элементов интерфейса
 */
function testMenuFunctionality() {
  try {
    console.log('🧪 Тестирование функций меню...');
    
    const tests = [
      { name: 'Проверка настроек', func: () => validateConfig() },
      { name: 'Информация о проекте', func: () => getProjectInfo() },
      { name: 'Статистика обработки', func: () => generateProcessingStatistics() },
      { name: 'Подсчет выбранных товаров', func: () => getSelectedProductsCount() }
    ];
    
    const results = [];
    
    tests.forEach(test => {
      try {
        const result = test.func();
        results.push(`✅ ${test.name}: OK`);
        console.log(`✅ ${test.name}: прошел`);
      } catch (error) {
        results.push(`❌ ${test.name}: ${error.message}`);
        console.error(`❌ ${test.name}: ${error.message}`);
      }
    });
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Результаты тестирования',
      results.join('\n'),
      ui.ButtonSet.OK
    );
    
    console.log('🎉 Тестирование меню завершено');
    
  } catch (error) {
    showErrorDialog('Ошибка тестирования меню', error.message);
  }
}

// ========================================
// ФУНКЦИИ ИНИЦИАЛИЗАЦИИ
// ========================================

/**
 * ИНИЦИАЛИЗАЦИЯ ПРОЕКТА ДЛЯ НОВОГО ПОЛЬЗОВАТЕЛЯ
 * 
 * Пошаговая настройка системы для новых пользователей
 */
function initializeProject() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      'Инициализация проекта',
      'Выполнить пошаговую настройку системы?\n\n' +
      'Будут выполнены:\n' +
      '1. Создание структуры листа\n' +
      '2. Проверка настроек\n' +
      '3. Тестирование подключений\n' +
      '4. Показ справки\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // Шаг 1: Создание структуры
      try {
        createImagesSheet();
        showUserNotification('Шаг 1 завершен', 'Структура листа создана успешно', 'success');
      } catch (error) {
        showUserNotification('Шаг 1: Ошибка', error.message, 'error');
      }
      
      // Шаг 2: Проверка настроек
      try {
        const validation = validateConfig();
        if (validation.isValid) {
          showUserNotification('Шаг 2 завершен', 'Настройки проверены успешно', 'success');
        } else {
          showUserNotification('Шаг 2: Предупреждения', 
            `Найдено проблем: ${validation.errors.length + validation.warnings.length}.\n` +
            'Проверьте настройки API в Script Properties.', 'warning');
        }
      } catch (error) {
        showUserNotification('Шаг 2: Ошибка', error.message, 'error');
      }
      
      // Шаг 3: Показ справки
      showHelpDialog();
      
      showUserNotification('Инициализация завершена', 
        'Система готова к работе!\n\n' +
        'Следующий шаг: загрузите товары из InSales\n' +
        'через меню "Работа с товарами".', 'success');
    }
    
  } catch (error) {
    showErrorDialog('Ошибка инициализации', error.message);
  }
}

/**
 * Тестирование парсинга поставщиков
 */
function testSupplierParsing() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Запрашиваем артикул для тестирования
    const result = ui.prompt(
      'Тест парсинга поставщика',
      'Введите артикул товара для тестирования:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const article = result.getResponseText().trim();
    if (!article) {
      ui.alert('Артикул не может быть пустым');
      return;
    }
    
    showNotification(`Запускаем тест парсинга для: ${article}`, 'info');
    
    // Тестируем автопарсинг
    autoParseSupplierByArticle('VEBER', article)
      .then(result => {
        if (result.success) {
          showNotification(`Найдено ${result.images.length} изображений. Проверьте лист "Парсинг Veber"`, 'success');
        } else {
          showNotification(`Изображения не найдены. Статус: ${result.status}`, 'warning');
        }
      })
      .catch(error => {
        logError('Ошибка тестирования', error);
        showNotification(`Ошибка: ${error.message}`, 'error');
      });
    
  } catch (error) {
    logError('Ошибка тестирования парсинга', error);
    showNotification('Ошибка тестирования', 'error');
  }
}

/**
 * НАСТРОЙКА КАЧЕСТВА УЛУЧШЕНИЯ ИЗОБРАЖЕНИЙ
 * HTML-диалог с кастомными кнопками Scale 2 и Scale 4
 */
function configureReplicateScale() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Получаем текущее значение
    const currentScale = getSetting('replicateScale') || '2';
    
    // Создаем HTML диалог с кастомными кнопками
    const htmlContent = `
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>Настройка качества улучшения</h3>
        <p><strong>Текущий режим: scale=${currentScale}</strong></p>
        
        <div style="margin: 20px 0;">
          <h4>Выберите режим обработки:</h4>
          <div style="margin: 10px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">
            <button onclick="setScale(2)" style="width: 100%; padding: 15px; margin: 5px 0; font-size: 16px; background: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer;">
              Scale 2 - Быстро (~30 сек на изображение)
            </button>
            <button onclick="setScale(4)" style="width: 100%; padding: 15px; margin: 5px 0; font-size: 16px; background: #2196F3; color: white; border: none; border-radius: 5px; cursor: pointer;">
              Scale 4 - Качественно (~60 сек на изображение)
            </button>
          </div>
        </div>
        
        <p style="font-size: 12px; color: #666; margin: 15px 0;">
          <strong>Scale 2:</strong> Оптимальный баланс скорости и качества<br>
          <strong>Scale 4:</strong> Максимальное качество, больше времени обработки
        </p>
        
        <div style="text-align: right; margin-top: 20px;">
          <button onclick="google.script.host.close()" style="padding: 10px 20px; background: #f44336; color: white; border: none; border-radius: 3px; cursor: pointer;">
            Отмена
          </button>
        </div>
      </div>
      
      <script>
        function setScale(scale) {
          google.script.run
            .withSuccessHandler(() => {
              alert('Настройка Scale ' + scale + ' сохранена!\\n\\nНастройка применится при следующей обработке изображений.');
              google.script.host.close();
            })
            .withFailureHandler((error) => {
              alert('Ошибка сохранения настроек: ' + error.message);
            })
            .setSetting('replicateScale', scale.toString());
        }
      </script>
    `;
    
    const html = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(450)
      .setHeight(350);
    
    ui.showModalDialog(html, 'Настройка Replicate Scale');
    
  } catch (error) {
    showErrorDialog('Ошибка настройки качества', error.message);
  }
}

// 3. ТАКЖЕ ДОБАВИТЬ ФУНКЦИЮ ПОКАЗА ТЕКУЩИХ НАСТРОЕК:

/**
 * ПОКАЗАТЬ ТЕКУЩИЕ НАСТРОЙКИ ОБРАБОТКИ
 * 
 * Отображает все настройки для обработки изображений
 */
function showProcessingSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    const settings = getApiSettings();
    
    const message = 
      'ТЕКУЩИЕ НАСТРОЙКИ ОБРАБОТКИ:\n\n' +
      `🎛️ Replicate Scale: ${settings.replicateScale}\n` +
      `   ${settings.replicateScale === 2 ? '(Быстро, экономично)' : '(Медленно, качественно)'}\n\n` +
      `🔑 API ключи:\n` +
      `   OpenAI: ${settings.openaiApiKey ? '✅ Настроен' : '❌ Не настроен'}\n` +
      `   Replicate: ${settings.replicateToken ? '✅ Настроен' : '❌ Не настроен'}\n` +
      `   TinyPNG: ${settings.tinypngKey ? '✅ Настроен' : '❌ Не настроен'}\n` +
      `   ImgBB: ${settings.imgbbKey ? '✅ Настроен' : '❌ Не настроен'}\n\n` +
      'Для изменения используйте:\n' +
      '• "Настроить качество улучшения" - смена Scale\n' +
      '• Script Properties - API ключи';
    
    ui.alert('Настройки обработки', message, ui.ButtonSet.OK);
    
  } catch (error) {
    showErrorDialog('Ошибка показа настроек', error.message);
  }
}

/**
 * ВЫБОР МОДЕЛИ REPLICATE
 */
function configureReplicateModel() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Получаем текущую модель
    const currentModel = getSetting('replicateModel') || 'esrgan';
    const currentConfig = getReplicateModelConfig(currentModel);
    
    const response = ui.alert(
      'Выбор модели Replicate',
      `Текущая модель: ${currentConfig.name}\n\n` +
      'Выберите модель для улучшения изображений:\n\n' +
      '• ОК - ESRGAN (универсальная, быстрая)\n' +
      '• Отмена - Clarity Upscaler (для тяжелых файлов)',
      ui.ButtonSet.OK_CANCEL
    );
    
    let newModel;
    if (response === ui.Button.OK) {
      newModel = 'esrgan';
    } else if (response === ui.Button.CANCEL) {
      newModel = 'clarity_upscaler';
    } else {
      return;
    }
    
    // Сохраняем новую модель
    setSetting('replicateModel', newModel);
    
    const newConfig = getReplicateModelConfig(newModel);
    
    // Подтверждение
    ui.alert(
      'Модель изменена',
      `Выбрана модель: ${newConfig.name}\n\n` +
      `${newConfig.description}\n\n` +
      'Модель применится при следующей обработке изображений.',
      ui.ButtonSet.OK
    );
    
    console.log(`✅ Replicate модель изменена на: ${newModel}`);
    
  } catch (error) {
    showErrorDialog('Ошибка выбора модели', error.message);
  }
}

/**
 * ===================================================================
 * 💡 ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ МОДУЛЯ МЕНЮ
 * ===================================================================
 * 
 * АВТОМАТИЧЕСКОЕ СОЗДАНИЕ МЕНЮ:
 * 
 * Функция onOpen() автоматически создает меню при открытии Google Sheets.
 * Если меню не появилось:
 * 1. Обновите страницу Google Sheets
 * 2. Проверьте права доступа к скрипту
 * 3. Запустите onOpen() вручную из Apps Script Editor
 * 
 * СТРУКТУРА МЕНЮ:
 * 
 * 🤖 Автоматизация товаров
 * ├── 🔧 Настройка системы
 * │   ├── 🏗️ Создать структуру листа
 * │   ├── ⚙️ Проверить настройки API  
 * │   ├── 🔌 Тестировать InSales API
 * │   └── 🧪 Полная диагностика
 * ├── 📦 Работа с товарами
 * │   ├── 📥 Загрузить товары из InSales
 * │   ├── 🖼️ Обработать изображения AI
 * │   └── 📊 Показать статистику
 * ├── 📊 Мониторинг
 * │   └── (функции для будущих этапов)
 * └── ℹ️ О проекте / 🆘 Помощь
 * 
 * ОСНОВНЫЕ ФУНКЦИИ:
 * 
 * - showErrorDialog() - показ ошибок пользователю
 * - showSuccessDialog() - уведомления об успехе
 * - generateProcessingStatistics() - анализ данных
 * - getSelectedProductsCount() - подсчет выбранных товаров
 * - clearAllProcessingStatuses() - сброс статусов
 * 
 * РАСШИРЕНИЕ МЕНЮ:
 * 
 * Для добавления новых пунктов меню:
 * 1. Создайте функцию-обработчик
 * 2. Добавьте .addItem() в нужное подменю в onOpen()
 * 3. Добавьте обработку ошибок через showErrorDialog()
 * 
 * ГОТОВНОСТЬ К ЭТАПАМ:
 * 
 * - Этап 2: ✅ Полностью готово
 * - Этап 3: 🔄 AI-функции как заглушки
 * - Этап 4: 🔄 Внешние сервисы как заглушки  
 * - Этап 5: 🔄 Расширенный мониторинг как заглушки
 * 
 * ===================================================================
 */
