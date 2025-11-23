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

     // НОВЫЙ ФУНКЦИОНАЛ: Импорт полных карточек
     .addItem('🆕 Импорт товаров от поставщиков', 'showFullProductImportDialog')
     .addItem('📤 Создать товары в InSales', 'createProductsInInsalesMenu')
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
     .addItem('📋 Обновить структуру таблицы', 'updateSheetStructure')
     .addItem('🧹 Очистить статусы обработки', 'clearProcessingStatuses')
     .addSeparator()
     .addItem('📖 Управление справочником параметров', 'showSpecificationReferenceMenu')
     .addItem('🔄 Управление параметрами', 'showUnifiedParameterDialog')
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


// ========================================
// НОВЫЙ ФУНКЦИОНАЛ: ИМПОРТ ПОЛНЫХ КАРТОЧЕК ТОВАРОВ
// ========================================

/**
 * ДИАЛОГ ИМПОРТА ТОВАРОВ ОТ ПОСТАВЩИКОВ
 *
 * Показывает пользовательский интерфейс для пошагового импорта:
 * 1. Парсинг полных карточек от поставщиков
 * 2. Нормализация характеристик
 * 3. AI-рерайт описаний
 * 4. Проверка дубликатов
 * 5. Создание в InSales
 */
function showFullProductImportDialog() {
  try {
    const ui = SpreadsheetApp.getUi();

    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
      background: #f5f5f5;
    }
    .container {
      background: white;
      border-radius: 8px;
      padding: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
    }
    .step {
      background: #f8f9fa;
      border-left: 4px solid #1a73e8;
      padding: 15px;
      margin: 10px 0;
      border-radius: 4px;
    }
    .step h3 {
      margin-top: 0;
      color: #333;
    }
    .btn {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      margin: 5px;
    }
    .btn:hover {
      background: #1557b0;
    }
    .btn-success {
      background: #34a853;
    }
    .btn-success:hover {
      background: #2d8e47;
    }
    .btn-warning {
      background: #fbbc04;
      color: #333;
    }
    .btn-warning:hover {
      background: #e5aa04;
    }
    .info-box {
      background: #e8f0fe;
      border: 1px solid #d2e3fc;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
    }
    .warning-box {
      background: #fef7e0;
      border: 1px solid #fce8b2;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
    }
    .input-group {
      margin: 15px 0;
    }
    .input-group label {
      display: block;
      margin-bottom: 5px;
      font-weight: bold;
    }
    .input-group input {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      box-sizing: border-box;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🆕 Импорт товаров от поставщиков</h2>

    <div class="info-box">
      <strong>📋 Инструкция:</strong><br>
      1. Введите артикулы товаров (по одному на строку)<br>
      2. Выберите поставщика<br>
      3. Система автоматически выполнит все этапы обработки<br>
      4. Проверьте результаты в таблице перед публикацией
    </div>

    <div class="input-group">
      <label>Артикулы товаров:</label>
      <textarea id="articles" rows="10" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"
        placeholder="Введите артикулы, по одному на строку. Например:
VEBER-10x42
STURMAN-8x32
БН-123"></textarea>
    </div>

    <div class="input-group">
      <label>Поставщик:</label>
      <select id="supplier" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
        <option value="veber">Veber.ru</option>
        <option value="sturman">Sturman.ru</option>
      </select>
    </div>

    <div class="warning-box">
      <strong>⚠️ Внимание:</strong> Процесс может занять несколько минут на товар.
      Рекомендуется импортировать не более 10 товаров за раз.
    </div>

    <div class="step">
      <h3>Этапы обработки:</h3>
      <ol>
        <li>🔍 Парсинг полной карточки (описание, характеристики, изображения, цена)</li>
        <li>📊 Нормализация характеристик по справочнику</li>
        <li>🤖 AI-рерайт описания для уникальности</li>
        <li>🔎 Проверка дубликатов в InSales</li>
        <li>✅ Запись в таблицу для ручной проверки</li>
      </ol>
    </div>

    <div style="text-align: center; margin-top: 20px;">
      <button class="btn btn-success" onclick="runFullImport()">
        🚀 Начать импорт
      </button>
      <button class="btn" onclick="runStepByStep()">
        📝 Пошаговый режим
      </button>
      <button class="btn btn-warning" onclick="google.script.host.close()">
        ❌ Отмена
      </button>
    </div>
  </div>

  <script>
    function runFullImport() {
      const articles = document.getElementById('articles').value.trim();
      const supplier = document.getElementById('supplier').value;

      if (!articles) {
        alert('Пожалуйста, введите артикулы товаров');
        return;
      }

      const articleList = articles.split('\\n').map(a => a.trim()).filter(a => a);

      if (articleList.length === 0) {
        alert('Список артикулов пуст');
        return;
      }

      if (articleList.length > 30) {
        alert('Максимум 30 товаров за раз. Пожалуйста, уменьшите количество.');
        return;
      }

      if (!confirm(\`Начать импорт \${articleList.length} товаров от \${supplier}?\\n\\nПроцесс может занять \${articleList.length * 2} минут.\`)) {
        return;
      }

      document.body.innerHTML = '<div class="container"><h2>⏳ Импорт в процессе...</h2><p>Пожалуйста, не закрывайте это окно.</p></div>';

      google.script.run
        .withSuccessHandler(onImportSuccess)
        .withFailureHandler(onImportError)
        .executeFullProductImport(articleList, supplier);
    }

    function runStepByStep() {
      alert('Пошаговый режим:\\n\\n1. Используйте меню "🔍 Спарсить у поставщиков"\\n2. Затем выберите нужные функции вручную');
      google.script.host.close();
    }

    function onImportSuccess(result) {
      alert(\`✅ Импорт завершен!\\n\\nУспешно: \${result.success}\\nОшибки: \${result.errors}\\n\\nПроверьте результаты в таблице.\`);
      google.script.host.close();
    }

    function onImportError(error) {
      alert('❌ Ошибка импорта: ' + error.message);
      google.script.host.close();
    }
  </script>
</body>
</html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(700);

    ui.showModalDialog(htmlOutput, 'Импорт товаров от поставщиков');

  } catch (error) {
    handleError(error, 'Диалог импорта товаров');
    SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось открыть диалог импорта: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * ВЫПОЛНЕНИЕ ПОЛНОГО ИМПОРТА ТОВАРОВ
 *
 * @param {Array<string>} articles - Список артикулов
 * @param {string} supplier - Поставщик (veber, sturman)
 * @returns {Object} Результат импорта
 */
function executeFullProductImport(articles, supplier) {
  try {
    logInfo(`🚀 Полный импорт ${articles.length} товаров от ${supplier}`);

    // ✅ ИСПРАВЛЕНИЕ: Очищаем список ненормализованных значений в начале импорта
    // Это предотвращает накопление значений из предыдущих импортов
    clearUnnormalizedValues();

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < articles.length; i++) {
      const article = articles[i];

      try {
        logInfo(`[${i + 1}/${articles.length}] Импорт ${article}`);

        // 1. Парсинг полной карточки
        let productData;
        if (supplier === 'veber') {
          productData = parseVeberFullProduct(article);
        } else if (supplier === 'sturman') {
          productData = parseSturmanFullProduct(article);
        } else {
          throw new Error(`Неизвестный поставщик: ${supplier}`);
        }

        if (!productData) {
          throw new Error('Не удалось спарсить товар');
        }

        // ✅ Отладочное логирование
        logInfo(`   📊 Спарсенные данные: Цена="${productData.price}", Остаток="${productData.stock}", Гарантия="${productData.warranty || 'нет'}"`);

        // 2. Нормализация характеристик
        const specsRaw = JSON.parse(productData.specifications || '{}');

        // 🆕 Добавляем гарантию в характеристики (она парсится отдельно на Veber.ru)
        if (productData.warranty) {
          specsRaw['Гарантия'] = productData.warranty;
          logInfo(`   ✅ Добавлена гарантия в характеристики: "${productData.warranty}"`);
        }

        const normalized = normalizeSpecifications(specsRaw, supplier, article);  // ✅ ИСПРАВЛЕНИЕ: Передаем артикул для связи с ненормализованными значениями

        // 3. AI-рерайт описания
        const aiResult = generateProductDescription({
          article: article,
          productName: productData.title,
          description: productData.description,
          specifications: normalized.normalized,
          brand: productData.brand,
          categories: productData.categories
        });

        // 4. Проверка дубликатов
        const matchResult = checkProductDuplicate(article, productData.title);

        // 5. Запись в таблицу
        const supplierImages = Array.isArray(productData.images)
          ? productData.images.join('\n')
          : (productData.images || '');

        // ✅ Преобразуем остаток: "В наличии" → 5, остальное → 0
        const stockQuantity = (productData.stock && productData.stock.toLowerCase().includes('наличии')) ? 5 : 0;

        // ✅ ИСПРАВЛЕНИЕ: Синхронизация порядка параметров в колонках P и Q
        // Используем новую функцию formatRawSpecsInNormalizedOrder для упорядочивания сырых данных
        const normalizedOrder = normalized.normalized;  // Нормализованные параметры

        writeFullProductData({
          article: article,
          productName: productData.title,
          description: productData.description,
          descriptionRewritten: aiResult.rewrittenDescription,
          shortDescription: aiResult.shortDescription,
          specificationsRaw: formatRawSpecsInNormalizedOrder(specsRaw, normalizedOrder, supplier),  // ✅ НОВАЯ функция для синхронизации
          specificationsNormalized: JSON.stringify(normalizedOrder),                                 // ✅ ИСПРАВЛЕНО: JSON вместо текстового формата
          price: productData.price,
          stock: stockQuantity,  // ✅ Числовое значение остатка
          categories: productData.categories,
          brand: productData.brand,
          supplierImages: supplierImages,
          matchStatus: matchResult.matchStatus,
          matchConfidence: matchResult.confidence,
          importStatus: 'Импортирован, требует проверки'
        });

        logInfo(`✅ [${i + 1}/${articles.length}] ${article}: успешно импортирован`);
        successCount++;

        // Пауза между товарами
        if (i < articles.length - 1) {
          Utilities.sleep(3000);
        }

      } catch (error) {
        logError(`❌ [${i + 1}/${articles.length}] ${article}: ${error.message}`, error);
        errors.push(`${article}: ${error.message}`);
        errorCount++;
      }
    }

    logInfo(`✅ Импорт завершен: успешно ${successCount}, ошибок ${errorCount}`);

    return {
      success: successCount,
      errors: errorCount,
      errorDetails: errors
    };

  } catch (error) {
    handleError(error, 'Полный импорт товаров');
    throw error;
  }
}

// ========================================
// УПРАВЛЕНИЕ СПРАВОЧНИКОМ ПАРАМЕТРОВ
// ========================================

/**
 * МЕНЮ УПРАВЛЕНИЯ СПРАВОЧНИКОМ ПАРАМЕТРОВ
 *
 * Показывает диалог с опциями для работы со справочником
 */
function showSpecificationReferenceMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    let message = '📖 УПРАВЛЕНИЕ СПРАВОЧНИКОМ ПАРАМЕТРОВ\n\n';

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const paramCount = data.length - 1; // Минус заголовок
      message += `✅ Справочник существует\n`;
      message += `   Параметров в справочнике: ${paramCount}\n\n`;
      message += 'Выберите действие:\n\n';
      message += '• [Да] - открыть лист справочника\n';
      message += '• [Нет] - пересоздать заново из SPEC_MAPPING\n';
      message += '• [Отмена] - загрузить данные из CSV файла';
    } else {
      message += '⚠️ Справочник не найден\n\n';
      message += 'Справочник параметров необходим для корректной нормализации характеристик товаров.\n\n';
      message += 'Выберите действие:\n\n';
      message += '• [Да] - создать справочник из SPEC_MAPPING (10 параметров)\n';
      message += '• [Нет] - импортировать из CSV файла (~50 параметров)\n';
      message += '• [Отмена] - закрыть диалог';
    }

    const response = ui.alert(
      'Справочник параметров',
      message,
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.YES) {
      if (sheet) {
        // Открываем лист
        ss.setActiveSheet(sheet);
        ui.alert('✅ Лист справочника открыт', 'Теперь вы можете редактировать параметры, синонимы и функции нормализации.', ui.ButtonSet.OK);
      } else {
        // Создаем справочник из базовых данных
        const createdSheet = initializeSpecificationReferenceSheet();
        if (createdSheet) {
          ss.setActiveSheet(createdSheet);
          ui.alert(
            '✅ Справочник создан',
            `Создан лист "${SHEET_NAMES.SPEC_REFERENCE}" с ${createdSheet.getLastRow() - 1} параметрами.\n\n` +
            'Вы можете:\n' +
            '• Добавлять новые параметры\n' +
            '• Редактировать синонимы для поставщиков\n' +
            '• Включать/отключать параметры через чекбокс\n' +
            '• Настраивать функции нормализации',
            ui.ButtonSet.OK
          );
        }
      }
    } else if (response === ui.Button.NO) {
      if (sheet) {
        // Пересоздать справочник
        const confirm = ui.alert(
          '⚠️ Подтверждение',
          'Это действие удалит текущий справочник и создаст новый из базовых данных SPEC_MAPPING.\n\n' +
          'Все ваши изменения будут потеряны!\n\nПродолжить?',
          ui.ButtonSet.YES_NO
        );

        if (confirm === ui.Button.YES) {
          const createdSheet = initializeSpecificationReferenceSheet();
          if (createdSheet) {
            ss.setActiveSheet(createdSheet);
            ui.alert('✅ Справочник пересоздан', 'Лист справочника обновлен с базовыми параметрами.', ui.ButtonSet.OK);
          }
        }
      } else {
        // Импорт из CSV
        showCSVImportDialog();
      }
    } else if (response === ui.Button.CANCEL) {
      // Третья кнопка
      if (sheet) {
        // Если справочник существует - импорт из CSV
        showCSVImportDialog();
      }
      // Если справочника нет - просто закрываем
    }

  } catch (error) {
    handleError(error, 'Меню справочника параметров');
    SpreadsheetApp.getUi().alert(
      'Ошибка',
      'Не удалось открыть меню справочника: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ДИАЛОГ ИМПОРТА CSV
 */
function showCSVImportDialog() {
  const ui = SpreadsheetApp.getUi();

  const message = '📄 ИМПОРТ ИЗ CSV\n\n' +
    'Для импорта параметров из CSV файла:\n\n' +
    '1. Откройте CSV файл в текстовом редакторе\n' +
    '2. Скопируйте всё содержимое (Ctrl+A, Ctrl+C)\n' +
    '3. Нажмите ОК и вставьте в следующее окно\n\n' +
    'Формат CSV:\n' +
    '• Разделитель: запятая (,)\n' +
    '• Кодировка: UTF-8\n' +
    '• Первая строка: заголовки\n\n' +
    'Готовы продолжить?';

  const response = ui.alert('Импорт CSV', message, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    // Показываем промпт для ввода CSV
    const csvResponse = ui.prompt(
      'Вставьте содержимое CSV',
      'Вставьте сюда полное содержимое CSV файла (Ctrl+V):',
      ui.ButtonSet.OK_CANCEL
    );

    if (csvResponse.getSelectedButton() === ui.Button.OK) {
      const csvText = csvResponse.getResponseText();

      if (!csvText || csvText.trim().length < 10) {
        ui.alert('Ошибка', 'CSV данные не были введены или слишком короткие', ui.ButtonSet.OK);
        return;
      }

      try {
        // Создаем справочник из CSV
        const sheet = initializeSpecificationReferenceSheet(csvText);

        if (sheet) {
          SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
          const paramCount = sheet.getLastRow() - 1;
          ui.alert(
            '✅ Импорт завершен',
            `Справочник создан из CSV файла.\n\n` +
            `Импортировано параметров: ${paramCount}\n\n` +
            `Лист "${SHEET_NAMES.SPEC_REFERENCE}" активирован для редактирования.`,
            ui.ButtonSet.OK
          );
        }
      } catch (error) {
        ui.alert(
          '❌ Ошибка импорта',
          `Не удалось импортировать CSV:\n\n${error.message}\n\n` +
          'Проверьте формат файла и попробуйте снова.',
          ui.ButtonSet.OK
        );
      }
    }
  }
}


// ========================================
// СОЗДАНИЕ ТОВАРОВ В INSALES
// ========================================

/**
 * Основное меню создания товаров в InSales
 */
function createProductsInInsalesMenu() {
  try {
    logInfo('📤 Запуск создания товаров в InSales');

    const ui = SpreadsheetApp.getUi();

    // Читаем выбранные товары
    const selectedProducts = readSelectedProductsForImport();

    if (selectedProducts.length === 0) {
      ui.alert(
        '⚠️ Нет выбранных товаров',
        'Пожалуйста, отметьте чекбоксами товары, которые нужно создать в InSales.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Сохраняем выбранные товары во временное свойство для использования после диалога
    PropertiesService.getScriptProperties().setProperty(
      'temp_selected_products',
      JSON.stringify(selectedProducts.map(p => p.sku))  // Сохраняем только SKU
    );

    // Показываем диалог выбора категории
    showCategoryPickerDialog(selectedProducts.length);

  } catch (error) {
    handleError(error, 'Меню создания товаров в InSales');
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка',
      `Произошла ошибка при создании товаров:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/**
 * Создание товаров в InSales (основная логика)
 *
 * @param {Array} products - Массив товаров для создания
 * @returns {Object} Результат импорта
 */
function createProductsInInsales(products) {
  const result = {
    total: products.length,
    created: 0,
    skipped: 0,
    errors: 0,
    details: []
  };

  try {
    logInfo(`📦 Начинаем создание ${products.length} товаров в InSales`);

    for (let i = 0; i < products.length; i++) {
      const product = products[i];
      const productNum = i + 1;

      logInfo(`\n${'='.repeat(60)}`);
      logInfo(`📦 Товар ${productNum}/${products.length}: "${product.title}" (${product.sku})`);
      logInfo(`${'='.repeat(60)}`);

      try {
        // Обновляем статус на "Импорт..."
        updateImportStatus(product.sku, STATUS_VALUES.IMPORT.IMPORTING);

        // ПРОВЕРКА ДУБЛИКАТОВ УДАЛЕНА:
        // Раньше здесь была проверка через checkProductExists(), которая делала сотни API запросов.
        // Теперь полагаемся на InSales API - если товар с таким SKU существует,
        // API вернёт ошибку (обычно 422 Unprocessable Entity), и мы обработаем её ниже.

        // Создаем товар напрямую
        logInfo('📦 Создаем товар в InSales');
        const createdProduct = createInsalesProduct(product);

        if (!createdProduct || !createdProduct.id) {
          throw new Error('Не удалось создать товар - API не вернул ID');
        }

        logInfo(`✅ Товар создан с ID: ${createdProduct.id}`);

        // Шаг 3: Добавляем изображения
        if (product.imageUrls && product.imageUrls.length > 0) {
          logInfo(`📸 Добавляем ${product.imageUrls.length} изображений`);
          const imagesAdded = addProductImages(createdProduct.id, product.imageUrls);
          logInfo(`✅ Добавлено ${imagesAdded} изображений`);
        } else {
          logWarning('⚠️ Нет изображений для добавления');
        }

        // Шаг 4: Обновляем статус
        updateImportStatus(product.sku, STATUS_VALUES.IMPORT.IMPORTED, createdProduct.id);

        result.created++;
        result.details.push({
          sku: product.sku,
          title: product.title,
          status: 'Создан ✅',
          insalesId: createdProduct.id,
          insalesUrl: `https://binokl.shop/admin/products/${createdProduct.id}`
        });

        logInfo(`✅ Товар "${product.title}" успешно импортирован (${productNum}/${products.length})`);

      } catch (productError) {
        handleError(productError, `Создание товара ${product.sku}`);

        // Проверяем, не является ли ошибка дубликатом SKU
        const errorMessage = productError.message || '';
        const isDuplicate = errorMessage.includes('уже существует') ||
                           errorMessage.includes('already exists') ||
                           errorMessage.includes('duplicate') ||
                           errorMessage.includes('SKU') ||
                           errorMessage.includes('422');

        if (isDuplicate) {
          logWarning(`⚠️ Товар с артикулом "${product.sku}" уже существует в InSales`);
          updateImportStatus(product.sku, STATUS_VALUES.IMPORT.EXISTS);

          result.skipped++;
          result.details.push({
            sku: product.sku,
            title: product.title,
            status: 'Пропущен (уже существует)',
            error: 'Товар с таким артикулом уже есть в InSales'
          });
        } else {
          updateImportStatus(product.sku, STATUS_VALUES.IMPORT.ERROR);

          result.errors++;
          result.details.push({
            sku: product.sku,
            title: product.title,
            status: 'Ошибка ❌',
            error: productError.message
          });
        }
      }

      // Пауза между запросами
      if (i < products.length - 1) {
        Utilities.sleep(IMPORT_SETTINGS.API_DELAY_MS);
      }
    }

    logInfo(`\n${'='.repeat(60)}`);
    logInfo(`✅ ИМПОРТ ЗАВЕРШЕН`);
    logInfo(`   Всего товаров: ${result.total}`);
    logInfo(`   Создано: ${result.created}`);
    logInfo(`   Пропущено: ${result.skipped}`);
    logInfo(`   Ошибок: ${result.errors}`);
    logInfo(`${'='.repeat(60)}\n`);

    return result;

  } catch (error) {
    handleError(error, 'Создание товаров в InSales');
    throw error;
  }
}


/**
 * Показ диалога с результатами импорта
 *
 * @param {Object} result - Результат импорта
 */
function showImportResultDialog(result) {
  try {
    const ui = SpreadsheetApp.getUi();

    let message = `📊 РЕЗУЛЬТАТЫ ИМПОРТА\n\n`;
    message += `Всего товаров: ${result.total}\n`;
    message += `✅ Создано: ${result.created}\n`;
    message += `⚠️ Пропущено (уже существуют): ${result.skipped}\n`;
    message += `❌ Ошибок: ${result.errors}\n\n`;

    if (result.details.length > 0) {
      message += `ДЕТАЛИ:\n\n`;

      // Показываем первые 10 товаров
      const limit = Math.min(10, result.details.length);
      for (let i = 0; i < limit; i++) {
        const detail = result.details[i];
        message += `${i + 1}. ${detail.sku} - ${detail.status}\n`;
      }

      if (result.details.length > 10) {
        message += `\n... и еще ${result.details.length - 10} товаров\n`;
      }
    }

    message += `\nПроверьте логи выполнения для подробностей.`;

    const title = result.errors === 0 && result.created > 0
      ? '✅ Импорт завершен успешно'
      : result.errors > 0
      ? '⚠️ Импорт завершен с ошибками'
      : 'ℹ️ Импорт завершен';

    ui.alert(title, message, ui.ButtonSet.OK);

  } catch (error) {
    handleError(error, 'Показ результатов импорта');
  }
}


/**
 * ДИАЛОГ ВЫБОРА КАТЕГОРИИ ДЛЯ ИМПОРТА
 * ====================================
 */

/**
 * Показывает диалог выбора категории перед импортом товаров
 *
 * @param {number} productCount - Количество товаров для импорта
 */
function showCategoryPickerDialog(productCount) {
  try {
    logInfo(`📂 Открытие диалога выбора категории для ${productCount} товаров`);

    // Создаем HTML диалог из файла
    const htmlTemplate = HtmlService.createTemplateFromFile('CategoryPickerDialog');

    // Передаем параметры в HTML шаблон
    htmlTemplate.productCount = productCount;

    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(650)
      .setHeight(700);

    // Показываем модальный диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📂 Выбор категории для импорта');

    logInfo('✅ Диалог выбора категории открыт');

  } catch (error) {
    handleError(error, 'Открытие диалога выбора категории');
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка',
      `Не удалось открыть диалог: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/**
 * Возвращает список категорий из InSales для диалога выбора
 *
 * @returns {Array} Массив объектов категорий
 */
async function getCategoriesForImport() {
  try {
    logInfo('📋 Загрузка категорий из InSales для диалога');

    // ОПТИМИЗАЦИЯ: Загружаем только категории БЕЗ подсчета товаров
    // Подсчет товаров делает 1000+ запросов и занимает минуты
    const categories = await makeInsalesRequest('GET', INSALES_ENDPOINTS.COLLECTIONS);

    if (!categories || !Array.isArray(categories)) {
      logWarning('⚠️ Не удалось загрузить категории');
      return [];
    }

    logInfo(`📊 Загружено ${categories.length} категорий из InSales`);

    // Строим иерархию без подсчета товаров
    const categoryMap = {};
    categories.forEach(cat => {
      categoryMap[cat.id] = {
        id: cat.id,
        title: cat.title,
        parent_id: cat.parent_id,
        children: []
      };
    });

    // Связываем родителей и детей
    const rootCategories = [];
    categories.forEach(cat => {
      const category = categoryMap[cat.id];
      if (cat.parent_id && categoryMap[cat.parent_id]) {
        categoryMap[cat.parent_id].children.push(category);
      } else {
        rootCategories.push(category);
      }
    });

    // Преобразуем в плоский список с путями
    const flatCategories = [];

    function processCategory(category, level = 0, parentPath = '') {
      const path = parentPath ? `${parentPath} / ${category.title}` : category.title;

      flatCategories.push({
        id: category.id,
        title: category.title,
        path: path,
        level: level,
        parentId: category.parent_id
      });

      // Рекурсивно обрабатываем подкатегории
      if (category.children && category.children.length > 0) {
        category.children.forEach(child => {
          processCategory(child, level + 1, path);
        });
      }
    }

    // Обрабатываем все корневые категории
    rootCategories.forEach(rootCategory => {
      processCategory(rootCategory, 0, '');
    });

    logInfo(`✅ Загружено ${flatCategories.length} категорий для диалога`);

    return flatCategories;

  } catch (error) {
    handleError(error, 'Загрузка категорий для диалога');
    throw new Error(`Не удалось загрузить категории: ${error.message}`);
  }
}


/**
 * Запускает импорт товаров с несколькими выбранными категориями
 *
 * @param {Array<number>} categoryIds - Массив ID выбранных категорий
 * @returns {Object} Результат импорта
 */
function startProductImportWithCategories(categoryIds) {
  try {
    logInfo(`📦 Запуск импорта товаров в категории: ${categoryIds.join(', ')}`);

    // Проверяем, что передан массив
    if (!Array.isArray(categoryIds) || categoryIds.length === 0) {
      throw new Error('Не выбраны категории для импорта');
    }

    // Получаем сохраненные SKU товаров
    const tempProductsSku = PropertiesService.getScriptProperties().getProperty('temp_selected_products');

    if (!tempProductsSku) {
      throw new Error('Не найдены данные о выбранных товарах. Попробуйте снова.');
    }

    const selectedSku = JSON.parse(tempProductsSku);

    // Очищаем временное свойство
    PropertiesService.getScriptProperties().deleteProperty('temp_selected_products');

    // Первая категория будет основной (collection_id)
    // Остальные добавим через collections_ids
    const mainCategoryId = categoryIds[0];

    // Читаем товары с основной категорией
    const selectedProducts = readSelectedProductsForImport(mainCategoryId).filter(p =>
      selectedSku.includes(p.sku)
    );

    if (selectedProducts.length === 0) {
      throw new Error('Не удалось найти товары для импорта');
    }

    // Добавляем массив всех категорий к каждому товару
    selectedProducts.forEach(product => {
      product.collectionIds = categoryIds;  // Все категории для товара
    });

    logInfo(`✅ Подготовлено ${selectedProducts.length} товаров для импорта в ${categoryIds.length} категорий`);

    // Запускаем процесс создания
    const result = createProductsInInsales(selectedProducts);

    // Показываем результат
    showImportResultDialog(result);

    return {
      success: true,
      categoryIds: categoryIds,
      productsCount: selectedProducts.length
    };

  } catch (error) {
    handleError(error, 'Запуск импорта с выбранными категориями');
    throw error;
  }
}


/**
 * УСТАРЕВШАЯ: Запускает импорт товаров с одной категорией (для обратной совместимости)
 *
 * @deprecated Используйте startProductImportWithCategories() вместо этой функции
 * @param {number} categoryId - ID выбранной категории
 * @returns {Object} Результат импорта
 */
function startProductImportWithCategory(categoryId) {
  // Переадресуем на новую функцию с массивом из одной категории
  return startProductImportWithCategories([categoryId]);
}

// ========================================
// ДИАЛОГ СОПОСТАВЛЕНИЯ ПАРАМЕТРОВ
// ========================================

/**
 * Показывает диалог сопоставления параметров товара со справочником
 *
 * Позволяет пользователю:
 * - Просмотреть сопоставленные и несопоставленные параметры
 * - Вручную сопоставить параметры поставщика с параметрами справочника
 * - Создать новые параметры в справочнике
 * - Применить изменения локально (для одного товара) или глобально (для всех)
 *
 * @deprecated Используйте showUnifiedParameterDialog() вместо этой функции
 */
function showParameterMappingDialog() {
  try {
    logInfo('🔄 Открываем диалог сопоставления параметров');

    // Получаем выделенные товары
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    // Ищем первый выделенный товар (с чекбоксом)
    let selectedArticle = null;
    for (let i = 1; i < data.length; i++) {
      const isSelected = data[i][IMAGES_COLUMNS.CHECKBOX - 1];

      if (isSelected === true) {
        selectedArticle = String(data[i][IMAGES_COLUMNS.ARTICLE - 1]);
        break;
      }
    }

    if (!selectedArticle) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Товар не выбран',
        'Пожалуйста, выберите товар с помощью чекбокса в первой колонке таблицы.',
        ui.ButtonSet.OK
      );
      return;
    }

    logInfo(`   Выбран товар: ${selectedArticle}`);

    // Создаем HTML диалог
    const htmlTemplate = HtmlService.createTemplateFromFile('ParameterMappingDialog');

    // Передаем артикул через URL параметр
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(1200)
      .setHeight(700)
      .append(`<script>window.history.replaceState(null, '', '?article=${encodeURIComponent(selectedArticle)}');</script>`);

    // Показываем модальный диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔄 Сопоставление параметров товара');

    logInfo('✅ Диалог открыт');

  } catch (error) {
    handleError(error, 'Открытие диалога сопоставления параметров');
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка',
      `Не удалось открыть диалог: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
