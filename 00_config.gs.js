/**
 * ПОЛУЧЕНИЕ ОСНОВНОГО РАБОЧЕГО ЛИСТА
 * 
 * Безопасное получение единственного рабочего листа проекта
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Лист обработки изображений
 * @throws {Error} Если лист не найден
 */
function getImagesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.IMAGES);
  
  if (!sheet) {
    throw new Error(`Лист "${SHEET_NAMES.IMAGES}" не найден. Создайте структуру листа сначала.`);
  }
  
  return sheet;
}

/**
 * ПОЛУЧЕНИЕ ЛЮБОГО ЛИСТА ПО НАЗВАНИЮ
 * 
 * Универсальная функция получения лист/**
 * ===================================================================
 * МОДУЛЬ 00_config.gs - НАСТРОЙКИ И КОНСТАНТЫ
 * ===================================================================
 * 
 * Назначение: Центральное место для всех настроек проекта
 * Версия: 1.0
 * Проект: Обработка изображений товаров
 * 
 * Этот модуль содержит:
 * - Константы названий листов и колонок
 * - Функции чтения настроек из Google Sheets
 * - Валидация конфигурации
 * - Глобальные настройки проекта
 */

// =============================================================================
// 📊 КОНСТАНТЫ НАЗВАНИЙ ЛИСТОВ
// =============================================================================

/**
 * НАЗВАНИЯ ЛИСТОВ В GOOGLE SHEETS (МИНИМАЛЬНАЯ СТРУКТУРА)
 * 
 * Используется только один рабочий лист для максимальной простоты
 * ⚠️ ВАЖНО: Изменение этой константы повлияет на весь проект!
 */
const SHEET_NAMES = {
  IMAGES: 'Обработка изображений',         // Основной рабочий лист
  SPEC_REFERENCE: 'Справочник параметров'  // Служебный лист с эталонными параметрами
};

// =============================================================================
// 📋 КОНСТАНТЫ КОЛОНОК ЛИСТА "ОБРАБОТКА ИЗОБРАЖЕНИЙ"
// =============================================================================

/**
 * СТРУКТУРА КОЛОНОК ОСНОВНОГО РАБОЧЕГО ЛИСТА (УПРОЩЕННАЯ)
 * 
 * Добавлена колонка ID InSales для технических нужд (получаем динамически из API)
 * Убраны все лишние справочники и навигация
 */
const IMAGES_COLUMNS = {
  CHECKBOX: 1,              // A - ☑️ Чекбокс для выбора товаров
  ARTICLE: 2,               // B - Артикул товара
  INSALES_ID: 3,            // C - ID InSales
  PRODUCT_NAME: 4,          // D - Название товара
  ORIGINAL_IMAGES: 5,       // E - Исходные изображения из InSales
  SUPPLIER_IMAGES: 6,       // F - Парсинг Поставщика (НОВЫЙ)
  ADDITIONAL_IMAGES: 7,     // G - Дополнительные фото (НОВЫЙ)
  PROCESSED_IMAGES: 8,      // H - Обработанные изображения (СДВИГ +2)
  ALT_TAGS: 9,              // I - Alt-теги (СДВИГ +2)
  SEO_FILENAMES: 10,        // J - SEO имена файлов (СДВИГ +2)
  PROCESSING_STATUS: 11,    // K - Статус обработки (СДВИГ +2)
  INSALES_STATUS: 12,       // L - Статус InSales (СДВИГ +2)

  // ========== НОВЫЕ КОЛОНКИ ДЛЯ ИМПОРТА ТОВАРОВ ==========
  DESCRIPTION: 13,              // M - Описание от поставщика
  DESCRIPTION_REWRITTEN: 14,    // N - Рерайт описания (AI)
  SHORT_DESCRIPTION: 15,        // O - Краткое описание
  SPECIFICATIONS_RAW: 16,       // P - Характеристики (JSON от поставщика)
  SPECIFICATIONS_NORMALIZED: 17,// Q - Нормализованные характеристики (JSON)
  PRICE: 18,                    // R - Цена поставщика
  STOCK: 19,                    // S - Остаток
  CATEGORIES: 20,               // T - Категории
  BRAND: 21,                    // U - Бренд
  SERIES: 22,                   // V - Серия
  WEIGHT: 23,                   // W - Вес, г
  DIMENSIONS: 24,               // X - Габариты (ДxШxВ)
  PACKAGE_CONTENTS: 25,         // Y - Комплектация
  MATCH_STATUS: 26,             // Z - Статус сопоставления (новый/существующий/дубль)
  MATCH_CONFIDENCE: 27,         // AA - Уверенность в совпадении (%)
  IMPORT_STATUS: 28             // AB - Статус импорта
};

// =============================================================================
// 📦 КОНСТАНТЫ КОЛОНОК ЛИСТА "СПРАВОЧНИК ТОВАРОВ"
// =============================================================================

/**
 * СТРУКТУРА КОЛОНОК СПРАВОЧНИКА ТОВАРОВ (ОБНОВЛЕННАЯ)
 *
 * Общий справочник для связи между проектами и хранения базовой информации о товарах
 * Технические данные (ID, категории) хранятся здесь, в рабочем листе не отображаются
 */
const PRODUCTS_COLUMNS = {
  ARTICLE: 1,            // A - Артикул товара (основной ключ)
  INSALES_ID: 2,         // B - ID товара в InSales (технический, скрытый)
  PRODUCT_NAME: 3,       // C - Полное название товара
  CATEGORIES: 4,         // D - Все категории товара (через запятую)
  PRICE: 5,              // E - Актуальная цена
  IMAGES_COUNT: 6,       // F - Количество изображений
  STATUS: 7,             // G - Статус товара в InSales (активен/неактивен)
  DATE_ADDED: 8,         // H - Дата добавления в справочник
  DATE_UPDATED: 9,       // I - Дата последнего обновления
  ADMIN_LINK: 10         // J - Прямая ссылка в админку InSales
};

// =============================================================================
// 📋 КОНСТАНТЫ КОЛОНОК ЛИСТА "СПРАВОЧНИК ПАРАМЕТРОВ"
// =============================================================================

/**
 * СТРУКТУРА СПРАВОЧНИКА ПАРАМЕТРОВ ДЛЯ НОРМАЛИЗАЦИИ ХАРАКТЕРИСТИК
 *
 * Этот лист хранит эталонные названия параметров и маппинг синонимов от поставщиков
 */
const SPEC_REFERENCE_COLUMNS = {
  CHECKBOX: 1,              // A - Чекбокс (активность параметра)
  PARAMETER_NAME: 2,        // B - Параметр (эталонное название)
  FIELD_TYPE: 3,            // C - Тип поля (enum/число/текст/формат)
  ALLOWED_VALUES: 4,        // D - Допустимые значения (enum) - через точку с запятой
  REQUIRED: 5,              // E - Обязательное поле (да/нет)
  VEBER_SYNONYMS: 6,        // F - Синонимы Veber (через точку с запятой)
  STURMAN_SYNONYMS: 7,      // G - Синонимы Sturman (через точку с запятой)
  NORMALIZER_FUNCTION: 8,   // H - Функция нормализации (extractNumber/toUpperCase/custom)
  DESCRIPTION: 9            // I - Описание параметра
};

// =============================================================================
// ⚙️ КОНСТАНТЫ НАСТРОЕК
// =============================================================================

/**
 * НАЗВАНИЯ ПАРАМЕТРОВ В ЛИСТЕ "НАСТРОЙКИ"
 * 
 * Используются для чтения API ключей и других настроек из листа
 */
const SETTINGS_PARAMS = {
  // InSales API настройки
  INSALES_API_KEY: 'InSalesAPIKey',
  INSALES_PASSWORD: 'InSalesPassword', 
  INSALES_SHOP: 'InSalesShop',
  
  // Внешние API для обработки изображений
  REPLICATE_TOKEN: 'ReplicateToken',     // ИИ-увеличение изображений
  REPLICATE_SCALE: 'ReplicateScale',     // ИИ-степень улучшения
  REPLICATE_MODEL: 'ReplicateModel',
  TINYPNG_KEY: 'TinyPNGKey',            // Сжатие изображений
  IMGBB_KEY: 'imgbbKey',                // Хостинг изображений
  
  // AI для генерации контента с анализом изображений
  OPENAI_API_KEY: 'OpenAIAPIKey',      // Общий API ключ OpenAI
  OPENAI_ALT_ASSISTANT_ID: 'OpenAIAltAssistantID', // ID ассистента для анализа изображений и генерации alt-тегов
  AI_IMAGE_ANALYSIS: 'AI_Image_Analysis', // Включить анализ изображений (true/false)
  
  // Уведомления (опционально)
  TELEGRAM_TOKEN: 'TelegramToken',       // Токен бота Telegram
  TELEGRAM_CHAT_ID: 'TelegramChatID'    // ID чата для уведомлений
};

/**
 * КЛЮЧИ ДЛЯ SCRIPT PROPERTIES
 * 
 * Используются для безопасного хранения API ключей в Script Properties
 */
const SCRIPT_PROPERTIES_KEYS = {
  // InSales API настройки
  INSALES_API_KEY: 'insalesApiKey',
  INSALES_PASSWORD: 'insalesPassword', 
  INSALES_SHOP: 'insalesShop',
  
  // Внешние API для обработки изображений
  REPLICATE_TOKEN: 'replicateToken',
  REPLICATE_SCALE: 'replicateScale',
  REPLICATE_MODEL: 'replicateModel',
  TINYPNG_KEY: 'tinypngKey',
  IMGBB_KEY: 'imgbbKey',
  
  // AI для генерации контента
  OPENAI_API_KEY: 'openaiApiKey',
  OPENAI_ALT_ASSISTANT_ID: 'openaiAltAssistantId',
  AI_IMAGE_ANALYSIS: 'aiImageAnalysis',
  
  // Уведомления
  TELEGRAM_TOKEN: 'telegramToken',
  TELEGRAM_CHAT_ID: 'telegramChatId'
};
// =============================================================================
// 🎯 КОНСТАНТЫ ЗНАЧЕНИЙ
// =============================================================================

/**
 * СТАНДАРТНЫЕ ЗНАЧЕНИЯ СТАТУСОВ
 *
 * Используются для унификации статусов во всех модулях
 */
const STATUS_VALUES = {
  // Статусы обработки изображений
  PROCESSING: {
    NOT_PROCESSED: 'Не обработано',
    PROCESSING: 'Обработка...',
    COMPLETED: 'Обработано',
    ERROR: 'Ошибка'
  },

  // Статусы отправки в InSales
  INSALES: {
    NOT_SENT: 'Не отправлено',
    SENDING: 'Отправка...',
    SENT: 'Отправлено ✅',
    ERROR: 'Ошибка отправки'
  },

  // Статусы товаров
  PRODUCT: {
    ACTIVE: 'Активен',
    INACTIVE: 'Неактивен',
    DRAFT: 'Черновик'
  },

  // Статусы импорта товаров
  IMPORT: {
    NOT_IMPORTED: 'Не импортирован',
    IMPORTING: 'Импорт...',
    IMPORTED: 'Импортирован ✅',
    EXISTS: 'Уже существует',
    ERROR: 'Ошибка импорта'
  }
};

// =============================================================================
// 🌐 КОНСТАНТЫ INSALES API
// =============================================================================

/**
 * ЭНДПОИНТЫ INSALES API
 *
 * Все доступные эндпоинты для работы с InSales REST API
 */
const INSALES_ENDPOINTS = {
  // Товары
  PRODUCTS: '/admin/products.json',
  PRODUCT_BY_ID: '/admin/products/{product_id}.json',
  PRODUCT_IMAGES: '/admin/products/{product_id}/images.json',
  PRODUCT_IMAGE_BY_ID: '/admin/products/{product_id}/images/{image_id}.json',

  // Категории (коллекции)
  COLLECTIONS: '/admin/collections.json',
  COLLECTION_BY_ID: '/admin/collections/{collection_id}.json',

  // Связи товар-категория
  COLLECTS: '/admin/collects.json',
  COLLECT_BY_ID: '/admin/collects/{collect_id}.json',

  // Характеристики
  PROPERTIES: '/admin/properties.json',
  PROPERTY_BY_ID: '/admin/properties/{property_id}.json'
};

/**
 * КАТЕГОРИЯ ПО УМОЛЧАНИЮ ДЛЯ НОВЫХ ТОВАРОВ
 *
 * ID категории "Новинки" для размещения импортированных товаров
 */
const DEFAULT_CATEGORY_ID = 9069712;

/**
 * НАСТРОЙКИ ИМПОРТА ТОВАРОВ
 *
 * Параметры для создания товаров в InSales
 */
const IMPORT_SETTINGS = {
  // Создавать товары скрытыми для ручной проверки
  CREATE_HIDDEN: true,

  // Количество товара по умолчанию
  DEFAULT_QUANTITY_IN_STOCK: 5,      // Если "в наличии"
  DEFAULT_QUANTITY_OUT_OF_STOCK: 0,  // Если "под заказ"

  // Максимальное количество фото на товар
  MAX_IMAGES_PER_PRODUCT: 10,

  // Задержка между запросами (мс) для избежания rate limit
  API_DELAY_MS: 500
};

// =============================================================================
// 📖 ФУНКЦИИ ЧТЕНИЯ НАСТРОЕК
// =============================================================================

/**
 * ПОЛУЧЕНИЕ НАСТРОЕК API ИЗ SCRIPT PROPERTIES
 * 
 * Читает все API ключи из безопасного хранилища Script Properties
 * 
 * @returns {Object} Объект с настройками API
 * @example
 * const settings = getApiSettings();
 * console.log(settings.insalesApiKey);
 */
function getApiSettings() {
  try {
    console.log('📖 Читаем настройки API из Script Properties...');
    
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    // Возвращаем структурированный объект настроек
    const apiSettings = {
      // InSales настройки
      insalesApiKey: allProperties[SCRIPT_PROPERTIES_KEYS.INSALES_API_KEY] || '',
      insalesPassword: allProperties[SCRIPT_PROPERTIES_KEYS.INSALES_PASSWORD] || '',
      insalesShop: allProperties[SCRIPT_PROPERTIES_KEYS.INSALES_SHOP] || '',
      
      // Внешние API
      replicateToken: allProperties[SCRIPT_PROPERTIES_KEYS.REPLICATE_TOKEN] || '',
      replicateScale: parseInt(allProperties[SCRIPT_PROPERTIES_KEYS.REPLICATE_SCALE] || '2'),
      replicateModel: allProperties[SCRIPT_PROPERTIES_KEYS.REPLICATE_MODEL] || 'esrgan',
      tinypngKey: allProperties[SCRIPT_PROPERTIES_KEYS.TINYPNG_KEY] || '',
      imgbbKey: allProperties[SCRIPT_PROPERTIES_KEYS.IMGBB_KEY] || '',
      
      // AI для генерации контента с анализом изображений
      openaiApiKey: allProperties[SCRIPT_PROPERTIES_KEYS.OPENAI_API_KEY] || '',
      openaiAltAssistantId: allProperties[SCRIPT_PROPERTIES_KEYS.OPENAI_ALT_ASSISTANT_ID] || '',
      aiImageAnalysis: allProperties[SCRIPT_PROPERTIES_KEYS.AI_IMAGE_ANALYSIS] === 'true' || true,
      
      // Уведомления (опционально)
      telegramToken: allProperties[SCRIPT_PROPERTIES_KEYS.TELEGRAM_TOKEN] || '',
      telegramChatId: allProperties[SCRIPT_PROPERTIES_KEYS.TELEGRAM_CHAT_ID] || ''
    };
    
    console.log('✅ Настройки API успешно загружены из Script Properties');
    return apiSettings;
    
  } catch (error) {
    console.error('❌ Ошибка чтения настроек API:', error.message);
    throw new Error('Не удалось загрузить настройки API: ' + error.message);
  }
}

const REPLICATE_MODELS = {
  ESRGAN: {
    name: 'ESRGAN (стандартная)',
    version: 'f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa',
    description: 'Универсальная модель, быстрая обработка',
    maxScale: 4
  },
  CLARITY_UPSCALER: {
    name: 'Clarity Upscaler (тяжелые файлы)',
    version: 'dfad41707589d68ecdccd1dfa600d55a208f9310748e44bfe35b4a6291453d5e',
    description: 'Лучше работает с крупными изображениями',
    maxScale: 4
  }
};

function getReplicateModelConfig(modelKey) {
  switch (modelKey.toLowerCase()) {
    case 'clarity':
    case 'clarity_upscaler':
      return REPLICATE_MODELS.CLARITY_UPSCALER;
    case 'esrgan':
    default:
      return REPLICATE_MODELS.ESRGAN;
  }
}

/**
 * ПОЛУЧЕНИЕ КОНКРЕТНОЙ НАСТРОЙКИ ПО КЛЮЧУ
 * 
 * Читает одну настройку из Script Properties
 * 
 * @param {string} propertyKey - Ключ свойства из SCRIPT_PROPERTIES_KEYS
 * @returns {string} Значение настройки или пустая строка
 * @example
 * const apiKey = getSetting(SCRIPT_PROPERTIES_KEYS.INSALES_API_KEY);
 */
function getSetting(propertyKey) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const value = properties.getProperty(propertyKey);
    return value || '';
    
  } catch (error) {
    console.error(`❌ Ошибка чтения свойства "${propertyKey}":`, error.message);
    return '';
  }
}

/**
 * УСТАНОВКА ЗНАЧЕНИЯ НАСТРОЙКИ
 * 
 * Записывает или обновляет значение настройки в Script Properties
 * 
 * @param {string} propertyKey - Ключ свойства
 * @param {string} propertyValue - Значение свойства
 * @example
 * setSetting(SCRIPT_PROPERTIES_KEYS.INSALES_API_KEY, 'your_api_key_here');
 */
function setSetting(propertyKey, propertyValue) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(propertyKey, propertyValue);
    console.log(`✅ Установлено свойство "${propertyKey}"`);
    
  } catch (error) {
    console.error(`❌ Ошибка записи свойства "${propertyKey}":`, error.message);
    throw error;
  }
}

/**
 * МАССОВАЯ УСТАНОВКА НАСТРОЕК
 * 
 * Устанавливает несколько настроек одновременно
 * 
 * @param {Object} settingsObject - Объект с настройками
 * @example
 * setMultipleSettings({
 *   [SCRIPT_PROPERTIES_KEYS.INSALES_API_KEY]: 'key1',
 *   [SCRIPT_PROPERTIES_KEYS.OPENAI_API_KEY]: 'key2'
 * });
 */
function setMultipleSettings(settingsObject) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties(settingsObject);
    console.log(`✅ Установлено ${Object.keys(settingsObject).length} настроек`);
    
  } catch (error) {
    console.error('❌ Ошибка массовой записи настроек:', error.message);
    throw error;
  }
}

// =============================================================================
// ✅ ФУНКЦИИ ВАЛИДАЦИИ НАСТРОЕК
// =============================================================================

/**
 * ВАЛИДАЦИЯ КОНФИГУРАЦИИ ПРОЕКТА
 * 
 * Проверяет корректность всех критически важных настроек
 * 
 * @returns {Object} Результат валидации с детальной информацией
 * @example
 * const validation = validateConfig();
 * if (!validation.isValid) {
 *   console.log('Ошибки:', validation.errors);
 * }
 */
function validateConfig() {
  console.log('🔍 Начинаем валидацию конфигурации проекта...');
  
  const validation = {
    isValid: true,
    errors: [],
    warnings: [],
    checkedItems: []
  };
  
  try {
    // Проверяем существование рабочего листа
    console.log('📊 Проверяем существование рабочего листа...');
    
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.IMAGES);
      if (sheet) {
        validation.checkedItems.push(`✅ Лист "${SHEET_NAMES.IMAGES}" существует`);
      } else {
        validation.errors.push(`❌ Лист "${SHEET_NAMES.IMAGES}" не найден`);
        validation.isValid = false;
      }
    } catch (error) {
      validation.errors.push(`❌ Ошибка проверки листа "${SHEET_NAMES.IMAGES}": ${error.message}`);
      validation.isValid = false;
    }
    
    // Проверяем критически важные настройки API
    console.log('🔑 Проверяем настройки API...');
    
    try {
      const settings = getApiSettings();
      
      // Критически важные настройки InSales
      if (!settings.insalesApiKey) {
        validation.errors.push('❌ Не настроен InSales API Key');
        validation.isValid = false;
      } else {
        validation.checkedItems.push('✅ InSales API Key настроен');
      }
      
      if (!settings.insalesPassword) {
        validation.errors.push('❌ Не настроен InSales Password');
        validation.isValid = false;
      } else {
        validation.checkedItems.push('✅ InSales Password настроен');
      }
      
      if (!settings.insalesShop) {
        validation.errors.push('❌ Не настроен InSales Shop');
        validation.isValid = false;
      } else {
        validation.checkedItems.push('✅ InSales Shop настроен');
      }
      
      // Предупреждения о дополнительных API
      if (!settings.replicateToken) {
        validation.warnings.push('⚠️ Не настроен Replicate Token (ИИ-увеличение будет недоступно)');
      } else {
        validation.checkedItems.push('✅ Replicate Token настроен');
      }
      
      if (!settings.tinypngKey) {
        validation.warnings.push('⚠️ Не настроен TinyPNG Key (сжатие будет недоступно)');
      } else {
        validation.checkedItems.push('✅ TinyPNG Key настроен');
      }
      
      if (!settings.imgbbKey) {
        validation.warnings.push('⚠️ Не настроен ImgBB Key (загрузка на хостинг будет недоступна)');
      } else {
        validation.checkedItems.push('✅ ImgBB Key настроен');
      }
      
      if (!settings.openaiApiKey) {
        validation.warnings.push('⚠️ Не настроен OpenAI API Key (AI-генерация alt-тегов будет недоступна)');
      } else {
        validation.checkedItems.push('✅ OpenAI API Key настроен');
        
        // Дополнительная проверка ассистента для alt-тегов с анализом изображений
        if (settings.openaiAltAssistantId) {
          validation.checkedItems.push('✅ OpenAI Alt Assistant ID настроен (будет использоваться анализ изображений для генерации alt-тегов)');
          
          if (settings.aiImageAnalysis) {
            validation.checkedItems.push('✅ Анализ изображений включен (максимальная точность alt-тегов)');
          } else {
            validation.warnings.push('⚠️ Анализ изображений отключен (alt-теги будут генерироваться только по названию)');
          }
        } else {
          validation.warnings.push('⚠️ OpenAI Alt Assistant ID не настроен (будет использоваться обычный Chat API без анализа изображений)');
        }
      }
      
    } catch (settingsError) {
      validation.errors.push('❌ Ошибка чтения настроек API: ' + settingsError.message);
      validation.isValid = false;
    }
    
    // Итоговая информация
    console.log(`📋 Валидация завершена:`);
    console.log(`✅ Проверок пройдено: ${validation.checkedItems.length}`);
    console.log(`❌ Ошибок найдено: ${validation.errors.length}`);
    console.log(`⚠️ Предупреждений: ${validation.warnings.length}`);
    
    return validation;
    
  } catch (error) {
    validation.errors.push('❌ Критическая ошибка валидации: ' + error.message);
    validation.isValid = false;
    console.error('❌ Критическая ошибка валидации:', error.message);
    return validation;
  }
}

/**
 * БЫСТРАЯ ПРОВЕРКА ГОТОВНОСТИ К РАБОТЕ
 * 
 * Упрощенная проверка основных требований
 * 
 * @returns {boolean} true если система готова к работе
 */
function isSystemReady() {
  try {
    const validation = validateConfig();
    
    if (validation.isValid) {
      console.log('✅ Система готова к работе');
      return true;
    } else {
      console.log('❌ Система не готова к работе. Ошибки:', validation.errors);
      return false;
    }
    
  } catch (error) {
    console.error('❌ Ошибка проверки готовности системы:', error.message);
    return false;
  }
}

// =============================================================================
// 🛠️ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ПОЛУЧЕНИЯ ЛИСТОВ
// =============================================================================

/**
 * ПОЛУЧЕНИЕ ЛИСТА "НАСТРОЙКИ"
 * 
 * Безопасное получение листа с настройками
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Лист настроек
 * @throws {Error} Если лист не найден
 */
function getSettingsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!sheet) {
    throw new Error(`Лист "${SHEET_NAMES.SETTINGS}" не найден. Создайте структуру листов сначала.`);
  }
  
  return sheet;
}

/**
 * ПОЛУЧЕНИЕ ЛИСТА "ОБРАБОТКА ИЗОБРАЖЕНИЙ"
 * 
 * Безопасное получение основного рабочего листа
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Лист обработки изображений
 * @throws {Error} Если лист не найден
 */
function getImagesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.IMAGES);
  
  if (!sheet) {
    throw new Error(`Лист "${SHEET_NAMES.IMAGES}" не найден. Создайте структуру листов сначала.`);
  }
  
  return sheet;
}

/**
 * ПОЛУЧЕНИЕ ЛИСТА "СПРАВОЧНИК ТОВАРОВ"
 * 
 * Безопасное получение листа со справочником товаров
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Лист справочника товаров
 * @throws {Error} Если лист не найден
 */
function getProductsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PRODUCTS);
  
  if (!sheet) {
    throw new Error(`Лист "${SHEET_NAMES.PRODUCTS}" не найден. Создайте структуру листов сначала.`);
  }
  
  return sheet;
}

/**
 * ПОЛУЧЕНИЕ ЛЮБОГО ЛИСТА ПО НАЗВАНИЮ
 * 
 * Универсальная функция получения листа с проверкой существования
 * 
 * @param {string} sheetName - Название листа
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Найденный лист
 * @throws {Error} Если лист не найден
 */
function getSheetByName(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Лист "${sheetName}" не найден`);
  }
  
  return sheet;
}

// =============================================================================
// 📋 ИНФОРМАЦИОННЫЕ ФУНКЦИИ
// =============================================================================

/**
 * ПОЛУЧЕНИЕ ИНФОРМАЦИИ О ПРОЕКТЕ
 * 
 * Возвращает основную информацию о конфигурации проекта
 * 
 * @returns {Object} Информация о проекте
 */
function getProjectInfo() {
  return {
    projectName: 'Обработка изображений товаров',
    version: '1.0',
    moduleVersion: '00_config.js v1.0',
    sheetsCount: Object.keys(SHEET_NAMES).length,
    columnsCount: Object.keys(IMAGES_COLUMNS).length,
    settingsCount: Object.keys(SETTINGS_PARAMS).length,
    author: 'AI Assistant + User',
    created: new Date().toLocaleString('ru-RU'),
    description: 'Система автоматической обработки изображений товаров с интеграцией InSales'
  };
}

/**
 * ВЫВОД СПРАВОЧНОЙ ИНФОРМАЦИИ В КОНСОЛЬ
 * 
 * Показывает всю конфигурацию проекта в удобном виде
 */
function showConfigInfo() {
  const info = getProjectInfo();
  
  console.log('='.repeat(50));
  console.log(`📋 КОНФИГУРАЦИЯ ПРОЕКТА "${info.projectName}"`);
  console.log('='.repeat(50));
  
  console.log('\n📊 ЛИСТЫ GOOGLE SHEETS:');
  for (const [key, name] of Object.entries(SHEET_NAMES)) {
    console.log(`  ${key}: "${name}"`);
  }
  
  console.log('\n📋 КОЛОНКИ ОСНОВНОГО ЛИСТА:');
  for (const [key, column] of Object.entries(IMAGES_COLUMNS)) {
    const letter = String.fromCharCode(64 + column); // Преобразуем номер в букву
    console.log(`  ${letter}${column}: ${key}`);
  }
  
  console.log('\n⚙️ НАСТРОЙКИ API:');
  for (const [key, param] of Object.entries(SETTINGS_PARAMS)) {
    console.log(`  ${key}: "${param}"`);
  }
  
  console.log('\n📈 СТАТИСТИКА:');
  console.log(`  Листов: ${info.sheetsCount}`);
  console.log(`  Колонок в основном листе: ${info.columnsCount}`);
  console.log(`  Параметров настроек: ${info.settingsCount}`);
  
  console.log('\n' + '='.repeat(50));
}

// =============================================================================
// 🧪 ФУНКЦИИ ТЕСТИРОВАНИЯ МОДУЛЯ
// =============================================================================

/**
 * ТЕСТИРОВАНИЕ МОДУЛЯ CONFIG
 * 
 * Комплексное тестирование всех функций модуля
 * Используется для проверки работоспособности после создания или изменения
 */
function testConfigModule() {
  console.log('🧪 Начинаем тестирование модуля 00_config.gs...');
  
  try {
    // Тест 1: Показ информации о проекте
    console.log('\n📋 Тест 1: Информация о проекте');
    const projectInfo = getProjectInfo();
    console.log(`✅ Проект: ${projectInfo.projectName} v${projectInfo.version}`);
    
    // Тест 2: Проверка констант
    console.log('\n📊 Тест 2: Проверка констант');
    console.log(`✅ Листов настроено: ${Object.keys(SHEET_NAMES).length}`);
    console.log(`✅ Колонок в основном листе: ${Object.keys(IMAGES_COLUMNS).length}`);
    console.log(`✅ Параметров настроек: ${Object.keys(SETTINGS_PARAMS).length}`);
    
    // Тест 3: Валидация конфигурации
    console.log('\n🔍 Тест 3: Валидация конфигурации');
    const validation = validateConfig();
    console.log(`✅ Валидация выполнена. Статус: ${validation.isValid ? 'OK' : 'ОШИБКИ'}`);
    
    if (validation.errors.length > 0) {
      console.log('❌ Найденные ошибки:');
      validation.errors.forEach(error => console.log(`   ${error}`));
    }
    
    if (validation.warnings.length > 0) {
      console.log('⚠️ Предупреждения:');
      validation.warnings.forEach(warning => console.log(`   ${warning}`));
    }
    
    // Тест 4: Проверка готовности системы
    console.log('\n🚀 Тест 4: Готовность системы');
    const isReady = isSystemReady();
    console.log(`✅ Система готова к работе: ${isReady ? 'ДА' : 'НЕТ'}`);
    
    console.log('\n🎉 Тестирование модуля 00_config.gs завершено!');
    console.log('📋 Модуль готов к использованию в проекте');
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании модуля config:', error.message);
    throw error;
  }
}

/**
 * Обновление данных существующего товара по артикулу
 * 
 * Назначение: Изменяет данные товара в листе без создания дубликатов
 * Логика работы:
 * 1. Находит строку товара по артикулу
 * 2. Обновляет только указанные поля
 * 3. Сохраняет существующие статусы и комментарии
 * 4. Логирует изменения
 * 
 * @param {string} article - Артикул товара для поиска
 * @param {Object} updateData - Объект с полями для обновления
 * @returns {boolean} true если обновление успешно, false при ошибке
 */
function updateProductData(article, updateData) {
  try {
    logInfo(`🔄 Обновляем данные товара: ${article}`, {
      fieldsToUpdate: Object.keys(updateData)
    });
    
    // Получаем рабочий лист
    const sheet = getWorkingSheet();
    if (!sheet) {
      throw new Error('Не удалось получить рабочий лист');
    }
    
    // Получаем заголовки и данные
    const headers = getSheetHeaders();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      logWarning('⚠️ Лист пуст, нечего обновлять');
      return false;
    }
    
    // Находим индекс колонки артикула
    const articleColumnIndex = headers.indexOf('article');
    if (articleColumnIndex === -1) {
      throw new Error('Колонка "article" не найдена в заголовках');
    }
    
    // Ищем строку с нужным артикулом
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) { // Начинаем с 1, пропуская заголовки
      if (data[i][articleColumnIndex] === article) {
        targetRow = i + 1; // +1 для номера строки в Google Sheets (1-based)
        break;
      }
    }
    
    if (targetRow === -1) {
      logWarning(`⚠️ Товар с артикулом "${article}" не найден в листе`);
      return false;
    }
    
    // Обновляем поля
    let updatedFields = [];
    Object.entries(updateData).forEach(([fieldName, newValue]) => {
      const columnIndex = headers.indexOf(fieldName);
      
      if (columnIndex === -1) {
        logWarning(`⚠️ Колонка "${fieldName}" не найдена, пропускаем`);
        return;
      }
      
      // Получаем текущее значение
      const currentValue = data[targetRow - 1][columnIndex];
      
      // Обновляем только если значение изменилось
      if (currentValue !== newValue) {
        sheet.getRange(targetRow, columnIndex + 1).setValue(newValue);
        updatedFields.push({
          field: fieldName,
          oldValue: currentValue,
          newValue: newValue
        });
        
        logInfo(`📝 Обновлено поле "${fieldName}": "${currentValue}" → "${newValue}"`);
      }
    });
    
    if (updatedFields.length === 0) {
      logInfo(`ℹ️ Товар "${article}" уже актуален, изменений не требуется`);
      return true;
    }
    
    // Обновляем timestamp последнего изменения (если есть такая колонка)
    const lastUpdatedIndex = headers.indexOf('lastUpdated');
    if (lastUpdatedIndex !== -1) {
      sheet.getRange(targetRow, lastUpdatedIndex + 1).setValue(new Date().toISOString());
    }
    
    logInfo(`✅ Товар "${article}" успешно обновлен`, {
      rowNumber: targetRow,
      updatedFieldsCount: updatedFields.length,
      updatedFields: updatedFields
    });
    
    return true;
    
  } catch (error) {
    handleError(error, 'Обновление данных товара', {
      function: 'updateProductData',
      article: article,
      updateData: updateData
    });
    return false;
  }
}

/**
 * ===================================================================
 * 💡 ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ МОДУЛЯ
 * ===================================================================
 * 
 * ОСНОВНЫЕ ФУНКЦИИ ДЛЯ ДРУГИХ МОДУЛЕЙ:
 * 
 * 1. getApiSettings() - получить все настройки API
 * 2. getSetting(paramName) - получить конкретную настройку
 * 3. validateConfig() - проверить корректность настроек
 * 4. getImagesSheet() - получить основной рабочий лист
 * 5. getSettingsSheet() - получить лист настроек
 * 
 * КОНСТАНТЫ ДЛЯ ИСПОЛЬЗОВАНИЯ:
 * 
 * - SHEET_NAMES.IMAGES - название основного листа
 * - IMAGES_COLUMNS.ARTICLE - колонка с артикулами
 * - STATUS_VALUES.PROCESSING.COMPLETED - статус "Обработано"
 * - SETTINGS_PARAMS.INSALES_API_KEY - параметр API ключа
 * 
 * ПРИМЕР ИСПОЛЬЗОВАНИЯ В ДРУГИХ МОДУЛЯХ:
 * 
 * // Получение настроек
 * const settings = getApiSettings();
 * console.log(settings.insalesApiKey);
 * 
 * // Работа с листами
 * const sheet = getImagesSheet();
 * const data = sheet.getRange(2, IMAGES_COLUMNS.ARTICLE, 10, 1).getValues();
 * 
 * // Проверка готовности
 * if (!isSystemReady()) {
 *   throw new Error('Система не настроена');
 * }
 * 
 * ===================================================================
 */