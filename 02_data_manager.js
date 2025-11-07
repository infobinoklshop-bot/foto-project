/**
 * ===================================================================
 * МОДУЛЬ 02_data_manager.gs - УПРАВЛЕНИЕ ДАННЫМИ
 * ===================================================================
 * 
 * Назначение: Управление структурой и данными основного рабочего листа
 * Версия: 1.0
 * Проект: Обработка изображений товаров
 * 
 * Этот модуль содержит:
 * - Создание структуры основного листа
 * - Функции чтения/записи данных
 * - Управление форматированием
 * - Валидация структуры листа
 */

// =============================================================================
// 🏗️ СОЗДАНИЕ И НАСТРОЙКА СТРУКТУРЫ ЛИСТА
// =============================================================================

/**
 * СОЗДАНИЕ ОСНОВНОГО ЛИСТА "ОБРАБОТКА ИЗОБРАЖЕНИЙ"
 * 
 * Создает единственный рабочий лист с полной настройкой структуры
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Созданный лист
 * @throws {Error} Если создание не удалось
 * @example
 * const sheet = createImagesSheet();
 */
function createImagesSheet() {
  try {
    logInfo('🏗️ Начинаем создание основного листа "Обработка изображений"...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Проверяем, существует ли лист уже
    const existingSheet = spreadsheet.getSheetByName(SHEET_NAMES.IMAGES);
    if (existingSheet) {
      logWarning(`Лист "${SHEET_NAMES.IMAGES}" уже существует`);
      
      // Спрашиваем пользователя через UI (если возможно) или просто предупреждаем
      logInfo('Настраиваем существующий лист...');
      setupSheetStructure(existingSheet);
      return existingSheet;
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(SHEET_NAMES.IMAGES);
    logInfo(`✅ Лист "${SHEET_NAMES.IMAGES}" создан`);
    
    // Настраиваем структуру
    setupSheetStructure(sheet);
    
    logInfo('🎉 Основной лист успешно создан и настроен!');
    return sheet;
    
  } catch (error) {
    const errorMsg = `Ошибка создания листа "${SHEET_NAMES.IMAGES}": ${error.message}`;
    logCritical(errorMsg);
    handleError(error, 'Создание основного листа');
    throw new Error(errorMsg);
  }
}

/**
 * НАСТРОЙКА СТРУКТУРЫ ЛИСТА
 * 
 * Полная настройка заголовков, форматирования и защиты
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист для настройки
 * @throws {Error} Если настройка не удалась
 */
function setupSheetStructure(sheet) {
  try {
    logInfo('⚙️ Настраиваем структуру листа...');
    
    if (!sheet) {
      throw new Error('Лист не передан для настройки');
    }
    
    // 1. Создаем заголовки
    setupHeaders(sheet);
    
    // 2. Настраиваем форматирование
    setupFormatting(sheet);
    
    // 3. Настраиваем ширину колонок
    setupColumnWidths(sheet);
    
    // 4. Добавляем чекбоксы в колонку A
    setupCheckboxes(sheet);
    
    // 5. Защищаем заголовки
    protectHeaders(sheet);
    
    // 6. Закрепляем первую строку
    sheet.setFrozenRows(1);
    
    logInfo('✅ Структура листа настроена успешно');
    
  } catch (error) {
    const errorMsg = `Ошибка настройки структуры листа: ${error.message}`;
    logError(errorMsg);
    handleError(error, 'Настройка структуры листа');
    throw new Error(errorMsg);
  }
}

/**
 * СОЗДАНИЕ ЗАГОЛОВКОВ КОЛОНОК
 * 
 * Устанавливает заголовки согласно структуре IMAGES_COLUMNS
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 */
function setupHeaders(sheet) {
  try {
    logInfo('📋 Создаем заголовки колонок...');
   
    const headers = [
      '☑️ Обработать',           // A - CHECKBOX
      'Артикул',                // B - ARTICLE
      'ID InSales',             // C - INSALES_ID
      'Название товара',         // D - PRODUCT_NAME
      'Исходные изображения',    // E - ORIGINAL_IMAGES
      'Парсинг Поставщика',     // F - SUPPLIER_IMAGES
      'Дополнительные фото',    // G - ADDITIONAL_IMAGES
      'Обработанные изображения', // H - PROCESSED_IMAGES
      'Alt-теги',               // I - ALT_TAGS
      'SEO имена файлов',       // J - SEO_FILENAMES
      'Статус обработки',       // K - PROCESSING_STATUS
      'Статус InSales',         // L - INSALES_STATUS
      // === НОВЫЕ КОЛОНКИ ДЛЯ ИМПОРТА ===
      'Описание поставщика',    // M - DESCRIPTION
      'Описание (рерайт AI)',   // N - DESCRIPTION_REWRITTEN
      'Краткое описание',       // O - SHORT_DESCRIPTION
      'Характеристики (сырые)', // P - SPECIFICATIONS_RAW
      'Характеристики (норм.)', // Q - SPECIFICATIONS_NORMALIZED
      'Цена',                   // R - PRICE
      'Остаток',                // S - STOCK
      'Категории',              // T - CATEGORIES
      'Бренд',                  // U - BRAND
      'Серия',                  // V - SERIES
      'Вес, г',                 // W - WEIGHT
      'Габариты',               // X - DIMENSIONS
      'Комплектация',           // Y - PACKAGE_CONTENTS
      'Статус сопоставления',   // Z - MATCH_STATUS
      'Совпадение, %',          // AA - MATCH_CONFIDENCE
      'Статус импорта'          // AB - IMPORT_STATUS
    ];
    
    // Записываем заголовки в первую строку
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    logInfo(`✅ Заголовки созданы: ${headers.length} колонок`);
    
  } catch (error) {
    logError('Ошибка создания заголовков: ' + error.message);
    throw error;
  }
}

/**
 * НАСТРОЙКА ФОРМАТИРОВАНИЯ ЛИСТА
 * 
 * Применяет стили к заголовкам и основным областям листа
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 */
function setupFormatting(sheet) {
  try {
    logInfo('🎨 Настраиваем форматирование...');
    
    // Количество колонок для форматирования
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    
    // Форматирование заголовков (первая строка)
    const headerRange = sheet.getRange(1, 1, 1, numColumns);
    
    headerRange
      .setBackground('#4285f4')           // Синий фон (Google Blue)
      .setFontColor('#ffffff')            // Белый текст
      .setFontWeight('bold')              // Жирный шрифт
      .setFontSize(11)                    // Размер шрифта
      .setHorizontalAlignment('center')    // Выравнивание по центру
      .setVerticalAlignment('middle')      // Вертикальное выравнивание по центру
      .setWrap(true);                     // Перенос текста
    
    // Устанавливаем высоту строки заголовков
    sheet.setRowHeight(1, 40);
    
    // Форматирование области данных (начиная со второй строки)
    // Создаем несколько строк для примера данных
    const dataStartRow = 2;
    const initialDataRows = 50; // Создаем 50 строк для начала
    
    if (sheet.getMaxRows() < dataStartRow + initialDataRows) {
      sheet.insertRows(sheet.getMaxRows(), (dataStartRow + initialDataRows) - sheet.getMaxRows());
    }
    
    const dataRange = sheet.getRange(dataStartRow, 1, initialDataRows, numColumns);
    
    dataRange
      .setBackground('#ffffff')           // Белый фон для данных
      .setFontColor('#000000')            // Черный текст
      .setFontSize(10)                    // Размер шрифта для данных
      .setVerticalAlignment('top')        // Выравнивание по верху
      .setWrap(false);                    // Без переноса для данных
    
    // Добавляем границы
    const fullRange = sheet.getRange(1, 1, dataStartRow + initialDataRows - 1, numColumns);
    fullRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Чередующиеся цвета строк для лучшей читаемости
    for (let row = dataStartRow; row < dataStartRow + initialDataRows; row += 2) {
      const evenRowRange = sheet.getRange(row, 1, 1, numColumns);
      evenRowRange.setBackground('#f8f9fa'); // Светло-серый для четных строк
    }
    
    logInfo('✅ Форматирование применено');
    
  } catch (error) {
    logError('Ошибка настройки форматирования: ' + error.message);
    throw error;
  }
}

/**
 * НАСТРОЙКА ШИРИНЫ КОЛОНОК
 * 
 * Устанавливает оптимальную ширину для каждой колонки
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 */
function setupColumnWidths(sheet) {
  try {
    logInfo('📏 Настраиваем ширину колонок...');
    
    // Определяем ширину для каждой колонки (в пикселях)
    const columnWidths = {
      [IMAGES_COLUMNS.CHECKBOX]: 80,           // A - Чекбокс
      [IMAGES_COLUMNS.ARTICLE]: 120,           // B - Артикул
      [IMAGES_COLUMNS.INSALES_ID]: 100,        // C - ID InSales
      [IMAGES_COLUMNS.PRODUCT_NAME]: 200,      // D - Название товара
      [IMAGES_COLUMNS.ORIGINAL_IMAGES]: 150,   // E - Исходные изображения
      [IMAGES_COLUMNS.PROCESSED_IMAGES]: 150,  // F - Обработанные изображения
      [IMAGES_COLUMNS.ALT_TAGS]: 180,          // G - Alt-теги
      [IMAGES_COLUMNS.SEO_FILENAMES]: 180,     // H - SEO имена файлов
      [IMAGES_COLUMNS.PROCESSING_STATUS]: 130, // I - Статус обработки
      [IMAGES_COLUMNS.INSALES_STATUS]: 130     // J - Статус InSales
    };
    
    // Применяем ширину для каждой колонки
    for (const [columnKey, columnIndex] of Object.entries(IMAGES_COLUMNS)) {
      const width = columnWidths[columnIndex];
      if (width) {
        sheet.setColumnWidth(columnIndex, width);
      }
    }
    
    logInfo('✅ Ширина колонок настроена');
    
  } catch (error) {
    logError('Ошибка настройки ширины колонок: ' + error.message);
    throw error;
  }
}

/**
 * НАСТРОЙКА ЧЕКБОКСОВ В КОЛОНКЕ A
 * 
 * Добавляет чекбоксы для выбора товаров к обработке
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 */
function setupCheckboxes(sheet) {
  try {
    logInfo('☑️ Настраиваем чекбоксы...');
    
    // Настраиваем чекбоксы в колонке A начиная со второй строки
    const startRow = 2;
    const numRows = 50; // Начальное количество строк с чекбоксами
    
    const checkboxRange = sheet.getRange(startRow, IMAGES_COLUMNS.CHECKBOX, numRows, 1);
    
    // Добавляем валидацию данных - чекбоксы
    const checkboxValidation = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    
    checkboxRange.setDataValidation(checkboxValidation);
    
    // Устанавливаем значения по умолчанию (false - не отмечены)
    const defaultValues = Array(numRows).fill([false]);
    checkboxRange.setValues(defaultValues);
    
    // Выравниваем чекбоксы по центру
    checkboxRange.setHorizontalAlignment('center');
    
    logInfo(`✅ Чекбоксы настроены для ${numRows} строк`);
    
  } catch (error) {
    logError('Ошибка настройки чекбоксов: ' + error.message);
    throw error;
  }
}

/**
 * ЗАЩИТА ЗАГОЛОВКОВ ОТ РЕДАКТИРОВАНИЯ
 * 
 * Защищает первую строку с заголовками от случайного изменения
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 */
function protectHeaders(sheet) {
  try {
    logInfo('🔒 Защищаем заголовки от редактирования...');
    
    // Получаем диапазон заголовков (первая строка)
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    const headerRange = sheet.getRange(1, 1, 1, numColumns);
    
    // Создаем защиту для заголовков
    const protection = headerRange.protect();
    
    // Настраиваем описание защиты
    protection.setDescription('Заголовки таблицы (защищены от редактирования)');
    
    // Удаляем всех редакторов кроме владельца (если есть права)
    try {
      const me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
    } catch (permissionError) {
      logWarning('Не удалось настроить права доступа к защите: ' + permissionError.message);
    }
    
    logInfo('✅ Заголовки защищены от редактирования');
    
  } catch (error) {
    logWarning('Не удалось защитить заголовки: ' + error.message);
    // Не останавливаем выполнение, так как это не критично
  }
}

// =============================================================================
// 📊 ФУНКЦИИ ЧТЕНИЯ ДАННЫХ
// =============================================================================

/**
 * ЧТЕНИЕ ВСЕХ ДАННЫХ ИЗ ЛИСТА "ОБРАБОТКА ИЗОБРАЖЕНИЙ"
 * 
 * Читает все данные из основного листа, исключая заголовки
 * 
 * @returns {Array<Object>} Массив объектов с данными товаров
 * @throws {Error} Если чтение не удалось
 * @example
 * const data = readImagesData();
 * console.log(data[0].article); // Артикул первого товара
 */
function readImagesData() {
  try {
    logInfo('📖 Читаем данные из листа "Обработка изображений"...');
    
    // Получаем рабочий лист
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    
    // Определяем диапазон данных (исключая заголовки)
    const lastRow = sheet.getLastRow();
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    
    if (lastRow <= 1) {
      logInfo('📋 Лист пуст, возвращаем пустой массив');
      return [];
    }
    
    // Читаем данные начиная со второй строки
    const dataRange = sheet.getRange(2, 1, lastRow - 1, numColumns);
    const rawData = dataRange.getValues();
    
    // Преобразуем сырые данные в структурированные объекты
    const structuredData = rawData.map((row, index) => {
      const rowNumber = index + 2; // +2 потому что начинаем с второй строки
      
      return {
        rowNumber: rowNumber,
        checkbox: Boolean(row[IMAGES_COLUMNS.CHECKBOX - 1]),
        article: cleanString(row[IMAGES_COLUMNS.ARTICLE - 1] || ''),
        insalesId: cleanString(row[IMAGES_COLUMNS.INSALES_ID - 1] || ''),
        productName: cleanString(row[IMAGES_COLUMNS.PRODUCT_NAME - 1] || ''),
        originalImages: cleanString(row[IMAGES_COLUMNS.ORIGINAL_IMAGES - 1] || ''),
        processedImages: cleanString(row[IMAGES_COLUMNS.PROCESSED_IMAGES - 1] || ''),
        altTags: cleanString(row[IMAGES_COLUMNS.ALT_TAGS - 1] || ''),
        seoFilenames: cleanString(row[IMAGES_COLUMNS.SEO_FILENAMES - 1] || ''),
        processingStatus: cleanString(row[IMAGES_COLUMNS.PROCESSING_STATUS - 1] || ''),
        insalesStatus: cleanString(row[IMAGES_COLUMNS.INSALES_STATUS - 1] || '')
      };
    });
    
    // Фильтруем пустые строки (где нет артикула)
    const filteredData = structuredData.filter(item => item.article !== '');
    
    logInfo(`✅ Прочитано ${filteredData.length} записей товаров`);
    return filteredData;
    
  } catch (error) {
    const errorMsg = `Ошибка чтения данных из листа: ${error.message}`;
    logError(errorMsg);
    handleError(error, 'Чтение данных листа');
    throw new Error(errorMsg);
  }
}

/**
 * ЧТЕНИЕ ОТМЕЧЕННЫХ ТОВАРОВ
 * 
 * Возвращает только те товары, у которых установлен чекбокс
 * 
 * @returns {Array<Object>} Массив отмеченных товаров
 * @example
 * const selectedProducts = readSelectedProducts();
 */
function readSelectedProducts() {
  try {
    logInfo('☑️ Читаем отмеченные товары...');
    
    const allData = readImagesData();
    const selectedData = allData.filter(item => item.checkbox === true);
    
    logInfo(`✅ Найдено ${selectedData.length} отмеченных товаров из ${allData.length}`);
    return selectedData;
    
  } catch (error) {
    logError('Ошибка чтения отмеченных товаров: ' + error.message);
    throw error;
  }
}

/**
 * ПОИСК ТОВАРА ПО АРТИКУЛУ
 * 
 * Находит товар в листе по артикулу
 * 
 * @param {string} article - Артикул товара для поиска
 * @returns {Object|null} Данные товара или null если не найден
 * @example
 * const product = findProductByArticle('ART123');
 */
function findProductByArticle(article) {
  try {
    if (!article || typeof article !== 'string') {
      throw new Error('Артикул не указан или имеет неверный формат');
    }
    
    const cleanArticle = cleanString(article);
    logDebug(`🔍 Ищем товар с артикулом "${cleanArticle}"`);
    
    const allData = readImagesData();
    const foundProduct = allData.find(item => item.article === cleanArticle);
    
    if (foundProduct) {
      logDebug(`✅ Товар найден в строке ${foundProduct.rowNumber}`);
    } else {
      logDebug(`❌ Товар с артикулом "${cleanArticle}" не найден`);
    }
    
    return foundProduct || null;
    
  } catch (error) {
    logError(`Ошибка поиска товара по артикулу "${article}": ${error.message}`);
    return null;
  }
}

// =============================================================================
// ✏️ ФУНКЦИИ ЗАПИСИ ДАННЫХ
// =============================================================================

/**
 * ЗАПИСЬ ДАННЫХ ТОВАРА В ЛИСТ
 * 
 * Записывает или обновляет данные одного товара
 * 
 * @param {Object} productData - Данные товара для записи
 * @param {number} targetRow - Номер строки для записи (опционально)
 * @returns {number} Номер строки где записаны данные
 * @throws {Error} Если запись не удалась
 * @example
 * const newProduct = {
 *   article: 'ART123',
 *   productName: 'Название товара',
 *   originalImages: 'https://example.com/image.jpg'
 * };
 * const rowNumber = writeProductData(newProduct);
 */
function writeProductData(productData, targetRow = null) {
  try {
    logInfo('✏️ Записываем данные товара...');
    
    // Валидация входных данных
    if (!productData || typeof productData !== 'object') {
      throw new Error('Данные товара не переданы или имеют неверный формат');
    }
    
    if (!productData.article) {
      throw new Error('Артикул товара обязателен для записи');
    }
    
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    
    // Определяем строку для записи
    let rowToWrite = targetRow;
    
    if (!rowToWrite) {
      // Ищем существующую запись по артикулу
      const existingProduct = findProductByArticle(productData.article);
      
      if (existingProduct) {
        rowToWrite = existingProduct.rowNumber;
        logInfo(`📝 Обновляем существующий товар в строке ${rowToWrite}`);
      } else {
        // Находим первую пустую строку
        rowToWrite = sheet.getLastRow() + 1;
        logInfo(`➕ Добавляем новый товар в строку ${rowToWrite}`);
      }
    }
    
    // Подготавливаем данные для записи
    const rowData = [
      productData.checkbox !== undefined ? Boolean(productData.checkbox) : false,
      cleanString(productData.article || ''),
      cleanString(productData.insalesId || ''),
      cleanString(productData.productName || ''),
      cleanString(productData.originalImages || ''),
      cleanString(productData.processedImages || ''),
      cleanString(productData.altTags || ''),
      cleanString(productData.seoFilenames || ''),
      cleanString(productData.processingStatus || ''),
      cleanString(productData.insalesStatus || '')
    ];
    
    // Записываем данные в строку
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    const targetRange = sheet.getRange(rowToWrite, 1, 1, numColumns);
    targetRange.setValues([rowData]);
    
    // Если это новая строка, настраиваем форматирование
    if (rowToWrite > sheet.getLastRow() - 1) {
      setupRowFormatting(sheet, rowToWrite);
      
      // Добавляем чекбокс если это новая строка
      const checkboxRange = sheet.getRange(rowToWrite, IMAGES_COLUMNS.CHECKBOX, 1, 1);
      const checkboxValidation = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .build();
      checkboxRange.setDataValidation(checkboxValidation);
    }
    
    logInfo(`✅ Данные товара "${productData.article}" записаны в строку ${rowToWrite}`);
    return rowToWrite;
    
  } catch (error) {
    const errorMsg = `Ошибка записи данных товара: ${error.message}`;
    logError(errorMsg);
    handleError(error, 'Запись данных товара', {article: productData?.article});
    throw new Error(errorMsg);
  }
}

/**
 * ОБНОВЛЕНИЕ СТАТУСА ТОВАРА
 * 
 * Обновляет статус обработки или InSales для конкретного товара
 * 
 * @param {string} article - Артикул товара
 * @param {string} statusType - Тип статуса ('processing' или 'insales')
 * @param {string} statusValue - Новое значение статуса
 * @example
 * updateProductStatus('ART123', 'processing', 'Обработано');
 */
function updateProductStatus(article, statusType, statusValue) {
  try {
    logInfo(`🔄 Обновляем статус "${statusType}" для товара "${article}"...`);
    
    // Валидация параметров
    if (!article || !statusType || statusValue === undefined) {
      throw new Error('Не все параметры указаны для обновления статуса');
    }
    
    // Находим товар
    const product = findProductByArticle(article);
    if (!product) {
      throw new Error(`Товар с артикулом "${article}" не найден`);
    }
    
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    
    // Определяем колонку для обновления
    let columnToUpdate;
    switch (statusType.toLowerCase()) {
      case 'processing':
        columnToUpdate = IMAGES_COLUMNS.PROCESSING_STATUS;
        break;
      case 'insales':
        columnToUpdate = IMAGES_COLUMNS.INSALES_STATUS;
        break;
      default:
        throw new Error(`Неизвестный тип статуса: "${statusType}"`);
    }
    
    // Обновляем статус
    const targetCell = sheet.getRange(product.rowNumber, columnToUpdate);
    targetCell.setValue(cleanString(statusValue));
    
    logInfo(`✅ Статус "${statusType}" обновлен на "${statusValue}" для товара "${article}"`);
    
  } catch (error) {
    const errorMsg = `Ошибка обновления статуса: ${error.message}`;
    logError(errorMsg);
    handleError(error, 'Обновление статуса товара', {article, statusType, statusValue});
    throw new Error(errorMsg);
  }
}

/**
 * МАССОВАЯ ЗАПИСЬ ТОВАРОВ
 * 
 * Записывает массив товаров в лист одной операцией
 * 
 * @param {Array<Object>} productsArray - Массив данных товаров
 * @returns {Object} Результат операции
 * @example
 * const products = [
 *   {article: 'ART123', productName: 'Товар 1'},
 *   {article: 'ART124', productName: 'Товар 2'}
 * ];
 * const result = writeMultipleProducts(products);
 */
function writeMultipleProducts(productsArray) {
  try {
    logInfo(`📝 Начинаем массовую запись ${productsArray.length} товаров...`);
    
    if (!Array.isArray(productsArray) || productsArray.length === 0) {
      throw new Error('Массив товаров пуст или имеет неверный формат');
    }
    
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    const results = {
      successful: 0,
      failed: 0,
      errors: [],
      startRow: sheet.getLastRow() + 1
    };
    
    // Записываем каждый товар
    for (let i = 0; i < productsArray.length; i++) {
      try {
        const product = productsArray[i];
        writeProductData(product);
        results.successful++;
        
        // Логируем прогресс каждые 10 товаров
        if ((i + 1) % 10 === 0) {
          logInfo(`📊 Записано ${i + 1} из ${productsArray.length} товаров`);
        }
        
      } catch (productError) {
        results.failed++;
        results.errors.push({
          index: i,
          article: productsArray[i]?.article || 'Неизвестно',
          error: productError.message
        });
        logWarning(`Ошибка записи товара ${i + 1}: ${productError.message}`);
      }
    }
    
    logInfo(`✅ Массовая запись завершена. Успешно: ${results.successful}, Ошибок: ${results.failed}`);
    
    if (results.failed > 0) {
      logWarning(`⚠️ Некоторые товары не были записаны: ${JSON.stringify(results.errors, null, 2)}`);
    }
    
    return results;
    
  } catch (error) {
    const errorMsg = `Ошибка массовой записи товаров: ${error.message}`;
    logCritical(errorMsg);
    handleError(error, 'Массовая запись товаров');
    throw new Error(errorMsg);
  }
}

// =============================================================================
// 🎨 ФУНКЦИИ ФОРМАТИРОВАНИЯ
// =============================================================================

/**
 * НАСТРОЙКА ФОРМАТИРОВАНИЯ НОВОЙ СТРОКИ
 * 
 * Применяет стандартное форматирование к новой строке данных
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Рабочий лист
 * @param {number} rowNumber - Номер строки для форматирования
 */
function setupRowFormatting(sheet, rowNumber) {
  try {
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    const rowRange = sheet.getRange(rowNumber, 1, 1, numColumns);
    
    // Базовое форматирование
    rowRange
      .setFontSize(10)
      .setVerticalAlignment('top')
      .setWrap(false);
    
    // Чередующиеся цвета строк
    if (rowNumber % 2 === 0) {
      rowRange.setBackground('#f8f9fa'); // Светло-серый для четных строк
    } else {
      rowRange.setBackground('#ffffff'); // Белый для нечетных строк
    }
    
    // Границы
    rowRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Выравнивание чекбокса по центру
    const checkboxCell = sheet.getRange(rowNumber, IMAGES_COLUMNS.CHECKBOX);
    checkboxCell.setHorizontalAlignment('center');
    
  } catch (error) {
    logWarning(`Не удалось настроить форматирование строки ${rowNumber}: ${error.message}`);
  }
}

// =============================================================================
// 🧹 ФУНКЦИИ УПРАВЛЕНИЯ ДАННЫМИ
// =============================================================================

/**
 * ОЧИСТКА ДАННЫХ ЛИСТА
 * 
 * Очищает все данные кроме заголовков
 * 
 * @param {boolean} confirmClear - Подтверждение очистки
 * @returns {boolean} true если очистка выполнена
 * @example
 * const cleared = clearImagesData(true);
 */
function clearImagesData(confirmClear = false) {
  try {
    if (!confirmClear) {
      logWarning('Очистка данных отменена - не получено подтверждение');
      return false;
    }
    
    logInfo('🧹 Начинаем очистку данных листа...');
    
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      logInfo('📋 Лист уже пуст');
      return true;
    }
    
    // Очищаем данные начиная со второй строки
    const numColumns = Object.keys(IMAGES_COLUMNS).length;
    const dataRange = sheet.getRange(2, 1, lastRow - 1, numColumns);
    
    dataRange.clearContent();
    dataRange.clearDataValidations();
    
    // Удаляем лишние строки, оставляя заголовок и несколько строк для данных
    const rowsToKeep = 51; // Заголовок + 50 строк данных
    if (lastRow > rowsToKeep) {
      sheet.deleteRows(rowsToKeep + 1, lastRow - rowsToKeep);
    }
    
    // Восстанавливаем базовое форматирование и чекбоксы
    setupCheckboxes(sheet);
    
    logInfo('✅ Данные листа очищены');
    
    // Отправляем уведомление о важном действии
    sendNotification(`🧹 ОЧИСТКА ДАННЫХ
    
📋 Лист: ${SHEET_NAMES.IMAGES}
🕐 Время: ${formatDate(new Date(), 'full')}
⚠️ Все данные удалены (заголовки сохранены)`);
    
    return true;
    
  } catch (error) {
    const errorMsg = `Ошибка очистки данных листа: ${error.message}`;
    logCritical(errorMsg);
    handleError(error, 'Очистка данных листа');
    throw new Error(errorMsg);
  }
}

/**
 * СБРОС ВСЕХ ЧЕКБОКСОВ
 * 
 * Снимает отметки со всех чекбоксов в листе
 * 
 * @returns {number} Количество сброшенных чекбоксов
 */
function resetAllCheckboxes() {
  try {
    logInfo('☑️ Сбрасываем все чекбоксы...');
    
    const sheet = getSheet(SHEET_NAMES.IMAGES);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      logInfo('📋 Нет данных для сброса чекбоксов');
      return 0;
    }
    
    // Сбрасываем чекбоксы начиная со второй строки
    const checkboxRange = sheet.getRange(2, IMAGES_COLUMNS.CHECKBOX, lastRow - 1, 1);
    const falseValues = Array(lastRow - 1).fill([false]);
    checkboxRange.setValues(falseValues);
    
    logInfo(`✅ Сброшено ${lastRow - 1} чекбоксов`);
    return lastRow - 1;
    
  } catch (error) {
    logError('Ошибка сброса чекбоксов: ' + error.message);
    throw error;
  }
}

/**
 * ПОДСЧЕТ СТАТИСТИКИ ЛИСТА
 * 
 * Возвращает детальную статистику по данным листа
 * 
 * @returns {Object} Объект со статистикой
 */
function getSheetStatistics() {
  try {
    logInfo('📊 Собираем статистику листа...');
    
    const allData = readImagesData();
    
    const stats = {
      totalProducts: allData.length,
      selectedProducts: 0,
      processingStatuses: {},
      insalesStatuses: {},
      withImages: 0,
      withProcessedImages: 0,
      withAltTags: 0,
      withSeoFilenames: 0,
      lastUpdated: formatDate(new Date(), 'full')
    };
    
    // Анализируем каждый товар
    allData.forEach(product => {
      // Подсчет отмеченных товаров
      if (product.checkbox) {
        stats.selectedProducts++;
      }
      
      // Статистика по статусам обработки
      const procStatus = product.processingStatus || 'Не указано';
      stats.processingStatuses[procStatus] = (stats.processingStatuses[procStatus] || 0) + 1;
      
      // Статистика по статусам InSales
      const insalesStatus = product.insalesStatus || 'Не указано';
      stats.insalesStatuses[insalesStatus] = (stats.insalesStatuses[insalesStatus] || 0) + 1;
      
      // Подсчет заполненных полей
      if (product.originalImages) stats.withImages++;
      if (product.processedImages) stats.withProcessedImages++;
      if (product.altTags) stats.withAltTags++;
      if (product.seoFilenames) stats.withSeoFilenames++;
    });
    
    // Вычисляем проценты
    stats.percentages = {
      selected: stats.totalProducts > 0 ? Math.round((stats.selectedProducts / stats.totalProducts) * 100) : 0,
      withImages: stats.totalProducts > 0 ? Math.round((stats.withImages / stats.totalProducts) * 100) : 0,
      withProcessedImages: stats.totalProducts > 0 ? Math.round((stats.withProcessedImages / stats.totalProducts) * 100) : 0,
      withAltTags: stats.totalProducts > 0 ? Math.round((stats.withAltTags / stats.totalProducts) * 100) : 0,
      withSeoFilenames: stats.totalProducts > 0 ? Math.round((stats.withSeoFilenames / stats.totalProducts) * 100) : 0
    };
    
    logInfo(`✅ Статистика собрана: ${stats.totalProducts} товаров, ${stats.selectedProducts} отмечено`);
    return stats;
    
  } catch (error) {
    logError('Ошибка сбора статистики: ' + error.message);
    return {
      error: error.message,
      timestamp: formatDate(new Date(), 'full')
    };
  }
}

// =============================================================================
// ✅ ФУНКЦИИ ВАЛИДАЦИИ
// =============================================================================

/**
 * ВАЛИДАЦИЯ СТРУКТУРЫ ЛИСТА
 * 
 * Проверяет корректность структуры листа и целостность данных
 * 
 * @returns {Object} Результат валидации
 */
function validateSheetStructure() {
  try {
    logInfo('🔍 Проверяем структуру листа...');
    
    const validation = {
      isValid: true,
      errors: [],
      warnings: [],
      checkedItems: []
    };
    
    // Проверяем существование листа
    let sheet;
    try {
      sheet = getSheet(SHEET_NAMES.IMAGES);
      validation.checkedItems.push(`✅ Лист "${SHEET_NAMES.IMAGES}" существует`);
    } catch (sheetError) {
      validation.errors.push(`❌ Лист "${SHEET_NAMES.IMAGES}" не найден`);
      validation.isValid = false;
      return validation;
    }
    
    // Проверяем заголовки
    try {
      const expectedHeaders = [
        '☑️ Обработать', 'Артикул', 'ID InSales', 'Название товара',
        'Исходные изображения', 'Обработанные изображения', 'Alt-теги',
        'SEO имена файлов', 'Статус обработки', 'Статус InSales'
      ];
      
      const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
      
      for (let i = 0; i < expectedHeaders.length; i++) {
        if (actualHeaders[i] === expectedHeaders[i]) {
          validation.checkedItems.push(`✅ Заголовок колонки ${i + 1} корректен`);
        } else {
          validation.errors.push(`❌ Заголовок колонки ${i + 1} некорректен: ожидается "${expectedHeaders[i]}", найдено "${actualHeaders[i]}"`);
          validation.isValid = false;
        }
      }
    } catch (headerError) {
      validation.errors.push(`❌ Ошибка проверки заголовков: ${headerError.message}`);
      validation.isValid = false;
    }
    
    // Проверяем защиту заголовков
    try {
      const headerRange = sheet.getRange(1, 1, 1, Object.keys(IMAGES_COLUMNS).length);
      const protections = headerRange.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      
      if (protections.length > 0) {
        validation.checkedItems.push('✅ Заголовки защищены от редактирования');
      } else {
        validation.warnings.push('⚠️ Заголовки не защищены от редактирования');
      }
    } catch (protectionError) {
      validation.warnings.push('⚠️ Не удалось проверить защиту заголовков');
    }
    
    // Проверяем наличие данных и их корректность
    const dataCount = sheet.getLastRow() - 1; // Исключаем заголовок
    if (dataCount > 0) {
      validation.checkedItems.push(`✅ Найдено ${dataCount} строк с данными`);
      
      // Проверяем несколько строк на корректность
      try {
        const sampleData = readImagesData();
        const validArticles = sampleData.filter(item => item.article !== '').length;
        validation.checkedItems.push(`✅ Товаров с артикулами: ${validArticles}`);
        
        if (validArticles < sampleData.length) {
          validation.warnings.push(`⚠️ Найдено ${sampleData.length - validArticles} строк без артикулов`);
        }
      } catch (dataError) {
        validation.warnings.push(`⚠️ Ошибка анализа данных: ${dataError.message}`);
      }
    } else {
      validation.checkedItems.push('✅ Лист готов к заполнению данными');
    }
    
    // Итоговая информация
    logInfo(`📋 Валидация структуры завершена: ${validation.isValid ? 'УСПЕШНО' : 'ЕСТЬ ОШИБКИ'}`);
    logInfo(`✅ Проверок пройдено: ${validation.checkedItems.length}`);
    logInfo(`❌ Ошибок: ${validation.errors.length}`);
    logInfo(`⚠️ Предупреждений: ${validation.warnings.length}`);
    
    return validation;
    
  } catch (error) {
    logCritical('Критическая ошибка валидации структуры листа: ' + error.message);
    return {
      isValid: false,
      errors: [`Критическая ошибка: ${error.message}`],
      warnings: [],
      checkedItems: []
    };
  }
}

// =============================================================================
// 🧪 ФУНКЦИИ ТЕСТИРОВАНИЯ МОДУЛЯ
// =============================================================================

/**
 * ТЕСТИРОВАНИЕ МОДУЛЯ DATA_MANAGER
 * 
 * Комплексное тестирование всех функций модуля
 */
function testDataManagerModule() {
  console.log('🧪 Начинаем тестирование модуля 02_data_manager.gs...');
  
  try {
    // Тест 1: Создание структуры листа
    console.log('\n🏗️ Тест 1: Создание и настройка структуры листа');
    try {
      const sheet = createImagesSheet();
      console.log(`✅ Лист создан и настроен: ${sheet.getName()}`);
    } catch (createError) {
      console.log(`ℹ️ Лист уже существует или ошибка создания: ${createError.message}`);
    }
    
    // Тест 2: Валидация структуры
    console.log('\n🔍 Тест 2: Валидация структуры листа');
    const validation = validateSheetStructure();
    console.log(`✅ Валидация выполнена. Статус: ${validation.isValid ? 'OK' : 'ОШИБКИ'}`);
    
    if (validation.errors.length > 0) {
      console.log('❌ Найденные ошибки:');
      validation.errors.forEach(error => console.log(`   ${error}`));
    }
    
    // Тест 3: Запись тестовых данных
    console.log('\n📝 Тест 3: Запись тестовых данных');
    const testProduct = {
      article: 'TEST001',
      productName: 'Тестовый товар',
      originalImages: 'https://example.com/test.jpg',
      processingStatus: STATUS_VALUES.PROCESSING.NOT_PROCESSED,
      insalesStatus: STATUS_VALUES.INSALES.NOT_SENT
    };
    
    try {
      const rowNumber = writeProductData(testProduct);
      console.log(`✅ Тестовый товар записан в строку ${rowNumber}`);
    } catch (writeError) {
      console.log(`❌ Ошибка записи тестовых данных: ${writeError.message}`);
    }
    
    // Тест 4: Чтение данных
    console.log('\n📖 Тест 4: Чтение данных');
    try {
      const allData = readImagesData();
      console.log(`✅ Прочитано ${allData.length} записей`);
      
      if (allData.length > 0) {
        console.log(`✅ Первый товар: ${allData[0].article} - ${allData[0].productName}`);
      }
    } catch (readError) {
      console.log(`❌ Ошибка чтения данных: ${readError.message}`);
    }
    
    // Тест 5: Поиск товара
    console.log('\n🔍 Тест 5: Поиск товара по артикулу');
    try {
      const foundProduct = findProductByArticle('TEST001');
      if (foundProduct) {
        console.log(`✅ Тестовый товар найден: ${foundProduct.productName}`);
      } else {
        console.log(`ℹ️ Тестовый товар не найден (возможно, не был создан)`);
      }
    } catch (searchError) {
      console.log(`❌ Ошибка поиска товара: ${searchError.message}`);
    }
    
    // Тест 6: Обновление статуса
    console.log('\n🔄 Тест 6: Обновление статуса товара');
    try {
      updateProductStatus('TEST001', 'processing', STATUS_VALUES.PROCESSING.COMPLETED);
      console.log(`✅ Статус тестового товара обновлен`);
    } catch (updateError) {
      console.log(`❌ Ошибка обновления статуса: ${updateError.message}`);
    }
    
    // Тест 7: Статистика
    console.log('\n📊 Тест 7: Сбор статистики листа');
    try {
      const stats = getSheetStatistics();
      console.log(`✅ Статистика собрана:`);
      console.log(`   Всего товаров: ${stats.totalProducts}`);
      console.log(`   Отмечено: ${stats.selectedProducts}`);
      console.log(`   С изображениями: ${stats.withImages} (${stats.percentages?.withImages || 0}%)`);
    } catch (statsError) {
      console.log(`❌ Ошибка сбора статистики: ${statsError.message}`);
    }
    
    // Тест 8: Сброс чекбоксов
    console.log('\n☑️ Тест 8: Сброс чекбоксов');
    try {
      const resetCount = resetAllCheckboxes();
      console.log(`✅ Сброшено ${resetCount} чекбоксов`);
    } catch (resetError) {
      console.log(`❌ Ошибка сброса чекбоксов: ${resetError.message}`);
    }
    
    console.log('\n🎉 Тестирование модуля 02_data_manager.gs завершено!');
    console.log('📋 Модуль готов к использованию в проекте');
    
    return {
      success: true,
      message: 'Все тесты пройдены успешно',
      timestamp: formatDate(new Date(), 'full')
    };
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании модуля data_manager:', error.message);
    logCritical('Критическая ошибка в тестировании модуля data_manager', {
      module: '02_data_manager.gs',
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
function demonstrateDataManager() {
  console.log('🎭 Демонстрация возможностей модуля 02_data_manager.gs');
  console.log('='.repeat(60));
  
  // Демо 1: Создание и настройка листа
  console.log('\n🏗️ СОЗДАНИЕ И НАСТРОЙКА ЛИСТА:');
  console.log('- createImagesSheet() - создает основной лист с форматированием');
  console.log('- setupSheetStructure() - настраивает заголовки, стили, защиту');
  console.log('- validateSheetStructure() - проверяет корректность структуры');
  
  // Демо 2: Работа с данными
  console.log('\n📊 РАБОТА С ДАННЫМИ:');
  console.log('- writeProductData() - запись/обновление товара');
  console.log('- readImagesData() - чтение всех данных');
  console.log('- findProductByArticle() - поиск по артикулу');
  console.log('- readSelectedProducts() - только отмеченные товары');
  
  // Демо 3: Управление статусами
  console.log('\n🔄 УПРАВЛЕНИЕ СТАТУСАМИ:');
  console.log('- updateProductStatus() - обновление статуса обработки/InSales');
  console.log('- resetAllCheckboxes() - сброс всех отметок');
  
  // Демо 4: Аналитика
  console.log('\n📈 АНАЛИТИКА И СТАТИСТИКА:');
  console.log('- getSheetStatistics() - подробная статистика листа');
  console.log('- Подсчет товаров по статусам');
  console.log('- Проценты заполненности полей');
  
  // Демо 5: Примеры использования
  console.log('\n💡 ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:');
  
  const exampleProduct = {
    article: 'DEMO123',
    productName: 'Демонстрационный товар',
    originalImages: 'https://example.com/image1.jpg, https://example.com/image2.jpg',
    processingStatus: STATUS_VALUES.PROCESSING.NOT_PROCESSED,
    insalesStatus: STATUS_VALUES.INSALES.NOT_SENT
  };
  
  console.log('Пример товара для записи:');
  console.log(JSON.stringify(exampleProduct, null, 2));
  
  console.log('\nПример поиска и обновления:');
  console.log('const product = findProductByArticle("DEMO123");');
  console.log('if (product) {');
  console.log('  updateProductStatus("DEMO123", "processing", "Обработано");');
  console.log('}');
  
  console.log('\n' + '='.repeat(60));
  console.log('🎉 Демонстрация завершена!');
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
 * 🏗️ СОЗДАНИЕ И НАСТРОЙКА:
 * - createImagesSheet() - создание основного листа
 * - setupSheetStructure(sheet) - настройка структуры
 * - validateSheetStructure() - проверка корректности
 * 
 * 📖 ЧТЕНИЕ ДАННЫХ:
 * - readImagesData() - все данные листа
 * - readSelectedProducts() - только отмеченные товары
 * - findProductByArticle(article) - поиск по артикулу
 * - getSheetStatistics() - статистика листа
 * 
 * ✏️ ЗАПИСЬ ДАННЫХ:
 * - writeProductData(productData, targetRow) - запись товара
 * - writeMultipleProducts(productsArray) - массовая запись
 * - updateProductStatus(article, statusType, statusValue) - обновление статуса
 * 
 * 🧹 УПРАВЛЕНИЕ:
 * - clearImagesData(confirmClear) - очистка данных
 * - resetAllCheckboxes() - сброс всех отметок
 * 
 * СТРУКТУРА ДАННЫХ ТОВАРА:
 * {
 *   article: string,           // Артикул (обязательно)
 *   productName: string,       // Название товара
 *   insalesId: string,         // ID в InSales
 *   originalImages: string,    // Ссылки на исходные изображения
 *   processedImages: string,   // Ссылки на обработанные изображения
 *   altTags: string,          // Alt-теги для SEO
 *   seoFilenames: string,     // SEO имена файлов
 *   processingStatus: string, // Статус обработки
 *   insalesStatus: string,    // Статус InSales
 *   checkbox: boolean         // Отметка для обработки
 * }
 * 
 * ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:
 * 
 * // Создание листа при первом запуске
 * const sheet = createImagesSheet();
 * 
 * // Запись нового товара
 * const newProduct = {
 *   article: 'ART123',
 *   productName: 'Название товара',
 *   originalImages: 'https://example.com/image.jpg'
 * };
 * writeProductData(newProduct);
 * 
 * // Обновление статуса
 * updateProductStatus('ART123', 'processing', STATUS_VALUES.PROCESSING.COMPLETED);
 * 
 * // Получение отмеченных товаров для обработки
 * const selectedProducts = readSelectedProducts();
 * selectedProducts.forEach(product => {
 *   console.log(`Обрабатываем: ${product.article}`);
 * });
 * 
 * // Статистика
 * const stats = getSheetStatistics();
 * console.log(`Всего товаров: ${stats.totalProducts}`);
 * 
 * ===================================================================
 */

// =============================================================================
// 🆕 ФУНКЦИИ ДЛЯ ИМПОРТА ПОЛНЫХ КАРТОЧЕК ТОВАРОВ
// =============================================================================

/**
 * ЗАПИСЬ ПОЛНОЙ КАРТОЧКИ ТОВАРА (ДЛЯ ИМПОРТА)
 *
 * Записывает все данные товара, включая характеристики, описания, цену
 *
 * @param {Object} productData - Полные данные товара
 * @returns {boolean} true если успешно
 */
function writeFullProductData(productData) {
  try {
    logInfo(`📝 Записываем полную карточку товара: ${productData.article}`);

    const sheet = getImagesSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();

    // Ищем существующий товар
    let targetRow = -1;
    const articleCol = IMAGES_COLUMNS.ARTICLE - 1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][articleCol] === productData.article) {
        targetRow = i + 1;
        break;
      }
    }

    // Если не найден - создаем новую строку
    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
    }

    // Подготовка значений для всех колонок
    const rowData = new Array(Object.keys(IMAGES_COLUMNS).length).fill('');

    rowData[IMAGES_COLUMNS.CHECKBOX - 1] = false;
    rowData[IMAGES_COLUMNS.ARTICLE - 1] = productData.article || '';
    rowData[IMAGES_COLUMNS.INSALES_ID - 1] = productData.insalesId || '';
    rowData[IMAGES_COLUMNS.PRODUCT_NAME - 1] = productData.productName || '';
    rowData[IMAGES_COLUMNS.ORIGINAL_IMAGES - 1] = productData.originalImages || '';
    rowData[IMAGES_COLUMNS.SUPPLIER_IMAGES - 1] = productData.supplierImages || '';
    rowData[IMAGES_COLUMNS.ADDITIONAL_IMAGES - 1] = productData.additionalImages || '';
    rowData[IMAGES_COLUMNS.PROCESSED_IMAGES - 1] = productData.processedImages || '';
    rowData[IMAGES_COLUMNS.ALT_TAGS - 1] = productData.altTags || '';
    rowData[IMAGES_COLUMNS.SEO_FILENAMES - 1] = productData.seoFilenames || '';
    rowData[IMAGES_COLUMNS.PROCESSING_STATUS - 1] = productData.processingStatus || STATUS_VALUES.PROCESSING.NOT_PROCESSED;
    rowData[IMAGES_COLUMNS.INSALES_STATUS - 1] = productData.insalesStatus || STATUS_VALUES.INSALES.NOT_SENT;

    // Новые поля для импорта
    rowData[IMAGES_COLUMNS.DESCRIPTION - 1] = productData.description || '';
    rowData[IMAGES_COLUMNS.DESCRIPTION_REWRITTEN - 1] = productData.descriptionRewritten || '';
    rowData[IMAGES_COLUMNS.SHORT_DESCRIPTION - 1] = productData.shortDescription || '';
    rowData[IMAGES_COLUMNS.SPECIFICATIONS_RAW - 1] = productData.specificationsRaw || '';
    rowData[IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED - 1] = productData.specificationsNormalized || '';
    rowData[IMAGES_COLUMNS.PRICE - 1] = productData.price || '';
    rowData[IMAGES_COLUMNS.STOCK - 1] = productData.stock || '';
    rowData[IMAGES_COLUMNS.CATEGORIES - 1] = productData.categories || '';
    rowData[IMAGES_COLUMNS.BRAND - 1] = productData.brand || '';
    rowData[IMAGES_COLUMNS.SERIES - 1] = productData.series || '';
    rowData[IMAGES_COLUMNS.WEIGHT - 1] = productData.weight || '';
    rowData[IMAGES_COLUMNS.DIMENSIONS - 1] = productData.dimensions || '';
    rowData[IMAGES_COLUMNS.PACKAGE_CONTENTS - 1] = productData.packageContents || '';
    rowData[IMAGES_COLUMNS.MATCH_STATUS - 1] = productData.matchStatus || '';
    rowData[IMAGES_COLUMNS.MATCH_CONFIDENCE - 1] = productData.matchConfidence || '';
    rowData[IMAGES_COLUMNS.IMPORT_STATUS - 1] = productData.importStatus || 'Спарсено';

    // Записываем строку
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    logInfo(`✅ Товар ${productData.article} записан в строку ${targetRow}`);
    return true;

  } catch (error) {
    handleError(error, 'Запись полной карточки товара', {
      article: productData.article
    });
    return false;
  }
}

/**
 * ОБНОВЛЕНИЕ ОДНОГО ПОЛЯ ТОВАРА
 *
 * @param {string} article - Артикул товара
 * @param {string} fieldName - Название поля из IMAGES_COLUMNS
 * @param {any} value - Новое значение
 */
function updateProductField(article, fieldName, value) {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    const columnIndex = IMAGES_COLUMNS[fieldName.toUpperCase()];

    if (!columnIndex) {
      throw new Error(`Неизвестное поле: ${fieldName}`);
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][IMAGES_COLUMNS.ARTICLE - 1] === article) {
        sheet.getRange(i + 1, columnIndex).setValue(value);
        logInfo(`✅ Обновлено поле ${fieldName} для ${article}`);
        return true;
      }
    }

    logWarning(`⚠️ Товар ${article} не найден`);
    return false;

  } catch (error) {
    handleError(error, 'Обновление поля товара');
    return false;
  }
}
