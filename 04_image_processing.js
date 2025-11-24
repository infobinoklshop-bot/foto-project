// ========================================
// 04_image_processing.gs
// AI-обработка изображений через GPT-4 Vision
// ========================================

// ========================================
// КОНСТАНТЫ И НАСТРОЙКИ
// ========================================
const IMAGE_PROCESSING_CONFIG = {
  MAX_IMAGES_PER_BATCH: 5,
  AI_TIMEOUT_MS: 30000,
  RETRY_ATTEMPTS: 3,
  ALT_TAG_MAX_LENGTH: 125,
  FILENAME_MAX_LENGTH: 80,
  RETRY_DELAYS: [1000, 3000, 10000], // ms
  FALLBACK_ALT_TAG: "Изображение товара {productName}",
  FALLBACK_FILENAME: "product-image-{timestamp}",
  BATCH_SIZE: 5,
  PROGRESS_SAVE_INTERVAL: 10
};

// ========================================
// ОСНОВНЫЕ ФУНКЦИИ AI-ОБРАБОТКИ
// ========================================

/**
 * Показ диалога выбора изображений перед обработкой
 */
function showImageSelectionForProcessing() {
  try {
    logInfo('📋 Показываем диалог выбора изображений');
    
    // Получаем выбранные товары
    const products = getProductsForProcessing();
    
    if (products.length === 0) {
      showNotification('Отметьте чекбоксами товары в колонке A для обработки', 'warning');
      return;
    }
    
    // Пока работаем с первым товаром (можно расширить для множественного выбора)
    const product = products[0];
    
    // Парсим изображения из всех источников
    const originalImages = product.originalImages ? 
      product.originalImages.split(/[\n,]/).map(url => url.trim()).filter(url => url && url.startsWith('http')) : [];
    
    const supplierImages = product.supplierImages ? 
      product.supplierImages.split(/[\n,]/).map(url => url.trim()).filter(url => url && url.startsWith('http')) : [];
    
    const additionalImages = product.additionalImages ? 
      product.additionalImages.split(/[\n,]/).map(url => url.trim()).filter(url => url && url.startsWith('http')) : [];
    
    // Создаем HTML-диалог
    const html = HtmlService.createTemplateFromFile('ImageSelectionDialog');
    html.productName = product.productName;
    html.article = product.article;
    html.originalImages = originalImages;
    html.supplierImages = supplierImages;
    html.additionalImages = additionalImages;
    
    const htmlOutput = html.evaluate()
      .setWidth(1400)
      .setHeight(800);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Выбор изображений для обработки');
    
  } catch (error) {
    logError('Ошибка показа диалога выбора изображений', error);
    showNotification('Ошибка: ' + error.message, 'error');
  }
}

/**
* Главная функция workflow - обработка выбранных изображений (ИСПРАВЛЕННАЯ)
* Упрощенная версия для надежной работы
*/
async function processSelectedImages() {
  const startTime = new Date();
  const stats = {
    processed: 0,
    errors: 0,
    totalImages: 0
  };

  try {
    logInfo('Начинаем обработку выбранных изображений');

    // 1. Получение товаров для обработки
    const products = getProductsForProcessing();
    
    if (products.length === 0) {
      showNotification('Отметьте чекбоксами товары в колонке A для обработки', 'warning');
      return;
    }

    // 2. Фильтрация товаров с изображениями
    const productsWithImages = products.filter(p => 
      (p.originalImages && p.originalImages.trim()) ||
      (p.supplierImages && p.supplierImages.trim()) ||
      (p.additionalImages && p.additionalImages.trim())
    );
    
    if (productsWithImages.length === 0) {
      showNotification('У выбранных товаров нет изображений', 'warning');
      return;
    }

    logInfo(`Обрабатываем ${productsWithImages.length} товаров`);

    // 3. Обработка каждого товара
    for (let i = 0; i < productsWithImages.length; i++) {
      const product = productsWithImages[i];
      
      try {
        logInfo(`Обрабатываем товар ${i + 1}/${productsWithImages.length}: ${product.productName}`);
        
        // Устанавливаем статус "Обработка..."
        setProcessingStatusSimple(product.article, 'Обработка...');

        // Парсим URL изображений из ВСЕХ источников
        const parseUrls = (text) => {
          if (!text) return [];
          return text.split(/[,\n]/)
            .map(url => url.trim())
            .filter(url => url && url.startsWith('http'));
        };

        const allImageUrls = [
          ...parseUrls(product.originalImages),
          ...parseUrls(product.supplierImages),
          ...parseUrls(product.additionalImages)
        ].slice(0, 10); // Максимум 10 изображений

        if (allImageUrls.length === 0) {
          throw new Error('Нет валидных URL изображений');
        }

        logInfo(`Найдено ${allImageUrls.length} изображений из всех источников`);

        // Обрабатываем все изображения товара за один вызов
        logInfo(`Обрабатываем все изображения товара`);

        const analysisResults = await analyzeImageSimple(product);

        if (analysisResults && analysisResults.length > 0) {
          // Сохраняем результаты сразу
          const altTags = analysisResults.map(r => r.altTag).join('\n');
          const seoFilenames = analysisResults.map(r => r.seoFilename).join('\n');
          const processedUrls = analysisResults.map(r => r.processedImageUrl).filter(Boolean).join('\n');
          
          updateResultsSimple(product.article, altTags, seoFilenames, processedUrls);
          setProcessingStatusSimple(product.article, 'Обработано');
          
          stats.processed++;
          logInfo(`Товар обработан: ${analysisResults.length} изображений`);
        } else {
          throw new Error('Не удалось обработать изображения');
        }

      } catch (productError) {
        logError(`Ошибка обработки товара: ${productError.message}`);
        setProcessingStatusSimple(product.article, 'Ошибка');
        stats.errors++;
      }
    }

    // Показываем результат
    const processingTime = Math.floor((new Date() - startTime) / 1000);
    const message = `Обработано: ${stats.processed}, Ошибок: ${stats.errors}, Время: ${processingTime}сек`;
    
    showNotification(message, stats.errors > 0 ? 'warning' : 'success');
    
    return stats;
    
  } catch (error) {
    logError('Критическая ошибка обработки', error);
    showNotification('Критическая ошибка: ' + error.message, 'error');
    throw error;
  }
}

/**
 * Анализ одного изображения через GPT-4 Vision
 */
async function analyzeProductImage(imageUrl, productName, categoryInfo = '') {
  try {
    const systemPrompt = `Ты эксперт по анализу изображений товаров для интернет-магазина и SEO-оптимизации.

    ТВОЯ РОЛЬ:
    - Анализируешь изображения товаров для e-commerce
    - Создаешь SEO-оптимизированные alt-теги для accessibility и поисковой оптимизации
    - Генерируешь правильные имена файлов для веб-использования

    КРИТИЧЕСКИ ВАЖНО:
    - Название товара "${productName}" — это ФАКТ, не подлежащий изменению
    - НИКОГДА не меняй тип товара (бинокль на монокуляр, телескоп на подзорную трубу и т.д.)
    - В alt-теге и имени файла ВСЕГДА используй название товара как основу
    - Описывай только визуальные детали (ракурс, комплектацию, особенности), сохраняя название`;

    const userPrompt = `Товар: "${productName}" (это точное название, НЕ меняй его!)
    ${categoryInfo ? `Категория: ${categoryInfo}` : ''}

    Проанализируй изображение и создай:
    1. SEO-оптимизированный alt-тег (максимум 125 символов) — ОБЯЗАТЕЛЬНО начни с названия товара "${productName}", добавь описание ракурса/деталей
    2. SEO-имя файла (только латиница, максимум 60 символов) — транслитерируй название товара + ракурс/особенность

    ВАЖНО: На изображении именно "${productName}", даже если визуально похоже на другой прибор.`;
    
    const response = await callOpenAIVision(imageUrl, systemPrompt, userPrompt);
    return parseOpenAIResponse(response, 'image_analysis');
    
  } catch (error) {
    logError(`Ошибка анализа изображения ${imageUrl}`, error);
    throw error;
  }
}

/**
 * Генерация alt-тега на основе анализа изображения
 */
async function generateAltTag(imageAnalysis, productName) {
  try {
    // Если анализ уже содержит alt-тег, используем его
    if (imageAnalysis.alt_tag) {
      return validateAltTag(imageAnalysis.alt_tag, productName);
    }
    
    // Иначе генерируем на основе описания
    const description = imageAnalysis.analysis || imageAnalysis.description || '';
    const keyFeatures = imageAnalysis.key_features || [];
    const colors = imageAnalysis.colors || [];
    
    // Формируем alt-тег
    let altTag = productName;
    
    if (colors.length > 0) {
      altTag = `${colors[0]} ${altTag}`;
    }
    
    if (keyFeatures.length > 0) {
      altTag += ` ${keyFeatures.slice(0, 2).join(' ')}`;
    }
    
    return validateAltTag(altTag, productName);
    
  } catch (error) {
    logError('Ошибка генерации alt-тега', error);
    return IMAGE_PROCESSING_CONFIG.FALLBACK_ALT_TAG.replace('{productName}', productName);
  }
}

/**
 * Создание SEO-имени файла
 */
async function generateSeoFilename(imageAnalysis, productName, originalUrl) {
  try {
    // Если анализ уже содержит SEO-имя, используем его
    if (imageAnalysis.seo_filename) {
      return validateSeoFilename(imageAnalysis.seo_filename);
    }
    
    // Извлекаем расширение из оригинального URL
    const extension = originalUrl.split('.').pop().toLowerCase();
    
    // Генерируем имя на основе данных
    let filename = productName;
    
    if (imageAnalysis.colors && imageAnalysis.colors.length > 0) {
      filename = `${imageAnalysis.colors[0]} ${filename}`;
    }
    
    // Транслитерация и форматирование
    filename = transliterate(filename);
    filename = filename.toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .substring(0, IMAGE_PROCESSING_CONFIG.FILENAME_MAX_LENGTH);
    
    return `${filename}.${extension}`;
    
  } catch (error) {
    logError('Ошибка генерации SEO-имени файла', error);
    return IMAGE_PROCESSING_CONFIG.FALLBACK_FILENAME.replace('{timestamp}', Date.now()) + '.jpg';
  }
}

// ========================================
// ИНТЕГРАЦИЯ С OPENAI API
// ========================================

/**
 * Запрос к GPT-4 Vision API
 */
async function callOpenAIVision(imageUrl, systemPrompt, userPrompt) {
  const settings = getApiSettings();
  
  if (!settings.openaiApiKey) {
    throw new Error('OpenAI API ключ не настроен');
  }
  
  const requestBody = {
    model: "gpt-4-vision-preview",
    messages: [
      {
        role: "system",
        content: systemPrompt
      },
      {
        role: "user",
        content: [
          {
            type: "text",
            text: userPrompt
          },
          {
            type: "image_url",
            image_url: {
              url: imageUrl
            }
          }
        ]
      }
    ],
    max_tokens: 300,
    temperature: 0.3
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${settings.openaiApiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`OpenAI API error: ${responseCode} - ${response.getContentText()}`);
    }
    
    return JSON.parse(response.getContentText());
    
  } catch (error) {
    logError('Ошибка вызова OpenAI Vision API', error);
    throw error;
  }
}

/**
 * Обработка ответов OpenAI
 */
function parseOpenAIResponse(response, requestType) {
  try {
    if (!response.choices || response.choices.length === 0) {
      throw new Error('Пустой ответ от OpenAI');
    }
    
    const content = response.choices[0].message.content;
    
    // Пытаемся распарсить JSON ответ
    try {
      return JSON.parse(content);
    } catch (e) {
      // Если не JSON, парсим текстовый ответ
      return parseTextResponse(content, requestType);
    }
    
  } catch (error) {
    logError('Ошибка парсинга ответа OpenAI', error);
    throw error;
  }
}

// ========================================
// ИНТЕГРАЦИЯ С OPENAI ASSISTANTS API
// ========================================

/**
 * Создание нового thread для Assistant
 */
async function createAssistantThread() {
  const settings = getApiSettings();
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${settings.openaiApiKey}`,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v2'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/threads', options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Ошибка создания thread: ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * Отправка сообщения с изображением в thread
 */
async function sendImageToAssistant(threadId, imageUrl, productName) {
  const settings = getApiSettings();
  
  const messageBody = {
    role: "user",
    content: [
      {
        type: "text",
        text: `Товар: "${productName}" (ТОЧНОЕ название, НЕ меняй!). Создай alt-тег и SEO-имя файла, ОБЯЗАТЕЛЬНО сохраняя название "${productName}" как основу. Описывай только ракурс/детали изображения.`
      },
      {
        type: "image_url",
        image_url: {
          url: imageUrl
        }
      }
    ]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${settings.openaiApiKey}`,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v2'
    },
    payload: JSON.stringify(messageBody),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(
    `https://api.openai.com/v1/threads/${threadId}/messages`,
    options
  );
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Ошибка отправки сообщения: ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * Запуск анализа Assistant и получение результата
 */
async function runAssistantAnalysis(threadId, assistantId) {
  const settings = getApiSettings();
  
  // Запускаем run
  const runBody = {
    assistant_id: assistantId
  };
  
  const runOptions = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${settings.openaiApiKey}`,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v2'
    },
    payload: JSON.stringify(runBody),
    muteHttpExceptions: true
  };
  
  const runResponse = UrlFetchApp.fetch(
    `https://api.openai.com/v1/threads/${threadId}/runs`,
    runOptions
  );
  
  if (runResponse.getResponseCode() !== 200) {
    throw new Error(`Ошибка запуска run: ${runResponse.getContentText()}`);
  }
  
  const run = JSON.parse(runResponse.getContentText());
  
  // Ждем завершения run
  let status = run.status;
  let attempts = 0;
  const maxAttempts = 30; // 30 секунд максимум
  
  while (status === 'queued' || status === 'in_progress') {
    Utilities.sleep(1000); // Ждем 1 секунду
    
    const statusOptions = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${settings.openaiApiKey}`,
        'OpenAI-Beta': 'assistants=v2'
      },
      muteHttpExceptions: true
    };
    
    const statusResponse = UrlFetchApp.fetch(
      `https://api.openai.com/v1/threads/${threadId}/runs/${run.id}`,
      statusOptions
    );
    
    const statusData = JSON.parse(statusResponse.getContentText());
    status = statusData.status;
    
    attempts++;
    if (attempts >= maxAttempts) {
      throw new Error('Таймаут ожидания ответа от Assistant');
    }
  }
  
  if (status !== 'completed') {
    throw new Error(`Assistant run завершился с ошибкой: ${status}`);
  }
  
  // Получаем сообщения из thread
  const messagesOptions = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${settings.openaiApiKey}`,
      'OpenAI-Beta': 'assistants=v2'
    },
    muteHttpExceptions: true
  };
  
  const messagesResponse = UrlFetchApp.fetch(
    `https://api.openai.com/v1/threads/${threadId}/messages`,
    messagesOptions
  );
  
  const messages = JSON.parse(messagesResponse.getContentText());
  
  // Находим последнее сообщение от assistant
  const assistantMessage = messages.data.find(msg => msg.role === 'assistant');
  
  if (!assistantMessage) {
    throw new Error('Не найден ответ от Assistant');
  }
  
  return assistantMessage.content[0].text.value;
}

/**
 * Парсинг ответа Assistant
 */
function parseAssistantResponse(response) {
  try {
    // Если ответ уже JSON
    if (typeof response === 'object') {
      return response;
    }
    
    // Пытаемся найти JSON в тексте
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    
    // Если не нашли JSON, парсим текст
    return parseTextResponse(response, 'assistant');
    
  } catch (error) {
    logError('Ошибка парсинга ответа Assistant', error);
    throw error;
  }
}

/**
 * Валидация выходных данных Assistant
 */
function validateAssistantOutput(output, productName) {
  const validated = {
    analysis: output.analysis || '',
    alt_tag: validateAltTag(output.alt_tag || '', productName),
    seo_filename: validateSeoFilename(output.seo_filename || ''),
    confidence: output.confidence || 5,
    match_product_name: output.match_product_name !== false,
    key_features: output.key_features || [],
    colors: output.colors || []
  };
  
  // Если alt-тег пустой, генерируем fallback
  if (!validated.alt_tag) {
    validated.alt_tag = IMAGE_PROCESSING_CONFIG.FALLBACK_ALT_TAG.replace('{productName}', productName);
  }
  
  // Если имя файла пустое, генерируем fallback
  if (!validated.seo_filename) {
    validated.seo_filename = IMAGE_PROCESSING_CONFIG.FALLBACK_FILENAME.replace('{timestamp}', Date.now());
  }
  
  return validated;
}

// ========================================
// РАБОТА С GOOGLE SHEETS
// ========================================

/**
 * Получение товаров для обработки
 */
function getProductsForProcessing() {
  const sheet = getImagesSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const products = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const isSelected = row[IMAGES_COLUMNS.CHECKBOX - 1];
    const hasImages = row[IMAGES_COLUMNS.ORIGINAL_IMAGES - 1] || 
                      row[IMAGES_COLUMNS.SUPPLIER_IMAGES - 1] || 
                      row[IMAGES_COLUMNS.ADDITIONAL_IMAGES - 1];
    const processingStatus = row[IMAGES_COLUMNS.PROCESSING_STATUS - 1];

    if (isSelected && hasImages) {
      products.push({
        article: row[IMAGES_COLUMNS.ARTICLE - 1],
        productName: row[IMAGES_COLUMNS.PRODUCT_NAME - 1],
        originalImages: row[IMAGES_COLUMNS.ORIGINAL_IMAGES - 1],
        supplierImages: row[IMAGES_COLUMNS.SUPPLIER_IMAGES - 1],
        additionalImages: row[IMAGES_COLUMNS.ADDITIONAL_IMAGES - 1], // Добавить эту строку
        rowIndex: i + 1
      });
    }
  }

  return products;
}

/**
 * Обновление результатов в листе
 */
function updateProcessingResults(article, altTags, seoFilenames) {
  try {
    const sheet = getImagesSheet();
    const articleRange = sheet.getRange(2, IMAGES_COLUMNS.ARTICLE, sheet.getLastRow() - 1, 1);
    const articles = articleRange.getValues();
    
    for (let i = 0; i < articles.length; i++) {
      if (articles[i][0] === article) {
        const rowIndex = i + 2;
        
        // Обновляем alt-теги
        sheet.getRange(rowIndex, IMAGES_COLUMNS.ALT_TAGS).setValue(altTags);
        
        // Обновляем SEO-имена файлов
        sheet.getRange(rowIndex, IMAGES_COLUMNS.SEO_FILENAMES).setValue(seoFilenames);
        
        // Обновляем временную метку
        sheet.getRange(rowIndex, IMAGES_COLUMNS.LAST_UPDATED).setValue(new Date());
        
        break;
      }
    }
    
  } catch (error) {
    logError(`Ошибка обновления результатов для артикула ${article}`, error);
    throw error;
  }
}

/**
 * Обновление статусов обработки
 */
function setProcessingStatus(article, status, comment = '') {
  try {
    const sheet = getImagesSheet();
    const articleRange = sheet.getRange(2, IMAGES_COLUMNS.ARTICLE, sheet.getLastRow() - 1, 1);
    const articles = articleRange.getValues();
    
    for (let i = 0; i < articles.length; i++) {
      if (articles[i][0] === article) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, IMAGES_COLUMNS.PROCESSING_STATUS).setValue(status);
        
        if (comment) {
          const existingNotes = sheet.getRange(rowIndex, IMAGES_COLUMNS.NOTES).getValue() || '';
          const newNotes = `[${new Date().toLocaleString()}] ${comment}\n${existingNotes}`;
          sheet.getRange(rowIndex, IMAGES_COLUMNS.NOTES).setValue(newNotes.substring(0, 1000));
        }
        
        break;
      }
    }
    
  } catch (error) {
    logError(`Ошибка установки статуса для артикула ${article}`, error);
  }
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ========================================

/**
 * Валидация alt-тега
 */
function validateAltTag(altTag, productName) {
  if (!altTag) {
    return IMAGE_PROCESSING_CONFIG.FALLBACK_ALT_TAG.replace('{productName}', productName);
  }

  // Убираем лишние пробелы
  altTag = altTag.trim().replace(/\s+/g, ' ');

  // Заменяем * на x в обозначениях кратности (7*40 → 7x40, 8*42 → 8x42)
  altTag = altTag.replace(/(\d+)\s*\*\s*(\d+)/g, '$1x$2');

  // Ограничиваем длину
  if (altTag.length > IMAGE_PROCESSING_CONFIG.ALT_TAG_MAX_LENGTH) {
    altTag = altTag.substring(0, IMAGE_PROCESSING_CONFIG.ALT_TAG_MAX_LENGTH - 3) + '...';
  }

  // Убираем слова "изображение", "фото", "картинка"
  altTag = altTag.replace(/\b(изображение|фото|картинка)\b/gi, '').trim();

  return altTag;
}

/**
 * Валидация SEO-имени файла
 */
function validateSeoFilename(filename) {
 if (!filename) {
   return IMAGE_PROCESSING_CONFIG.FALLBACK_FILENAME.replace('{timestamp}', Date.now());
 }

 // Убираем точки в конце (AI иногда добавляет лишние)
 let name = filename.replace(/\.+$/, '');

 // Убираем расширение если есть (.webp, .jpg, .png и т.д.)
 name = name.replace(/\.(webp|jpg|jpeg|png|gif|bmp|svg)$/i, '');

 // Оставляем только латиницу, цифры и дефисы
 name = name.toLowerCase()
   .replace(/[^a-z0-9-]/g, '-')
   .replace(/-+/g, '-')
   .replace(/^-+|-+$/g, '');

 // Ограничиваем длину
 if (name.length > IMAGE_PROCESSING_CONFIG.FILENAME_MAX_LENGTH) {
   name = name.substring(0, IMAGE_PROCESSING_CONFIG.FILENAME_MAX_LENGTH);
 }

 // Если имя пустое после валидации
 if (!name) {
   name = IMAGE_PROCESSING_CONFIG.FALLBACK_FILENAME.replace('{timestamp}', Date.now());
 }

 return name; // Возвращаем ТОЛЬКО имя без расширения
}

/**
 * Транслитерация русского текста
 */
function transliterate(text) {
  const rules = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd',
    'е': 'e', 'ё': 'yo', 'ж': 'zh', 'з': 'z', 'и': 'i',
    'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n',
    'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't',
    'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch',
    'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '',
    'э': 'e', 'ю': 'yu', 'я': 'ya',
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D',
    'Е': 'E', 'Ё': 'Yo', 'Ж': 'Zh', 'З': 'Z', 'И': 'I',
    'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N',
    'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T',
    'У': 'U', 'Ф': 'F', 'Х': 'Kh', 'Ц': 'Ts', 'Ч': 'Ch',
    'Ш': 'Sh', 'Щ': 'Sch', 'Ъ': '', 'Ы': 'Y', 'Ь': '',
    'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya'
  };
  
  return text.split('').map(char => rules[char] || char).join('');
}

/**
 * Парсинг текстового ответа (fallback)
 */
function parseTextResponse(text, requestType) {
  const result = {
    analysis: text,
    alt_tag: '',
    seo_filename: '',
    confidence: 5,
    match_product_name: true,
    key_features: [],
    colors: []
  };
  
  // Пытаемся извлечь alt-тег
  const altMatch = text.match(/alt[^:]*:\s*"?([^"\n]+)"?/i);
  if (altMatch) {
    result.alt_tag = altMatch[1].trim();
  }
  
  // Пытаемся извлечь имя файла
  const filenameMatch = text.match(/(?:filename|имя файла)[^:]*:\s*"?([^"\n]+)"?/i);
  if (filenameMatch) {
    result.seo_filename = filenameMatch[1].trim();
  }
  
  // Пытаемся извлечь цвета
  const colorMatch = text.match(/(?:цвет|color)[^:]*:\s*([^\n]+)/i);
  if (colorMatch) {
    result.colors = colorMatch[1].split(/[,;]/).map(c => c.trim());
  }
  
  return result;
}

/**
 * Обработка ошибок API с retry-логикой
 */
async function handleAssistantApiError(error, article = '') {
  const errorMessage = error.toString();
  
  // Определяем тип ошибки
  if (errorMessage.includes('401')) {
    logError('Неверный API ключ OpenAI', { article });
    throw new Error('Неверный API ключ. Проверьте настройки.');
  } else if (errorMessage.includes('429')) {
    logError('Превышен лимит запросов OpenAI', { article });
    throw new Error('Превышен лимит запросов. Попробуйте позже.');
  } else if (errorMessage.includes('500')) {
    logError('Внутренняя ошибка сервера OpenAI', { article });
    throw new Error('Ошибка сервера OpenAI. Попробуйте позже.');
  } else {
    logError('Неизвестная ошибка API', { error: errorMessage, article });
    throw error;
  }
}

/**
 * Получение статистики обработки
 */
function getProcessingStatistics() {
  const sheet = getImagesSheet();
  const data = sheet.getDataRange().getValues();
  
  const stats = {
    total: 0,
    processed: 0,
    errors: 0,
    notProcessed: 0,
    withAltTags: 0,
    withSeoFilenames: 0
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[IMAGES_COLUMNS.PROCESSING_STATUS - 1];
    const altTags = row[IMAGES_COLUMNS.ALT_TAGS - 1];
    const seoFilenames = row[IMAGES_COLUMNS.SEO_FILENAMES - 1];
    
    stats.total++;
    
    if (status === STATUS_VALUES.PROCESSING.COMPLETED) {
      stats.processed++;
    } else if (status === STATUS_VALUES.PROCESSING.ERROR) {
      stats.errors++;
    } else if (status === STATUS_VALUES.PROCESSING.NOT_PROCESSED) {
      stats.notProcessed++;
    }
    
    if (altTags) stats.withAltTags++;
    if (seoFilenames) stats.withSeoFilenames++;
  }
  
  return stats;
}

// ========================================
// UI ФУНКЦИИ
// ========================================

/**
 * Показать диалог прогресса обработки
 */
function showProcessingProgress() {
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>Обработка изображений</h3>
      <div id="progress-info">
        <p>Инициализация...</p>
      </div>
      <div style="margin-top: 20px;">
        <div style="background: #f0f0f0; height: 20px; border-radius: 10px;">
          <div id="progress-bar" style="background: #4CAF50; height: 100%; width: 0%; border-radius: 10px; transition: width 0.3s;"></div>
        </div>
      </div>
      <div id="stats" style="margin-top: 20px; font-size: 14px; color: #666;">
      </div>
    </div>
    <script>
      // Обновление прогресса каждые 5 секунд
      setInterval(() => {
        google.script.run
          .withSuccessHandler(updateProgress)
          .getProcessingStatistics();
      }, 5000);
      
      function updateProgress(stats) {
        const processed = stats.processed + stats.errors;
        const total = stats.total;
        const percentage = total > 0 ? Math.round((processed / total) * 100) : 0;
        
        document.getElementById('progress-bar').style.width = percentage + '%';
        document.getElementById('progress-info').innerHTML = 
          '<p>Обработано: ' + processed + ' из ' + total + ' (' + percentage + '%)</p>';
        
        document.getElementById('stats').innerHTML = 
          '<p>✅ Успешно: ' + stats.processed + '<br>' +
          '❌ Ошибок: ' + stats.errors + '<br>' +
          '⏳ Ожидает: ' + stats.notProcessed + '</p>';
      }
    </script>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(300);
    
  SpreadsheetApp.getUi().showModelessDialog(html, 'Прогресс обработки');
}

// ========================================
// ТЕСТОВЫЕ ФУНКЦИИ
// ========================================

/**
 * Тест анализа одного изображения
 */
async function testImageAnalysis() {
  const testImageUrl = 'https://example.com/test-shoe.jpg';
  const testProductName = 'Кроссовки Nike Air Max черные размер 42';
  
  try {
    logInfo('Начало тестирования анализа изображения');
    
    const result = await analyzeProductImage(testImageUrl, testProductName);
    
    logInfo('Результат анализа:', result);
    
    // Проверка результата
    if (!result.alt_tag || !result.seo_filename) {
      throw new Error('Неполный результат анализа');
    }
    
    logInfo('Тест успешно завершен');
    return result;
    
  } catch (error) {
    logError('Ошибка тестирования', error);
    throw error;
  }
}

/**
 * Тест генерации alt-тега
 */
async function testAltTagGeneration() {
  const testAnalysis = {
    analysis: 'Черные кроссовки Nike Air Max на белой подошве',
    key_features: ['спортивная обувь', 'амортизация'],
    colors: ['черный', 'белый']
  };
  const testProductName = 'Кроссовки Nike Air Max';
  
  const altTag = await generateAltTag(testAnalysis, testProductName);
  
  logInfo('Сгенерированный alt-тег:', altTag);
  
  // Проверки
  if (altTag.length > IMAGE_PROCESSING_CONFIG.ALT_TAG_MAX_LENGTH) {
    throw new Error('Alt-тег превышает максимальную длину');
  }
  
  return altTag;
}

/**
 * Тест транслитерации
 */
function testTransliteration() {
  const testCases = [
    { input: 'Красные женские туфли', expected: 'Krasnye zhenskie tufli' },
    { input: 'Синий мужской пиджак', expected: 'Siniy muzhskoy pidzhak' },
    { input: 'Белые кроссовки Nike', expected: 'Belye krossovki Nike' }
  ];
  
  testCases.forEach(testCase => {
    const result = transliterate(testCase.input);
    logInfo(`Транслитерация: "${testCase.input}" -> "${result}"`);
    
    if (result !== testCase.expected) {
      logError(`Ошибка транслитерации: ожидалось "${testCase.expected}", получено "${result}"`);
    }
  });
}

// ========================================
// ЭКСПОРТИРУЕМЫЕ ФУНКЦИИ ДЛЯ МЕНЮ
// ========================================

/**
 * Запуск обработки выбранных изображений (для меню)
 */
function runImageProcessing() {
  const ui = SpreadsheetApp.getUi();
  
  // Проверка настроек
  const settings = getApiSettings();
  if (!settings.openaiApiKey) {
    ui.alert('Ошибка', 'OpenAI API ключ не настроен. Перейдите в Настройки.', ui.ButtonSet.OK);
    return;
  }
  
  if (!settings.openaiAltAssistantId) {
    ui.alert('Ошибка', 'ID Assistant не настроен. Перейдите в Настройки.', ui.ButtonSet.OK);
    return;
  }
  
  // Подтверждение запуска
  const result = ui.alert(
    'Подтверждение',
    'Начать AI-обработку выбранных изображений? Это может занять несколько минут.',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  // Показываем прогресс
  showProcessingProgress();
  
  // Запускаем обработку
  try {
    processSelectedImages();
  } catch (error) {
    ui.alert('Ошибка', 'Произошла ошибка: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Показать статистику обработки (для меню)
 */
function showProcessingStats() {
  const stats = getProcessingStatistics();
  
  const message = `
📊 Статистика обработки изображений:

Всего товаров: ${stats.total}
✅ Обработано успешно: ${stats.processed}
❌ С ошибками: ${stats.errors}
⏳ Не обработано: ${stats.notProcessed}

📝 С alt-тегами: ${stats.withAltTags}
📁 С SEO-именами: ${stats.withSeoFilenames}

Процент обработки: ${Math.round((stats.processed / stats.total) * 100)}%
  `;
  
  SpreadsheetApp.getUi().alert('Статистика обработки', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Сброс статусов обработки (для меню)
 */
function resetProcessingStatuses() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Подтверждение',
    'Сбросить все статусы обработки? Это позволит заново обработать все товары.',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  const sheet = getImagesSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const statusRange = sheet.getRange(2, IMAGES_COLUMNS.PROCESSING_STATUS, lastRow - 1, 1);
    statusRange.setValue(STATUS_VALUES.PROCESSING.NOT_PROCESSED);
    
    // Очищаем alt-теги и SEO-имена
    sheet.getRange(2, IMAGES_COLUMNS.ALT_TAGS, lastRow - 1, 1).clearContent();
    sheet.getRange(2, IMAGES_COLUMNS.SEO_FILENAMES, lastRow - 1, 1).clearContent();
  }
  
  ui.alert('Готово', 'Статусы обработки сброшены.', ui.ButtonSet.OK);
}

// ========================================
// ИНИЦИАЛИЗАЦИЯ МОДУЛЯ
// ========================================

/**
 * Проверка готовности модуля
 */
function checkImageProcessingModule() {
  const requiredFunctions = [
    'processSelectedImages',
    'analyzeProductImage',
    'generateAltTag',
    'generateSeoFilename',
    'callOpenAIVision',
    'createAssistantThread',
    'runAssistantAnalysis'
  ];
  
  const missingFunctions = [];
  
  requiredFunctions.forEach(funcName => {
    if (typeof this[funcName] !== 'function') {
      missingFunctions.push(funcName);
    }
  });
  
  if (missingFunctions.length > 0) {
    logError('Модуль 04_image_processing.gs не полностью загружен', {
      missing: missingFunctions
    });
    return false;
  }
  
  logInfo('✅ Модуль 04_image_processing.gs успешно загружен');
  return true;
}

/**
 * Упрощенное обновление статуса
 */
function setProcessingStatusSimple(article, status) {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === article) { // Колонка B - артикул
        sheet.getRange(i + 1, IMAGES_COLUMNS.PROCESSING_STATUS).setValue(status); // Колонка K - статус обработки
        return true;
      }
    }
    return false;
  } catch (error) {
    logError(`Ошибка обновления статуса: ${error.message}`);
    return false;
  }
}

/**
* Упрощенное сохранение результатов
*/
function updateResultsSimple(article, altTags, seoFilenames, processedUrls) {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === article) { // Колонка B - артикул
        
        // Форматируем данные для отображения в столбик (каждое значение с новой строки)
        const formattedProcessedUrls = processedUrls ? processedUrls.replace(/\s*\|\s*/g, '\n') : '';
        const formattedAltTags = altTags ? altTags.replace(/\s*\|\s*/g, '\n') : '';
        // Убираем точки в конце SEO-имен (AI иногда добавляет)
        const formattedSeoFilenames = seoFilenames ? seoFilenames
          .replace(/\s*\|\s*/g, '\n')
          .split('\n')
          .map(name => name.trim().replace(/\.+$/, ''))
          .join('\n') : '';
        
        // Сохраняем в ячейки
        sheet.getRange(i + 1, IMAGES_COLUMNS.PROCESSED_IMAGES).setValue(formattedProcessedUrls); // H - колонка 8
        sheet.getRange(i + 1, IMAGES_COLUMNS.ALT_TAGS).setValue(formattedAltTags); // I - колонка 9
        sheet.getRange(i + 1, IMAGES_COLUMNS.SEO_FILENAMES).setValue(formattedSeoFilenames); // J - колонка 10
        
        return true;
      }
    }
    return false;
  } catch (error) {
    logError(`Ошибка сохранения результатов: ${error.message}`);
    return false;
  }
}

/**
 * ОПТИМИЗИРОВАННЫЙ WORKFLOW - Упрощенная и надежная схема
 */

/**
 * Главная функция - упрощенная обработка изображений
 */
async function analyzeImageSimple(product) {
  try {
    const productName = product.productName;
    
    logInfo(`Начинаем оптимизированную обработку товара: ${productName}`);
    
    // Объединяем все изображения из разных источников
    const allImageUrls = await combineAllImages(
      product.originalImages, 
      product.supplierImages,
      product.additionalImages, 
      product.article
    );
    
    if (allImageUrls.length === 0) {
      throw new Error('Нет изображений для обработки');
    }

    // Ограничиваем до 10 изображений
    const imagesToProcess = allImageUrls.slice(0, 10);
    logInfo(`Обрабатываем ${imagesToProcess.length} изображений по оптимизированной схеме`);
    
    // ЭТАП 1: OpenAI анализ исходных изображений
    logInfo(`Этап 1: AI-анализ ${imagesToProcess.length} исходных изображений`);
    const analysisResults = await processOpenAIOriginals(imagesToProcess, productName);
    
    // ЭТАП 2: Replicate улучшение (ОПЦИОНАЛЬНО - можно пропустить)
    let enhancedImages;
    const settings = getApiSettings();

    if (settings.replicateToken) {
      logInfo(`Этап 2: Replicate улучшение с fallback`);
      enhancedImages = await processReplicateWithFallback(imagesToProcess);

      // Проверяем, все ли провалились
      const allFailed = enhancedImages.every(img => !img.wasEnhanced);
      if (allFailed) {
        logWarning('Replicate: все изображения провалились. Используем исходные.');
      }
    } else {
      logInfo(`Этап 2: Replicate ПРОПУЩЕН (токен не настроен)`);
      enhancedImages = imagesToProcess.map(url => ({ original: url, processed: url, wasEnhanced: false }));
    }
    
    // ЭТАП 3: Единая WebP оптимизация всех изображений
    logInfo(`Этап 3: WebP оптимизация всех изображений`);
    const finalResults = await processUnifiedWebPOptimization(enhancedImages, analysisResults);
    
    logInfo(`Оптимизированная обработка завершена: ${finalResults.length} изображений`);
    
    // ДОБАВИТЬ ФИНАЛЬНОЕ УВЕДОМЛЕНИЕ:
    showFinalProcessingResults(enhancedImages, finalResults, productName);
    
    return finalResults;
    
  } catch (error) {
    logError('Критическая ошибка оптимизированной обработки', error);
    
    // Fallback результат
    return [{
      altTag: `${product.productName} - изображение товара`,
      seoFilename: `product-${Date.now()}`,
      processedImageUrl: '',
      confidence: 1
    }];
  }
}

/**
 * ЭТАП 1: OpenAI анализ исходных изображений
 */
async function processOpenAIOriginals(imageUrls, productName) {
  try {
    const settings = getApiSettings();
    
    if (!settings.openaiAltAssistantId) {
      logWarning('OpenAI Assistant недоступен, генерируем базовые теги');
      return imageUrls.map((url, i) => ({
        altTag: `${productName} - изображение ${i + 1}`,
        seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-')
      }));
    }

    const BATCH_SIZE = 5;
    const allResults = [];
    
    // Разбиваем на группы по 5
    for (let i = 0; i < imageUrls.length; i += BATCH_SIZE) {
      const batch = imageUrls.slice(i, i + BATCH_SIZE);
      const groupNumber = Math.floor(i / BATCH_SIZE) + 1;
      
      logInfo(`OpenAI анализ группы ${groupNumber}: ${batch.length} исходных изображений`);
      
      try {
        const thread = await createAssistantThread();
        
        // Формируем сообщение с исходными изображениями
        const messageContent = [
          {
            type: "text",
            text: `ТОВАР: "${productName}" — это ТОЧНОЕ название, НЕ МЕНЯЙ ЕГО!

Проанализируй ${batch.length} изображений и создай для каждого:
1. SEO-оптимизированный alt-тег (до 125 символов) — ОБЯЗАТЕЛЬНО начинай с "${productName}", добавляй только описание ракурса/деталей
2. SEO-имя файла (латиница, до 60 символов, БЕЗ расширения) — транслитерируй "${productName}" + ракурс

КРИТИЧЕСКИ ВАЖНО:
- НИКОГДА не меняй тип товара (бинокль НЕ монокуляр, телескоп НЕ труба)
- Сохраняй "${productName}" как основу каждого alt-тега
- Описывай только визуальные детали: ракурс, комплектацию, цвет

Отвечай в JSON формате:
{
  "results": [
    {"altTag": "${productName} - ракурс/детали", "seoFilename": "transliterated-name-detail"},
    {"altTag": "${productName} - другой ракурс", "seoFilename": "transliterated-name-detail2"}
  ]
}`
          }
        ];

        // Добавляем исходные изображения
        batch.forEach((imageUrl, index) => {
          messageContent.push({
            type: "image_url",
            image_url: { url: imageUrl }
          });
        });

        // Отправляем и анализируем
        const messageResponse = UrlFetchApp.fetch(
          `https://api.openai.com/v1/threads/${thread.id}/messages`,
          {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${settings.openaiApiKey}`,
              'Content-Type': 'application/json',
              'OpenAI-Beta': 'assistants=v2'
            },
            payload: JSON.stringify({
              role: "user",
              content: messageContent
            }),
            muteHttpExceptions: true
          }
        );

        if (messageResponse.getResponseCode() !== 200) {
          throw new Error(`OpenAI message failed: ${messageResponse.getResponseCode()}`);
        }

        const analysisResult = await runAssistantAnalysis(thread.id, settings.openaiAltAssistantId);
        const parsed = parseAssistantBatchResponse(analysisResult, batch.length, productName);
        
        allResults.push(...parsed);
        
        // Пауза между группами
        if (i + BATCH_SIZE < imageUrls.length) {
          Utilities.sleep(2000);
        }
        
      } catch (groupError) {
        logError(`Ошибка OpenAI группы ${groupNumber}`, groupError);
        
        // Fallback для текущей группы
        const fallbackResults = batch.map((url, index) => ({
          altTag: `${productName} - изображение ${i + index + 1}`,
          seoFilename: transliterate(`${productName}-${i + index + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
        }));
        
        allResults.push(...fallbackResults);
      }
    }
    
    logInfo(`OpenAI анализ завершен: ${allResults.length} результатов`);
    return allResults;
    
  } catch (error) {
    logError('Критическая ошибка OpenAI анализа', error);
    
    // Fallback - генерируем базовые теги
    return imageUrls.map((url, i) => ({
      altTag: `${productName} - изображение ${i + 1}`,
      seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
    }));
  }
}

/**
 * Определение размеров изображения в пикселях по заголовкам JPEG/PNG
 * Скачивает только первые 64KB для анализа (быстро и экономно)
 */
async function getImageDimensions(imageUrl) {
  try {
    // Скачиваем изображение (частичная загрузка не поддерживается всеми серверами)
    const response = UrlFetchApp.fetch(imageUrl, {
      muteHttpExceptions: true,
      // Некоторые CDN не поддерживают Range, поэтому скачиваем полностью
      // но анализируем только первые байты
    });

    if (response.getResponseCode() !== 200) {
      return null;
    }

    const bytes = response.getBlob().getBytes();

    // Проверяем минимальный размер
    if (bytes.length < 24) {
      return null;
    }

    // PNG: magic bytes 89 50 4E 47 0D 0A 1A 0A
    if (bytes[0] === -119 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
      // PNG IHDR chunk: width в байтах 16-19, height в байтах 20-23 (big-endian)
      const width = ((bytes[16] & 0xFF) << 24) | ((bytes[17] & 0xFF) << 16) | ((bytes[18] & 0xFF) << 8) | (bytes[19] & 0xFF);
      const height = ((bytes[20] & 0xFF) << 24) | ((bytes[21] & 0xFF) << 16) | ((bytes[22] & 0xFF) << 8) | (bytes[23] & 0xFF);
      return { width, height, format: 'PNG' };
    }

    // JPEG: magic bytes FF D8 FF
    if (bytes[0] === -1 && bytes[1] === -40 && bytes[2] === -1) {
      // Ищем SOF0 (0xFFC0) или SOF2 (0xFFC2) маркер
      let offset = 2;
      while (offset < bytes.length - 9) {
        if (bytes[offset] === -1) { // 0xFF
          const marker = bytes[offset + 1] & 0xFF;

          // SOF0 (0xC0), SOF1 (0xC1), SOF2 (0xC2) - содержат размеры
          if (marker >= 0xC0 && marker <= 0xC3) {
            // Height в байтах offset+5,6, Width в offset+7,8 (big-endian)
            const height = ((bytes[offset + 5] & 0xFF) << 8) | (bytes[offset + 6] & 0xFF);
            const width = ((bytes[offset + 7] & 0xFF) << 8) | (bytes[offset + 8] & 0xFF);
            return { width, height, format: 'JPEG' };
          }

          // Пропускаем другие сегменты
          if (marker !== 0x00 && marker !== 0xFF && marker !== 0xD0 && marker !== 0xD8 && marker !== 0xD9) {
            const segmentLength = ((bytes[offset + 2] & 0xFF) << 8) | (bytes[offset + 3] & 0xFF);
            offset += 2 + segmentLength;
          } else {
            offset += 2;
          }
        } else {
          offset++;
        }

        // Ограничение поиска первыми 64KB
        if (offset > 65536) break;
      }
    }

    // WebP: RIFF....WEBP
    if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46 &&
        bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50) {
      // VP8 chunk для lossy WebP
      if (bytes[12] === 0x56 && bytes[13] === 0x50 && bytes[14] === 0x38 && bytes[15] === 0x20) {
        // VP8: размеры в байтах 26-29
        const width = ((bytes[27] & 0xFF) << 8) | (bytes[26] & 0xFF);
        const height = ((bytes[29] & 0xFF) << 8) | (bytes[28] & 0xFF);
        return { width: width & 0x3FFF, height: height & 0x3FFF, format: 'WebP' };
      }
    }

    return null; // Неизвестный формат

  } catch (error) {
    logWarning(`Ошибка определения размеров: ${error.message}`);
    return null;
  }
}

/**
 * ЭТАП 2: Replicate улучшение с проверкой размера (только для маленьких изображений)
 */
async function processReplicateWithFallback(imageUrls) {
  const settings = getApiSettings();

  if (!settings.replicateToken) {
    logWarning('Replicate Token недоступен, используем исходные изображения');
    return imageUrls.map(url => ({ original: url, processed: url, wasEnhanced: false }));
  }

  // Определяем модель
  const modelKey = settings.replicateModel || 'esrgan';
  const modelConfig = getReplicateModelConfig(modelKey);

  // ЛИМИТ: изображения >= 1500px уже качественные, Replicate не нужен
  const MIN_DIMENSION_FOR_SKIP = 1500;

  logInfo(`Replicate (${modelConfig.name}): анализируем ${imageUrls.length} изображений`);

  // Последовательная обработка с проверкой размера
  const results = [];

  for (let index = 0; index < imageUrls.length; index++) {
    const imageUrl = imageUrls[index];

    try {
      // ПРОВЕРКА РАЗМЕРА ИЗОБРАЖЕНИЯ В ПИКСЕЛЯХ
      const dimensions = await getImageDimensions(imageUrl);

      if (dimensions && dimensions.width && dimensions.height) {
        const maxSide = Math.max(dimensions.width, dimensions.height);

        // Пропускаем большие изображения — они уже качественные
        if (maxSide >= MIN_DIMENSION_FOR_SKIP) {
          logInfo(`Изображение ${index + 1}: ${dimensions.width}×${dimensions.height}px — качественное, Replicate не требуется`);
          results.push({ original: imageUrl, processed: imageUrl, wasEnhanced: false, skippedReason: 'already_high_quality' });
          continue;
        }

        logInfo(`Изображение ${index + 1}: ${dimensions.width}×${dimensions.height}px — маленькое, улучшаем`);
      } else {
        logWarning(`Изображение ${index + 1}: не удалось определить размер, пробуем Replicate`);
      }

      // Отправляем в Replicate только маленькие изображения
      logInfo(`Запускаем Replicate для изображения ${index + 1}`);

      const predictionResponse = UrlFetchApp.fetch('https://api.replicate.com/v1/predictions', {
        method: 'POST',
        headers: {
          'Authorization': `Token ${settings.replicateToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          version: modelConfig.version, // ИСПОЛЬЗОВАТЬ ВЕРСИЮ МОДЕЛИ
          input: {
            image: imageUrl,
            scale: Math.min(settings.replicateScale, modelConfig.maxScale) // ОГРАНИЧИТЬ SCALE
          }
        }),
        muteHttpExceptions: true
      });

      const responseCode = predictionResponse.getResponseCode();
      const responseText = predictionResponse.getContentText();

      if (responseCode !== 201) {
        // Детальная диагностика ошибки
        let errorDetails = `HTTP ${responseCode}`;
        try {
          const errorData = JSON.parse(responseText);
          errorDetails += `: ${errorData.detail || errorData.error || responseText.substring(0, 100)}`;
        } catch (e) {
          errorDetails += `: ${responseText.substring(0, 100)}`;
        }
        throw new Error(`Replicate prediction failed: ${errorDetails}`);
      }

      const prediction = JSON.parse(responseText);

      // Ждем результат (сокращен таймаут до 30 сек для ускорения)
      const result = await waitForReplicateResult(prediction.id, settings.replicateToken, 30);

      if (result.success) {
        logInfo(`Изображение ${index + 1} улучшено через Replicate (${modelConfig.name})`);
        results.push({ original: imageUrl, processed: result.outputUrl, wasEnhanced: true });
      } else {
        throw new Error(`Replicate processing failed: ${result.reason || 'unknown'}`);
      }

    } catch (error) {
      logWarning(`Replicate не удался для изображения ${index + 1}, используем исходное: ${error.message}`);
      results.push({ original: imageUrl, processed: imageUrl, wasEnhanced: false, error: error.message });
    }
  }

  // Статистика
  const enhancedCount = results.filter(r => r.wasEnhanced).length;
  const skippedHighQuality = results.filter(r => r.skippedReason === 'already_high_quality').length;
  const failedCount = results.filter(r => !r.wasEnhanced && !r.skippedReason).length;

  logInfo(`Replicate этап завершен: ${enhancedCount} улучшено, ${skippedHighQuality} пропущено (качественные), ${failedCount} ошибок`);

  return results;
}

function showReplicateFallbackNotification(enhancedCount, failedCount, totalCount) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = `РЕЗУЛЬТАТ УЛУЧШЕНИЯ ИЗОБРАЖЕНИЙ:\n\n`;
    
    if (enhancedCount > 0) {
      message += `✅ Улучшено через Replicate: ${enhancedCount} из ${totalCount}\n`;
    }
    
    if (failedCount > 0) {
      message += `⚠️ Без улучшения (исходные): ${failedCount} из ${totalCount}\n\n`;
      
      if (failedCount === totalCount) {
        message += `ВНИМАНИЕ: Все изображения обработаны БЕЗ улучшения!\n\n`;
        message += `Возможные причины:\n`;
        message += `• Проблемы с Replicate API\n`;
        message += `• Превышен лимит запросов\n`;
        message += `• Сетевые проблемы\n`;
        message += `• Неверный токен Replicate\n\n`;
        message += `Изображения всё равно будут оптимизированы в WebP формат.`;
      } else {
        message += `Частичные проблемы с Replicate API.\n`;
        message += `Обработка продолжена с исходными изображениями.`;
      }
    }
    
    const title = failedCount === totalCount ? '⚠️ Проблемы с улучшением' : 'ℹ️ Результат улучшения';
    
    ui.alert(title, message, ui.ButtonSet.OK);
    
    logInfo(`Replicate fallback notification: ${enhancedCount} enhanced, ${failedCount} failed`);
    
  } catch (error) {
    logError('Ошибка показа уведомления о fallback', error);
  }
}

/**
 * ЭТАП 3: Единая WebP оптимизация (фиксированное разрешение 3000px)
 */
async function processUnifiedWebPOptimization(enhancedImages, analysisResults) {
  const settings = getApiSettings();
  
  if (!settings.tinypngKey || !settings.imgbbKey) {
    logWarning('TinyPNG или ImgBB ключи недоступны, используем обработанные изображения');
    return enhancedImages.map((img, index) => {
      const analysis = analysisResults[index] || {
        altTag: `Изображение ${index + 1}`,
        seoFilename: `image-${index + 1}`
      };
      
      return {
        altTag: analysis.altTag,
        seoFilename: analysis.seoFilename,
        processedImageUrl: img.processed,
        confidence: img.wasEnhanced ? 7 : 5
      };
    });
  }

  // Увеличенный лимит: 15 изображений (~8 сек/шт = 2 мин)
  // Основная защита — таймаут MAX_TOTAL_TIME_MS ниже
  const MAX_WEBP_IMAGES = 15;
  const imagesToOptimize = enhancedImages.slice(0, MAX_WEBP_IMAGES);
  const skippedImages = enhancedImages.slice(MAX_WEBP_IMAGES);

  if (skippedImages.length > 0) {
    logWarning(`WebP: обрабатываем первые ${MAX_WEBP_IMAGES}, остальные ${skippedImages.length} получат только alt-теги`);
  }

  logInfo(`WebP оптимизация: ${imagesToOptimize.length} изображений`);

  const results = [];
  const startTime = Date.now();
  const MAX_TOTAL_TIME_MS = 240000; // 4 минуты максимум на весь этап (увеличено с 3)

  // ПОСЛЕДОВАТЕЛЬНАЯ обработка (вместо параллельной) для стабильности
  for (let index = 0; index < imagesToOptimize.length; index++) {
    // Проверка таймаута
    if (Date.now() - startTime > MAX_TOTAL_TIME_MS) {
      logWarning(`WebP таймаут после ${index} изображений. Остальные пропущены.`);
      // Добавляем оставшиеся как необработанные
      for (let j = index; j < imagesToOptimize.length; j++) {
        const analysis = analysisResults[j] || { altTag: `Изображение ${j + 1}`, seoFilename: `image-${j + 1}` };
        results.push({
          altTag: analysis.altTag,
          seoFilename: analysis.seoFilename,
          processedImageUrl: imagesToOptimize[j].processed,
          confidence: imagesToOptimize[j].wasEnhanced ? 6 : 4
        });
      }
      break;
    }

    const img = imagesToOptimize[index];
    const analysis = analysisResults[index] || { altTag: `Изображение ${index + 1}`, seoFilename: `image-${index + 1}` };

    try {
      logInfo(`WebP ${index + 1}/${imagesToOptimize.length}: обработка...`);

      const webpResult = await unifiedWebPConversion(img.processed, settings.tinypngKey);

      if (webpResult.success) {
        const finalUrl = await uploadToImgBB(webpResult.blob, settings.imgbbKey, analysis.seoFilename);

        if (finalUrl) {
          logInfo(`WebP ${index + 1}: OK (${webpResult.sizeKB}KB)`);
          results.push({
            altTag: analysis.altTag,
            seoFilename: analysis.seoFilename,
            processedImageUrl: finalUrl,
            confidence: 8
          });
          continue;
        }
      }

      // Fallback
      results.push({
        altTag: analysis.altTag,
        seoFilename: analysis.seoFilename,
        processedImageUrl: img.processed,
        confidence: img.wasEnhanced ? 6 : 4
      });

    } catch (error) {
      logError(`WebP ${index + 1}: ошибка`, error);
      results.push({
        altTag: analysis.altTag,
        seoFilename: analysis.seoFilename,
        processedImageUrl: img.processed,
        confidence: 3
      });
    }
  }

  // Добавляем пропущенные изображения (без WebP оптимизации)
  for (let i = 0; i < skippedImages.length; i++) {
    const idx = MAX_WEBP_IMAGES + i;
    const analysis = analysisResults[idx] || { altTag: `Изображение ${idx + 1}`, seoFilename: `image-${idx + 1}` };
    results.push({
      altTag: analysis.altTag,
      seoFilename: analysis.seoFilename,
      processedImageUrl: skippedImages[i].processed,
      confidence: skippedImages[i].wasEnhanced ? 5 : 3
    });
  }

  const optimizedCount = results.filter(r => r.confidence >= 7).length;
  const elapsedSec = Math.round((Date.now() - startTime) / 1000);
  logInfo(`WebP завершено за ${elapsedSec}с: ${optimizedCount}/${results.length} оптимизировано`);

  return results;
}

/**
 * Единая WebP конвертация с фиксированными параметрами
 */
async function unifiedWebPConversion(imageUrl, apiKey) {
  try {
    // Скачиваем изображение
    const imageResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    if (imageResponse.getResponseCode() !== 200) {
      return { success: false };
    }

    const imageBlob = imageResponse.getBlob();

    // Этап 1: Сжатие
    const shrinkResponse = UrlFetchApp.fetch('https://api.tinify.com/shrink', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`)
      },
      payload: imageBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (shrinkResponse.getResponseCode() !== 201) {
      return { success: false };
    }

    const shrinkData = JSON.parse(shrinkResponse.getContentText());

    // Этап 2: WebP конвертация с фиксированными параметрами
    const webpResponse = UrlFetchApp.fetch(shrinkData.output.url, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        convert: { type: "image/webp" },  // БЕЗ параметров качества
        resize: { method: "fit", width: 3000, height: 3000 }  // Фиксированное разрешение
      }),
      muteHttpExceptions: true
    });

    if (webpResponse.getResponseCode() === 200) {
      const webpBlob = webpResponse.getBlob();
      const sizeKB = Math.round(webpBlob.getBytes().length / 1024);
      
      return { success: true, blob: webpBlob, sizeKB: sizeKB };
    }

    return { success: false };
    
  } catch (error) {
    return { success: false };
  }
}

/**
 * ЭТАП 2: Пакетный анализ через OpenAI с разбивкой на группы по 5 изображений
 */
async function processBatchOpenAI(enhancedImages, productName) {
  try {
    const settings = getApiSettings();
    
    if (!settings.openaiAltAssistantId) {
      logWarning('OpenAI Assistant недоступен, генерируем базовые теги');
      return enhancedImages.map((img, i) => ({
        altTag: `${productName} - изображение ${i + 1}`,
        seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-')
      }));
    }

    const BATCH_SIZE = 5; // Максимум 5 изображений за раз
    const allResults = [];
    
    logInfo(`Разбиваем ${enhancedImages.length} изображений на группы по ${BATCH_SIZE}`);
    
    // Разбиваем на группы по 5
    for (let i = 0; i < enhancedImages.length; i += BATCH_SIZE) {
      const batch = enhancedImages.slice(i, i + BATCH_SIZE);
      const groupNumber = Math.floor(i / BATCH_SIZE) + 1;
      const totalGroups = Math.ceil(enhancedImages.length / BATCH_SIZE);
      
      logInfo(`Обрабатываем группу ${groupNumber}/${totalGroups}: ${batch.length} изображений`);
      
      try {
        // Создаем отдельный thread для каждой группы
        const thread = await createAssistantThread();
        
        // Формируем сообщение для группы
        const messageContent = [
          {
            type: "text",
            text: `Проанализируй ${batch.length} изображений товара "${productName}" и создай для каждого:
1. SEO-оптимизированный alt-тег (до 125 символов)
2. SEO-имя файла (только латиница, до 60 символов, БЕЗ расширения)

Отвечай в JSON формате:
{
  "results": [
    {"altTag": "alt текст 1", "seoFilename": "filename-1"},
    {"altTag": "alt текст 2", "seoFilename": "filename-2"}
  ]
}`
          }
        ];

        // Добавляем изображения группы
        batch.forEach((img, index) => {
          messageContent.push({
            type: "image_url",
            image_url: { url: img.enhanced }
          });
        });

        // Отправляем сообщение
        const messageResponse = UrlFetchApp.fetch(
          `https://api.openai.com/v1/threads/${thread.id}/messages`,
          {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${settings.openaiApiKey}`,
              'Content-Type': 'application/json',
              'OpenAI-Beta': 'assistants=v2'
            },
            payload: JSON.stringify({
              role: "user",
              content: messageContent
            }),
            muteHttpExceptions: true
          }
        );

        if (messageResponse.getResponseCode() !== 200) {
          throw new Error(`OpenAI message failed: ${messageResponse.getResponseCode()}`);
        }

        // Запускаем анализ группы
        const analysisResult = await runAssistantAnalysis(thread.id, settings.openaiAltAssistantId);
        const parsed = parseAssistantBatchResponse(analysisResult, batch.length, productName);
        
        logInfo(`Группа ${groupNumber} обработана: ${parsed.length} результатов`);
        allResults.push(...parsed);
        
        // Пауза между группами (кроме последней)
        if (i + BATCH_SIZE < enhancedImages.length) {
          logInfo('Пауза между группами...');
        }
        
      } catch (groupError) {
        logError(`Ошибка обработки группы ${groupNumber}`, groupError);
        
        // Fallback для текущей группы
        const fallbackResults = batch.map((img, index) => ({
          altTag: `${productName} - изображение ${i + index + 1}`,
          seoFilename: transliterate(`${productName}-${i + index + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
        }));
        
        allResults.push(...fallbackResults);
      }
    }
    
    logInfo(`Пакетный анализ завершен: ${allResults.length} результатов общим`);
    return allResults;
    
  } catch (error) {
    logError('Критическая ошибка пакетного OpenAI анализа', error);
    
    // Fallback - генерируем базовые теги для всех изображений
    return enhancedImages.map((img, i) => ({
      altTag: `${productName} - изображение ${i + 1}`,
      seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
    }));
  }
}

/**
 * ЭТАП 3: Параллельная WebP оптимизация
 */
async function processBatchOptimization(enhancedImages, analysisResults) {
  const settings = getApiSettings();
  
  if (!settings.tinypngKey || !settings.imgbbKey) {
    logWarning('TinyPNG или ImgBB ключи недоступны, используем исходные изображения');
    return enhancedImages.map((img, index) => {
      const analysis = analysisResults[index] || {
        altTag: `Изображение ${index + 1}`,
        seoFilename: `image-${index + 1}`
      };
      
      return {
        altTag: analysis.altTag,
        seoFilename: analysis.seoFilename,
        processedImageUrl: img.enhanced,
        confidence: 3
      };
    });
  }

  logInfo(`Запускаем параллельную WebP оптимизацию для ${enhancedImages.length} изображений`);
  
  // Создаем массив промисов для параллельной обработки
  const optimizationPromises = enhancedImages.map(async (img, index) => {
    const startTime = Date.now();
    
    try {
      const analysis = analysisResults[index] || {
        altTag: `Изображение ${index + 1}`,
        seoFilename: `image-${index + 1}`
      };

      logInfo(`Запускаем оптимизацию изображения ${index + 1}`);

      // Быстрая WebP оптимизация
      const optimized = await quickWebPOptimizationParallel(
        img.enhanced, 
        settings.tinypngKey, 
        settings.imgbbKey, 
        analysis.seoFilename,
        index + 1
      );
      
      const processingTime = Math.round((Date.now() - startTime) / 1000);
      
      if (optimized.success) {
        logInfo(`Изображение ${index + 1} оптимизировано за ${processingTime}сек: ${optimized.sizeKB}KB`);
        
        return {
          altTag: analysis.altTag,
          seoFilename: analysis.seoFilename,
          processedImageUrl: optimized.url,
          confidence: 8
        };
      } else {
        logWarning(`Оптимизация изображения ${index + 1} не удалась, используем исходное`);
        
        return {
          altTag: analysis.altTag,
          seoFilename: analysis.seoFilename,
          processedImageUrl: img.enhanced,
          confidence: 5
        };
      }
      
    } catch (error) {
      logError(`Ошибка оптимизации изображения ${index + 1}`, error);
      
      const analysis = analysisResults[index] || {
        altTag: `Изображение ${index + 1}`,
        seoFilename: `image-${index + 1}`
      };

      return {
        altTag: analysis.altTag,
        seoFilename: analysis.seoFilename,
        processedImageUrl: img.enhanced,
        confidence: 3
      };
    }
  });

  // Ждем завершения всех параллельных операций
  const results = await Promise.all(optimizationPromises);
  
  const successful = results.filter(r => r.confidence >= 7).length;
  logInfo(`Параллельная оптимизация завершена: ${successful}/${results.length} успешно оптимизированы`);
  
  return results;
}

/**
 * Улучшенная быстрая WebP оптимизация с контролем качества
 */
async function quickWebPOptimization(imageUrl, tinypngKey, imgbbKey, seoFilename) {
  try {
    // Скачиваем изображение
    const imageResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    if (imageResponse.getResponseCode() !== 200) {
      return { success: false };
    }

    const imageBlob = imageResponse.getBlob();
    const originalSizeKB = Math.round(imageBlob.getBytes().length / 1024);
    
    logInfo(`🔧 Начинаем качественную WebP оптимизацию: ${originalSizeKB}KB`);

    // Определяем стратегию на основе размера после Replicate
    const strategy = determineQualityStrategy(originalSizeKB);
    
    // WebP конвертация с контролем качества
    const webpResult = await convertToWebPQuality(imageBlob, tinypngKey, strategy);
    
    if (webpResult.success) {
      // Проверяем размер - должен быть 150-400KB
      if (webpResult.sizeKB < 150) {
        logInfo(`📐 Размер ${webpResult.sizeKB}KB слишком мал, увеличиваем качество`);
        
        // Повторная обработка с более высокими параметрами
        const betterStrategy = {
          width: strategy.width * 1.3,
          height: strategy.height * 1.3,
          quality: { min: 85, max: 95 }
        };
        
        const improvedResult = await convertToWebPQuality(imageBlob, tinypngKey, betterStrategy);
        if (improvedResult.success && improvedResult.sizeKB >= 150) {
          const finalUrl = await uploadToImgBB(improvedResult.blob, imgbbKey, seoFilename);
          if (finalUrl) {
            logInfo(`✅ Качественный WebP: ${improvedResult.sizeKB}KB`);
            return { success: true, url: finalUrl, sizeKB: improvedResult.sizeKB };
          }
        }
      }
      
      // Основной результат (150-400KB)
      if (webpResult.sizeKB <= 400) {
        const finalUrl = await uploadToImgBB(webpResult.blob, imgbbKey, seoFilename);
        if (finalUrl) {
          logInfo(`✅ Оптимальный WebP: ${webpResult.sizeKB}KB`);
          return { success: true, url: finalUrl, sizeKB: webpResult.sizeKB };
        }
      } else {
        // Слишком большой - применяем умеренное сжатие
        logInfo(`📉 Размер ${webpResult.sizeKB}KB превышает 400KB, применяем сжатие`);
        
        const compressedStrategy = {
          width: Math.round(strategy.width * 0.8),
          height: Math.round(strategy.height * 0.8),
          quality: { min: 75, max: 85 }
        };
        
        const compressedResult = await convertToWebPQuality(imageBlob, tinypngKey, compressedStrategy);
        if (compressedResult.success) {
          const finalUrl = await uploadToImgBB(compressedResult.blob, imgbbKey, seoFilename);
          if (finalUrl) {
            logInfo(`✅ Сжатый WebP: ${compressedResult.sizeKB}KB`);
            return { success: true, url: finalUrl, sizeKB: compressedResult.sizeKB };
          }
        }
      }
    }

    return { success: false };
    
  } catch (error) {
    logError(`Ошибка качественной оптимизации: ${error.message}`);
    return { success: false };
  }
}

/**
 * WebP конвертация с увеличенным разрешением БЕЗ параметров качества
 */
async function convertToWebPQuality(imageBlob, apiKey, strategy) {
  try {
    logInfo(`🎯 WebP параметры: ${strategy.width}x${strategy.height}, автоматическое качество`);
    
    // ЭТАП 1: Первичное сжатие (загружаем изображение)
    const shrinkResponse = UrlFetchApp.fetch('https://api.tinify.com/shrink', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`)
      },
      payload: imageBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (shrinkResponse.getResponseCode() !== 201) {
      logError(`❌ TinyPNG shrink failed: ${shrinkResponse.getResponseCode()}`);
      return { success: false };
    }

    const shrinkData = JSON.parse(shrinkResponse.getContentText());
    logInfo(`🔄 Первичное сжатие: ${Math.round(shrinkData.input.size / 1024)}KB → ${Math.round(shrinkData.output.size / 1024)}KB`);

    // ЭТАП 2: WebP конвертация БЕЗ параметров качества
    const webpPayload = {
      convert: {
        type: "image/webp"
        // Убираем все параметры качества
      },
      resize: {
        method: "fit",
        width: strategy.width,
        height: strategy.height
      }
    };

    const webpResponse = UrlFetchApp.fetch(shrinkData.output.url, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(webpPayload),
      muteHttpExceptions: true
    });

    if (webpResponse.getResponseCode() === 200) {
      const webpBlob = webpResponse.getBlob();
      const sizeKB = Math.round(webpBlob.getBytes().length / 1024);
      
      logInfo(`✅ WebP создан: ${sizeKB}KB`);
      return { success: true, blob: webpBlob, sizeKB: sizeKB };
    } else {
      logError(`❌ WebP conversion failed: ${webpResponse.getResponseCode()}`);
      return { success: false };
    }

  } catch (error) {
    logError(`💥 WebP conversion error: ${error.message}`);
    return { success: false };
  }
}

/**
 * Парсинг пакетного ответа от OpenAI Assistant
 */
function parseAssistantBatchResponse(response, expectedCount, productName) {
  try {
    // Пытаемся найти JSON в ответе
    const jsonMatch = response.match(/\{[\s\S]*"results"[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      if (parsed.results && Array.isArray(parsed.results)) {
        return parsed.results.map((result, i) => ({
          altTag: validateAltTag(result.altTag || `${productName} - изображение ${i + 1}`, productName),
          seoFilename: validateSeoFilename(result.seoFilename || `product-${i + 1}`)
        }));
      }
    }

    // Fallback парсинг
    const lines = response.split('\n').filter(line => line.trim());
    const results = [];
    
    for (let i = 0; i < expectedCount; i++) {
      results.push({
        altTag: `${productName} - изображение ${i + 1}`,
        seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
      });
    }
    
    return results;
    
  } catch (error) {
    logError('Ошибка парсинга пакетного ответа', error);
    
    // Создаем fallback результаты
    const results = [];
    for (let i = 0; i < expectedCount; i++) {
      results.push({
        altTag: `${productName} - изображение ${i + 1}`,
        seoFilename: transliterate(`${productName}-${i + 1}`).toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50)
      });
    }
    return results;
  }
}

/**
 * Анализ исходного изображения и выбор стратегии
 */
async function analyzeOriginalImage(imageUrl) {
  try {
    // Скачиваем изображение для определения размера
    const imageResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    
    if (imageResponse.getResponseCode() !== 200) {
      return { sizeKB: 0, strategy: 'fallback', needsEnhancement: false };
    }
    
    const sizeKB = Math.round(imageResponse.getBlob().getBytes().length / 1024);

    // Выбираем стратегию
    let strategy, needsEnhancement;

    if (sizeKB < 200) {          // было 100
      strategy = 'small_enhance';
      needsEnhancement = true;
    } else if (sizeKB <= 800) {   // было 500
      strategy = 'medium_enhance'; 
      needsEnhancement = true;
    } else if (sizeKB <= 1000) {  // было 1000
      strategy = 'large_optimize';
      needsEnhancement = false;
    } else {
      strategy = 'huge_compress';
      needsEnhancement = false;
    }

    logInfo(`Файл ${sizeKB}КБ: стратегия=${strategy}, улучшение=${needsEnhancement}`);

    return {
      sizeKB: sizeKB,
      strategy: strategy,
      needsEnhancement: needsEnhancement
    };

  } catch (error) {
    logError(`Ошибка анализа исходника: ${error.message}`);
    return { sizeKB: 0, strategy: 'fallback', needsEnhancement: false };
  }
}

/**
 * Умное улучшение изображения
 */
async function smartEnhanceImage(imageUrl, strategy) {
  try {
    const settings = getApiSettings();
    
    if (!settings.replicateToken) {
      logWarning('⚠️ Replicate Token не настроен, пропускаем улучшение');
      return { success: false, reason: 'no_token' };
    }

    // Параметры улучшения в зависимости от стратегии
    const enhanceParams = {
      'small_enhance': { scale: 4, model: 'esrgan' },    // Вернуть 2
      'medium_enhance': { scale: 4, model: 'esrgan' }    // Оставить 2
    };

    const params = enhanceParams[strategy] || enhanceParams['medium_enhance'];
    
    logInfo(`🚀 Запускаем Replicate ${params.model}, scale: ${params.scale}x`);

    logInfo(`Создаем payload для Replicate...`);
    const payload = {
      version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
      input: {
        image: imageUrl,
        scale: 4
      }
    };
    logInfo(`Payload создан, отправляем запрос...`);
    logInfo(`Replicate API URL: https://api.replicate.com/v1/predictions`);
    logInfo(`Replicate Token есть: ${!!settings.replicateToken}`);

    // Создаем prediction
    const predictionResponse = UrlFetchApp.fetch('https://api.replicate.com/v1/predictions', {
      method: 'POST',
      headers: {
        'Authorization': `Token ${settings.replicateToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
        input: {
          image: imageUrl,
          scale: 4
        }
      }),
      muteHttpExceptions: true
    });

    try {
      const responseCode = predictionResponse.getResponseCode();
      const responseText = predictionResponse.getContentText();
      logInfo(`Replicate HTTP код: ${responseCode}`);
      logInfo(`Replicate ответ: ${responseText}`);
    } catch (e) {
      logError(`Ошибка получения ответа: ${e.message}`);
    }

    logInfo(`Запрос отправлен, получаем ответ...`);

    if (predictionResponse.getResponseCode() !== 201) {
      return { success: false, reason: 'prediction_failed' };
    }

    const prediction = JSON.parse(predictionResponse.getContentText());

    // Ждем результат (максимум 45 секунд)
    const result = await waitForReplicateResult(prediction.id, settings.replicateToken, 45);
    
    if (result.success) {
      return {
        success: true,
        url: result.outputUrl,
        sizeKB: 200  // Фиксированное значение вместо проблемной функции
      };
    }

    return { success: false, reason: 'processing_failed' };

  } catch (error) {
    logError(`Ошибка улучшения: ${error.message}`);
    return { success: false, reason: error.message };
  }
}

/**
 * Ожидание результата от Replicate
 */
async function waitForReplicateResult(predictionId, token, maxWaitSeconds) {
  let attempts = 0;
  const maxAttempts = maxWaitSeconds;

  while (attempts < maxAttempts) {
    try {
      const statusResponse = UrlFetchApp.fetch(`https://api.replicate.com/v1/predictions/${predictionId}`, {
        headers: { 'Authorization': `Token ${token}` },
        muteHttpExceptions: true
      });

      const responseCode = statusResponse.getResponseCode();
      if (responseCode !== 200) {
        return { success: false, reason: `status_check_failed_${responseCode}` };
      }

      const statusData = JSON.parse(statusResponse.getContentText());

      if (statusData.status === 'succeeded' && statusData.output) {
        return { success: true, outputUrl: statusData.output };
      }

      if (statusData.status === 'failed') {
        // Детальная причина провала
        const errorMsg = statusData.error || statusData.logs || 'unknown_error';
        logWarning(`Replicate prediction failed: ${errorMsg}`);
        return { success: false, reason: `replicate_failed: ${errorMsg.substring(0, 100)}` };
      }

      if (statusData.status === 'canceled') {
        return { success: false, reason: 'canceled' };
      }

      // Ждем 2 секунды (вместо 1) чтобы уменьшить количество запросов
      Utilities.sleep(2000);
      attempts += 2;

    } catch (error) {
      return { success: false, reason: error.message };
    }
  }

  return { success: false, reason: `timeout_after_${maxWaitSeconds}s` };
}

/**
 * Умная WebP оптимизация
 */
async function smartWebPOptimization(imageUrl, currentSizeKB, config, seoFilename) {
  try {
    const settings = getApiSettings();
    
    if (!settings.tinypngKey) {
      return { success: false, reason: 'no_tinypng_key' };
    }

    // Определяем параметры WebP на основе текущего размера
    const webpParams = determineWebPParams(currentSizeKB, config);
    
    logInfo(`🎯 WebP параметры: ${webpParams.width}px x ${webpParams.height}px`);

    // Скачиваем изображение
    const imageResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    
    if (imageResponse.getResponseCode() !== 200) {
      return { success: false, reason: 'download_failed' };
    }

    const imageBlob = imageResponse.getBlob();

    // Прямая конвертация в WebP через TinyPNG
    let webpResult = await convertToWebPDirect(imageBlob, settings.tinypngKey, webpParams);
    
    if (webpResult.success) {
      // Проверяем минимальный размер
      if (webpResult.sizeKB < 150) {
        logInfo(`Размер ${webpResult.sizeKB}KB < 150KB, увеличиваем разрешение`);
        
        const largerParams = { width: 4500, height: 4500 };
        const largerResult = await convertToWebPDirect(imageBlob, settings.tinypngKey, largerParams);
        
        if (largerResult.success) {
          webpResult = largerResult;
          logInfo(`Новый размер: ${webpResult.sizeKB}KB`);
        }
      }
      if (webpResult.sizeKB <= config.MAX_SIZE_KB) {
        // Отлично! Загружаем финальную версию
        const finalUrl = await uploadToImgBB(webpResult.blob, settings.imgbbKey, seoFilename);
        
        return {
          success: true,
          finalUrl: finalUrl,
          finalSizeKB: webpResult.sizeKB,
          resolution: webpParams.width,
          confidence: 8
        };
      } else {
        // Размер превышает лимит - применяем дополнительную оптимизацию
        logInfo(`📐 Размер ${webpResult.sizeKB}KB > ${config.MAX_SIZE_KB}KB, применяем доп.сжатие`);
        
        const compressed = await applyFinalCompression(webpResult.blob, settings.tinypngKey, config.MAX_SIZE_KB);
        
        if (compressed.success) {
          const finalUrl = await uploadToImgBB(compressed.blob, settings.imgbbKey, seoFilename);
          
          return {
            success: true,
            finalUrl: finalUrl,
            finalSizeKB: compressed.sizeKB,
            resolution: compressed.resolution,
            confidence: 6
          };
        }
      }
    }

    return { success: false, reason: 'webp_conversion_failed' };

  } catch (error) {
    logError(`Ошибка WebP оптимизации: ${error.message}`);
    return { success: false, reason: error.message };
  }
}

/**
 * Определение параметров WebP для достижения 150-300 КБ
 */
function determineWebPParams(sizeKB, config) {
 if (sizeKB <= config.SMALL_IMAGE_KB) {
   // Маленькое - увеличиваем разрешение для достижения 150-300КБ
   return {
     width: 4000,
     height: 4000
   };
 } else if (sizeKB <= config.MEDIUM_IMAGE_KB) {
   // Среднее - высокое разрешение
   return {
     width: 2600,
     height: 2600
   };
 } else {
   // Большое - умеренное разрешение
   return {
     width: 3000,
     height: 3000
   };
 }
}

/**
 * Исправленная функция WebP конвертации с правильным двухэтапным процессом
 */
async function convertToWebPDirect(imageBlob, apiKey, params) {
  try {
    logInfo(`🔧 Начинаем WebP конвертацию: ${Math.round(imageBlob.getBytes().length / 1024)}KB`);
    
    // ЭТАП 1: Первичное сжатие (загружаем изображение)
    const shrinkResponse = UrlFetchApp.fetch('https://api.tinify.com/shrink', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`)
      },
      payload: imageBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (shrinkResponse.getResponseCode() !== 201) {
      logError(`❌ TinyPNG shrink failed: ${shrinkResponse.getResponseCode()}`);
      return { success: false, reason: 'shrink_failed' };
    }

    const shrinkData = JSON.parse(shrinkResponse.getContentText());
    logInfo(`✅ Первичное сжатие: ${Math.round(shrinkData.input.size / 1024)}KB → ${Math.round(shrinkData.output.size / 1024)}KB`);

    // ЭТАП 2: WebP конвертация с параметрами качества
      const webpPayload = {
        convert: {
          type: "image/webp",
        },
        resize: {
          method: "fit",
          width: params.width,
          height: params.height
        }
      };

    const webpResponse = UrlFetchApp.fetch(shrinkData.output.url, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(webpPayload),
      muteHttpExceptions: true
    });

    if (webpResponse.getResponseCode() === 200) {
      const webpBlob = webpResponse.getBlob();
      const sizeKB = Math.round(webpBlob.getBytes().length / 1024);
      
      logInfo(`✅ WebP создан: ${sizeKB}KB`);
      return { success: true, blob: webpBlob, sizeKB: sizeKB };
    } else {
      logError(`❌ WebP conversion failed: ${webpResponse.getResponseCode()}`);
      return { success: false, reason: 'webp_conversion_failed' };
    }

  } catch (error) {
    logError(`💥 WebP conversion error: ${error.message}`);
    return { success: false, reason: error.message };
  }
}

/**
 * Финальное сжатие при превышении лимита
 */
async function applyFinalCompression(imageBlob, apiKey, maxSizeKB) {
  try {
    logInfo('🗜️ Применяем финальное сжатие');

    // Более агрессивные параметры
    const payload = {
      convert: {
        type: "image/webp",
        quality: { min: 75, max: 85 }
      },
      resize: {
        method: "fit",
        width: 1200,
        height: 1200
      }
    };

    const response = UrlFetchApp.fetch('https://api.tinify.com/shrink', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`)
      },
      payload: imageBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 201) {
      return { success: false, reason: 'final_shrink_failed' };
    }

    const shrinkData = JSON.parse(response.getContentText());

    const finalResponse = UrlFetchApp.fetch(shrinkData.output.url, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (finalResponse.getResponseCode() === 200) {
      const finalBlob = finalResponse.getBlob();
      const finalSize = Math.round(finalBlob.getBytes().length / 1024);

      return {
        success: true,
        blob: finalBlob,
        sizeKB: finalSize,
        resolution: 1000
      };
    }

    return { success: false, reason: 'final_conversion_failed' };

  } catch (error) {
    return { success: false, reason: error.message };
  }
}

/**
 * Получение размера изображения в KB
 */
async function getImageSizeKB(imageUrl) {
  try {
    const response = UrlFetchApp.fetch(imageUrl, { 
      method: 'HEAD',
      muteHttpExceptions: true 
    });
    
    const contentLength = response.getHeaders()['Content-Length'] || 
                         response.getHeaders()['content-length'];
    
    return contentLength ? Math.round(parseInt(contentLength) / 1024) : 0;
    
  } catch (error) {
    return 0;
  }
}

/**
 * Улучшенная генерация alt-тега
 */
async function generateAltTagWithAssistant(imageUrl, productName) {
  try {
    const settings = getApiSettings();
    
    if (!settings.openaiAltAssistantId) {
      return `${productName} - изображение товара`;
    }

    const thread = await createAssistantThread();
    await sendImageToAssistant(thread.id, imageUrl, productName);
    const result = await runAssistantAnalysis(thread.id, settings.openaiAltAssistantId);
    
    const parsed = parseAssistantResponse(result);
    return parsed.alt_tag || `${productName} - изображение товара`;
    
  } catch (error) {
    logError(`Ошибка Assistant: ${error.message}`);
    return `${productName} - изображение товара`;
  }
}

/**
 * Генерация SEO filename
 */
async function generateSeoFilenameWithAssistant(imageUrl, productName, altTag) {
 try {
   const settings = getApiSettings();
  
   if (!settings.openaiAltAssistantId) {
     // Fallback - транслитерация
     const cleaned = transliterate(altTag || productName)
       .toLowerCase()
       .replace(/[^a-z0-9]/g, '-')
       .replace(/-+/g, '-')
       .replace(/^-|-$/g, '')
       .substring(0, 50);
     return cleaned || `product-${Date.now()}`;
   }

   // Используем тот же thread что и для alt-тега
   const thread = await createAssistantThread();
   await sendImageToAssistant(thread.id, imageUrl, `Создай SEO-имя файла для товара "${productName}". Только латиница, максимум 80 символов. БЕЗ расширения файла.`);
   const result = await runAssistantAnalysis(thread.id, settings.openaiAltAssistantId);
  
   const parsed = parseAssistantResponse(result);
   const seoName = parsed.seo_filename || parsed.filename || '';
   
   if (seoName) {
     return validateSeoFilename(seoName);
   }
   
   // Fallback
   const cleaned = transliterate(altTag || productName)
     .toLowerCase()
     .replace(/[^a-z0-9]/g, '-')
     .replace(/-+/g, '-')
     .replace(/^-|-$/g, '')
     .substring(0, 50);
   return cleaned || `product-${Date.now()}`;
   
 } catch (error) {
   logError(`Ошибка генерации SEO filename: ${error.message}`);
   // Fallback
   const cleaned = transliterate(altTag || productName)
     .toLowerCase()
     .replace(/[^a-z0-9]/g, '-')
     .replace(/-+/g, '-')
     .replace(/^-|-$/g, '')
     .substring(0, 50);
   return cleaned || `product-${Date.now()}`;
 }
}

/**
 * Fallback результат при ошибках
 */
async function createFallbackResult(imageUrl, productName) {
  return {
    altTag: `${productName} - изображение товара`,
    seoFilename: `product-${Date.now()}`,
    processedImageUrl: imageUrl, // Возвращаем исходное изображение
    confidence: 3
  };
}

async function uploadToImgBB(webpBlob, imgbbKey, seoFilename) {
  try {
    const WORKER_URL = 'https://misty-leaf-2392.eugeny-ermackow.workers.dev';
    
    const fileName = `${seoFilename || Date.now() + '-optimized'}`;
    const base64Data = Utilities.base64Encode(webpBlob.getBytes());
    
    logInfo(`Загружаем через Cloudflare Worker: ${Math.round(webpBlob.getBytes().length / 1024)}KB`);
    
    const response = UrlFetchApp.fetch(WORKER_URL, {
      method: 'POST',
      payload: {
        image: base64Data,
        name: fileName
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Worker HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }
    
    const result = JSON.parse(response.getContentText());
    if (result.success) {
      logInfo(`Загружено через Worker: ${result.data.url}`);
      return result.data.url;
    }
    
    throw new Error(`Worker error: ${result.error?.message || 'Unknown'}`);
    
  } catch (error) {
    logError(`Ошибка Worker загрузки: ${error.message}`);
    return null;
  }
}

/**
 * Улучшение изображения через Replicate API
 */
async function enhanceWithReplicate(imageUrl) {
  try {
    const settings = getApiSettings();
    
    if (!settings.replicateToken) {
      logWarning('Replicate Token не настроен, пропускаем улучшение');
      return imageUrl; // Возвращаем исходное изображение
    }

    logInfo(`Улучшаем изображение через Replicate: ${imageUrl}`);

    // Создаем prediction для улучшения изображения
    const predictionResponse = UrlFetchApp.fetch('https://api.replicate.com/v1/predictions', {
      method: 'POST',
      headers: {
        'Authorization': `Token ${settings.replicateToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        version: "f121d640bd286e1fdc67f9799164c1d5be36ff74576ee11c803ae5b665dd46aa",
        input: {
          image: imageUrl,
          scale: 4,
          face_enhance: false
        }
      }),
      muteHttpExceptions: true
    });

    if (predictionResponse.getResponseCode() !== 201) {
      throw new Error(`Replicate API error: ${predictionResponse.getResponseCode()}`);
    }

    const prediction = JSON.parse(predictionResponse.getContentText());
    const predictionId = prediction.id;

    // Ждем завершения обработки
    let status = prediction.status;
    let attempts = 0;
    const maxAttempts = 30; // 30 секунд максимум

    while (status === 'starting' || status === 'processing') {
      Utilities.sleep(1000);
      
      const statusResponse = UrlFetchApp.fetch(`https://api.replicate.com/v1/predictions/${predictionId}`, {
        headers: {
          'Authorization': `Token ${settings.replicateToken}`
        },
        muteHttpExceptions: true
      });

      const statusData = JSON.parse(statusResponse.getContentText());
      status = statusData.status;

      if (status === 'succeeded' && statusData.output) {
        // ДОБАВИТЬ ЭТИ СТРОКИ:
        logInfo(`Replicate результат URL: ${statusData.output}`);
        
        // Проверяем формат файла
        const isWebP = statusData.output.toLowerCase().includes('.webp');
        logInfo(`Replicate выдает WebP: ${isWebP}`);
        
        logInfo(`Изображение улучшено через Replicate: ${statusData.output}`);
        return statusData.output;
      }

      attempts++;
      if (attempts >= maxAttempts) {
        throw new Error('Таймаут ожидания Replicate');
      }
    }

    if (status === 'failed') {
      throw new Error('Replicate обработка не удалась');
    }

    // Fallback - возвращаем исходное изображение
    return imageUrl;

  } catch (error) {
    logError(`Ошибка Replicate: ${error.message}`);
    // В случае ошибки возвращаем исходное изображение
    return imageUrl;
  }
}

/**
 * Оптимизация изображения через TinyPNG до 300KB в WebP
 */
/**
 * Оптимизация изображения через TinyPNG до 300KB в WebP (ИСПРАВЛЕНО)
 */
/**
 * Оптимизация изображения через TinyPNG - прямая конвертация в WebP
 */
async function optimizeWithTinyPNG(imageUrl) {
  try {
    const settings = getApiSettings();
    
    if (!settings.tinypngKey) {
      logWarning('TinyPNG ключ не настроен, пропускаем оптимизацию');
      return imageUrl;
    }

    logInfo(`Оптимизируем изображение через TinyPNG: ${imageUrl}`);

    // Скачиваем улучшенное изображение от Replicate
    const imageResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    
    if (imageResponse.getResponseCode() !== 200) {
      throw new Error(`Не удалось скачать изображение: ${imageResponse.getResponseCode()}`);
    }

    const imageBlob = imageResponse.getBlob();
    const originalSize = imageBlob.getBytes().length;
    logInfo(`Размер после Replicate: ${Math.round(originalSize / 1024)} KB`);

    // Прямая конвертация в WebP с оптимизацией размера (один этап)
    const optimizeResponse = UrlFetchApp.fetch('https://api.tinify.com/shrink', {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${settings.tinypngKey}`)
      },
      payload: imageBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (optimizeResponse.getResponseCode() !== 201) {
      throw new Error(`TinyPNG error: ${optimizeResponse.getResponseCode()}`);
    }

    const optimizeData = JSON.parse(optimizeResponse.getContentText());
    const compressedUrl = optimizeData.output.url;

    // Конвертируем в WebP с контролем размера
    const webpPayload = {
      convert: {
        type: "image/webp"
      },
      resize: {
        method: "fit",
        width: 1400,
        height: 1400
      }
    };

    const webpResponse = UrlFetchApp.fetch(compressedUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${settings.tinypngKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(webpPayload),
      muteHttpExceptions: true
    });

    if (webpResponse.getResponseCode() === 200) {
      const webpBlob = webpResponse.getBlob();
      const finalSize = webpBlob.getBytes().length;
      
      logInfo(`Финальный WebP размер: ${Math.round(finalSize / 1024)} KB`);
      
      // Если размер больше 300KB, уменьшаем разрешение
      if (finalSize > 307200) {
        logInfo('Размер > 300KB, уменьшаем разрешение');
        return await resizeWebPSmaller(compressedUrl, settings.tinypngKey);
      }
      
      // Временно загружаем на ImgBB для получения прямой ссылки
      return await uploadWebPToImgBB(webpBlob, settings.imgbbKey);
      
    } else {
      throw new Error(`WebP конвертация не удалась: ${webpResponse.getResponseCode()}`);
    }
    
  } catch (error) {
    logError(`Ошибка TinyPNG оптимизации: ${error.message}`);
    return imageUrl; // Fallback к улучшенному Replicate изображению
  }
}

/**
 * Уменьшение разрешения WebP для достижения 300KB
 */
async function resizeWebPSmaller(imageUrl, apiKey) {
  const smallerPayload = {
    convert: {
      type: "image/webp"
    },
    resize: {
      method: "fit",
      width: 1000,
      height: 1000
    }
  };

  try {
    const response = UrlFetchApp.fetch(imageUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(smallerPayload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const smallerBlob = response.getBlob();
      logInfo(`Размер после уменьшения: ${Math.round(smallerBlob.getBytes().length / 1024)} KB`);
      
      const settings = getApiSettings();
      return await uploadWebPToImgBB(smallerBlob, settings.imgbbKey);
    }
    
    return imageUrl;
    
  } catch (error) {
    logError(`Ошибка уменьшения WebP: ${error.message}`);
    return imageUrl;
  }
}

/**
 * Загрузка WebP на ImgBB
 */
async function uploadWebPToImgBB(webpBlob, imgbbKey) {
  try {
    const base64Image = Utilities.base64Encode(webpBlob.getBytes());
    
    const uploadResponse = UrlFetchApp.fetch('https://api.imgbb.com/1/upload', {
      method: 'POST',
      payload: {
        key: imgbbKey,
        image: base64Image,
        name: `optimized-webp-${Date.now()}`
      },
      muteHttpExceptions: true
    });

    if (uploadResponse.getResponseCode() === 200) {
      const result = JSON.parse(uploadResponse.getContentText());
      if (result.success) {
        logInfo(`WebP загружен на ImgBB: ${Math.round(result.data.size / 1024)} KB`);
        return result.data.url;
      }
    }
    
    throw new Error('ImgBB загрузка не удалась');
    
  } catch (error) {
    logError(`Ошибка загрузки WebP на ImgBB: ${error.message}`);
    throw error;
  }
}

/**
 * Дополнительное сжатие WebP файла
 */
async function compressWebPFurther(imageUrl, apiKey) {
  try {
    const furtherPayload = {
      convert: {
        type: "image/webp"
      },
      resize: {
        method: "fit", 
        width: 800,
        height: 800
      }
    };

    const response = UrlFetchApp.fetch(imageUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(furtherPayload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const finalSize = response.getBlob().getBytes().length;
      logInfo(`Размер после дополнительного сжатия: ${Math.round(finalSize / 1024)} KB`);
      
      return await uploadOptimizedToTempService(response.getBlob());
    }
    
    return imageUrl;
    
  } catch (error) {
    logError(`Ошибка дополнительного сжатия: ${error.message}`);
    return imageUrl;
  }
}

/**
 * Временная загрузка через ImgBB для получения прямой ссылки
 */
async function uploadOptimizedToTempService(optimizedBlob) {
  try {
    const settings = getApiSettings();
    
    if (!settings.imgbbKey) {
      logWarning('ImgBB ключ недоступен для временной загрузки');
      return null;
    }

    const base64Image = Utilities.base64Encode(optimizedBlob.getBytes());
    
    const uploadResponse = UrlFetchApp.fetch('https://api.imgbb.com/1/upload', {
      method: 'POST',
      payload: {
        key: settings.imgbbKey,
        image: base64Image,
        name: `temp-optimized-${Date.now()}`
      },
      muteHttpExceptions: true
    });

    if (uploadResponse.getResponseCode() === 200) {
      const result = JSON.parse(uploadResponse.getContentText());
      if (result.success) {
        return result.data.url;
      }
    }
    
    return null;
    
  } catch (error) {
    logError(`Ошибка временной загрузки: ${error.message}`);
    return null;
  }
}

/**
 * Дополнительное уменьшение качества для достижения 300KB
 */
async function reduceQualityFurther(imageUrl, apiKey) {
  try {
    const reduceResponse = UrlFetchApp.fetch(imageUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`api:${apiKey}`)
      },
      payload: JSON.stringify({
        convert: {
          type: ["image/webp"],
          quality: {
            min: 60,
            max: 80
          }
        },
        resize: {
          method: "fit",
          width: 1000,
          height: 1000
        }
      }),
      muteHttpExceptions: true
    });

    if (reduceResponse.getResponseCode() === 200) {
      const finalBlob = reduceResponse.getBlob();
      logInfo(`Размер после дополнительной оптимизации: ${Math.round(finalBlob.getBytes().length / 1024)} KB`);
      return createTempUrl(finalBlob);
    }
    
    return imageUrl;
    
  } catch (error) {
    logError(`Ошибка дополнительной оптимизации: ${error.message}`);
    return imageUrl;
  }
}

/**
 * Создание временного URL для blob
 */
function createTempUrl(blob) {
  try {
    // Конвертируем blob в base64 data URL
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64}`;
  } catch (error) {
    logError(`Ошибка создания temp URL: ${error.message}`);
    return null;
  }
}

function testUploadSmallPng() {
  const imgbbKey = 'a6edc9fb34fbe3f5590da5a29cfc0573';

  // Попробуем найти любой файл-изображение
  const it = DriveApp.searchFiles(
    "mimeType contains 'image/' and trashed = false"
  );

  let blob;
  if (it.hasNext()) {
    const file = it.next();
    blob = file.getBlob();
    logInfo(`Нашёл в Drive: ${file.getName()} (${blob.getContentType()})`);
  } else {
    // Фолбэк: 1×1 PNG (прозрачный), встроенный в код
    logInfo('В Drive изображений не найдено — используем встроенный tiny PNG');
    const tinyPngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6kH8kAAAAASUVORK5CYII=';
    blob = Utilities.newBlob(Utilities.base64Decode(tinyPngBase64), 'image/png', 'tiny.png');
  }

  const url = uploadToImgBB(blob, imgbbKey, 'test-small');
  Logger.log('ImgBB URL: ' + url);
}

function debugImgbbKeyProps() {
  const sp = PropertiesService.getScriptProperties().getProperties();
  const up = PropertiesService.getUserProperties().getProperties();
  const dp = PropertiesService.getDocumentProperties().getProperties();

  function mask(v) {
    if (!v) return '(empty)';
    const s = v.trim();
    return s.length <= 8 ? s : s.slice(0,4) + '…' + s.slice(-4) + ` (len:${s.length})`;
  }
  function dump(props, label) {
    logInfo(`--- ${label} ---`);
    Object.keys(props).sort().forEach(k => {
      const v = props[k];
      logInfo(`${k} = ${mask(v)}`);
    });
  }

  dump(sp, 'SCRIPT PROPERTIES');
  dump(up, 'USER PROPERTIES');
  dump(dp, 'DOCUMENT PROPERTIES');

  // Явно получим ключ из Script Properties по ожидаемому имени
  const raw = (PropertiesService.getScriptProperties().getProperty('ImgBBKey') || '').toString();
  const trimmed = raw.replace(/\s+/g, '');
  logInfo(`ImgBBKey (raw masked): ${mask(raw)}`);
  logInfo(`ImgBBKey (no-whitespace masked): ${mask(trimmed)}`);

  // Простая валидация формата
  const ok = /^[A-Za-z0-9]{32}$/.test(trimmed);
  logInfo(`Format valid: ${ok}`);
}

/**
 * Объединение всех изображений товара с автопарсингом
 * Берет изображения из InSales + от поставщика + дополнительные (выбранные пользователем)
 */
async function combineAllImages(originalImages, supplierImages, additionalImages, article) {
  try {
    const allImages = [];
    
    // ПРИОРИТЕТ 1: Дополнительные изображения (выбранные пользователем в диалоге)
    if (additionalImages) {
      const additionalUrls = additionalImages.split(/[\n,]/)
        .map(url => url.trim())
        .filter(url => url && url.startsWith('http'));
      allImages.push(...additionalUrls);
      logInfo(`Добавлены выбранные пользователем: ${additionalUrls.length} изображений`);
    }
    
    // ПРИОРИТЕТ 2: Если нет выбранных, берем изображения поставщика
    if (allImages.length === 0 && supplierImages) {
      const supplierUrls = supplierImages.split(/[\n,]/)
        .map(url => url.trim())
        .filter(url => url && url.startsWith('http'));
      allImages.push(...supplierUrls);
      logInfo(`Добавлены изображения поставщика: ${supplierUrls.length}`);
    }
    
    // ПРИОРИТЕТ 3: Если и поставщика нет, берем исходные из InSales
    if (allImages.length === 0 && originalImages) {
      const insalesUrls = originalImages.split(/[\n,]/)
        .map(url => url.trim())
        .filter(url => url && url.startsWith('http'));
      allImages.push(...insalesUrls);
      logInfo(`Добавлены исходные из InSales: ${insalesUrls.length}`);
    }
    
    // Удаляем дубликаты
    const uniqueImages = [...new Set(allImages)];
    
    logInfo(`Итого для обработки: ${uniqueImages.length} уникальных изображений`);
    
    return uniqueImages;
    
  } catch (error) {
    logError('Ошибка объединения изображений', error);
    return [];
  }
}

/**
 * Обработка выбранных изображений из диалога
 * Вызывается из HTML-диалога после выбора пользователем
 */
function processSelectedImageUrls(selectedUrls) {
  try {
    logInfo(`Получены выбранные изображения: ${selectedUrls.length} шт`);
    
    // Получаем выбранные товары
    const products = getProductsForProcessing();
    
    if (products.length === 0) {
      throw new Error('Нет выбранных товаров');
    }
    
    // Работаем с первым выбранным товаром
    const product = products[0];
    
    // Сохраняем выбранные изображения в колонку G (Дополнительные изображения)
    const selectedImagesText = selectedUrls.join('\n');
    updateAdditionalImages(product.article, selectedImagesText);
    
    logInfo(`Сохранены выбранные изображения для товара: ${product.article}`);
    
    // Запускаем обработку товара с выбранными изображениями
    processSelectedImages();
    
    return { success: true, count: selectedUrls.length };
    
  } catch (error) {
    logError('Ошибка обработки выбранных изображений', error);
    throw error;
  }
}

/**
 * Обновление колонки "Дополнительные изображения" (G)
 */
function updateAdditionalImages(article, imagesText) {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === article) { // Колонка B - артикул
        sheet.getRange(i + 1, IMAGES_COLUMNS.ADDITIONAL_IMAGES).setValue(imagesText);
        logInfo(`Обновлены дополнительные изображения для ${article}`);
        return true;
      }
    }
    
    throw new Error(`Товар с артикулом ${article} не найден`);
    
  } catch (error) {
    logError('Ошибка обновления дополнительных изображений', error);
    throw error;
  }
}

/**
 * ПОКАЗАТЬ ФИНАЛЬНЫЕ РЕЗУЛЬТАТЫ ОБРАБОТКИ
 * 
 * @param {Array} enhancedImages - Результаты Replicate
 * @param {Array} finalResults - Финальные результаты после WebP
 * @param {string} productName - Название товара
 */
function showFinalProcessingResults(enhancedImages, finalResults, productName) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Анализируем результаты
    const totalImages = enhancedImages.length;
    const replicateEnhanced = enhancedImages.filter(img => img.wasEnhanced).length;
    const replicateFailed = totalImages - replicateEnhanced;
    const webpOptimized = finalResults.filter(r => r.confidence >= 7).length;
    const webpFailed = totalImages - webpOptimized;
    
    let message = `РЕЗУЛЬТАТЫ ОБРАБОТКИ "${productName}":\n\n`;
    
    // Этап Replicate
    message += `🔧 ЭТАП УЛУЧШЕНИЯ (Replicate):\n`;
    if (replicateEnhanced > 0) {
      message += `✅ Улучшено: ${replicateEnhanced} из ${totalImages}\n`;
    }
    if (replicateFailed > 0) {
      message += `⚠️ Без улучшения: ${replicateFailed} из ${totalImages}\n`;
    }
    
    // Этап WebP оптимизации
    message += `\n🎯 ЭТАП ОПТИМИЗАЦИИ (WebP):\n`;
    if (webpOptimized > 0) {
      message += `✅ Оптимизировано: ${webpOptimized} из ${totalImages}\n`;
    }
    if (webpFailed > 0) {
      message += `⚠️ Базовая обработка: ${webpFailed} из ${totalImages}\n`;
    }
    
    // Общий результат
    message += `\n📊 ОБЩИЙ РЕЗУЛЬТАТ:\n`;
    message += `• Всего обработано: ${totalImages} изображений\n`;
    message += `• Alt-теги созданы: ${finalResults.length}\n`;
    message += `• SEO-имена созданы: ${finalResults.length}\n`;
    
    // Предупреждения
    if (replicateFailed === totalImages) {
      message += `\n⚠️ ВНИМАНИЕ: Все изображения обработаны без улучшения Replicate.\n`;
      message += `Возможные причины:\n`;
      message += `• Проблемы с Replicate API\n`;
      message += `• Превышен лимит запросов\n`;
      message += `• Неверные настройки модели\n`;
    } else if (replicateFailed > 0) {
      message += `\n💡 ${replicateFailed} изображений обработаны без улучшения.`;
    }
    
    // Определяем тип уведомления
    const hasIssues = (replicateFailed > 0) || (webpFailed > 0);
    const title = hasIssues ? '⚠️ Обработка завершена с замечаниями' : '✅ Обработка успешно завершена';
    
    ui.alert(title, message, ui.ButtonSet.OK);
    
    logInfo(`Final processing notification shown for ${productName}: ${replicateEnhanced} enhanced, ${webpOptimized} optimized`);
    
  } catch (error) {
    logError('Ошибка показа финальных результатов', error);
  }
}