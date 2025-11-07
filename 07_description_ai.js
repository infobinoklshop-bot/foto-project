/**
 * ========================================
 * МОДУЛЬ 07: AI-РЕРАЙТ ОПИСАНИЙ
 * ========================================
 *
 * Генерация уникальных описаний товаров с использованием OpenAI Assistants API
 * Адаптировано из проекта /Users/evgenijermakov/Documents/Description
 */

// =============================================================================
// КОНСТАНТЫ АССИСТЕНТОВ
// =============================================================================

const AI_ASSISTANTS = {
  COPIER: 'asst_qVFYH8Q5qgzMKsvOOzwnAuHN',  // Копирайтер
  CRITIC: 'asst_4NCAFLDo2mOh8CJNKFKt5Qvu',  // Критик
  EDITOR: 'asst_QTKbLGXgFwDWrJasQpC5It1Z'   // Редактор
};

// Имена для логирования
const ASSISTANT_NAMES = {
  [AI_ASSISTANTS.COPIER]: 'КОПИРАЙТЕР',
  [AI_ASSISTANTS.CRITIC]: 'КРИТИК',
  [AI_ASSISTANTS.EDITOR]: 'РЕДАКТОР'
};

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * ГЕНЕРАЦИЯ ОПИСАНИЯ ТОВАРА
 *
 * @param {Object} productData - Данные товара
 * @param {string} productData.article - Артикул
 * @param {string} productData.productName - Название
 * @param {string} productData.description - Исходное описание от поставщика
 * @param {Object} productData.specifications - Нормализованные характеристики
 * @param {string} productData.brand - Бренд
 * @param {string} productData.categories - Категории
 * @returns {Object} { rewrittenDescription, shortDescription, quality }
 */
function generateProductDescription(productData) {
  try {
    logInfo(`🤖 Генерируем описание для ${productData.article}`);

    // Валидация входных данных
    if (!productData.description && !productData.productName) {
      throw new Error('Нет исходных данных для генерации описания');
    }

    // Формируем промпт для копирайтера
    const prompt = buildDescriptionPrompt(productData);

    // Генерируем описание через копирайтера
    const rawDescription = callOpenAIAssistant(
      prompt,
      AI_ASSISTANTS.COPIER,
      getOpenAIKey()
    );

    // Редактируем описание для улучшения качества
    const editedDescription = editDescription(rawDescription, productData);

    // Генерируем краткое описание (макс 250 символов)
    const shortDescription = generateShortDescription(editedDescription);

    // Оцениваем качество
    const quality = assessDescriptionQuality(editedDescription);

    logInfo(`✅ Описание сгенерировано: ${editedDescription.length} символов, качество: ${quality.score}%`);

    return {
      rewrittenDescription: editedDescription,
      shortDescription: shortDescription,
      quality: quality
    };

  } catch (error) {
    handleError(error, 'Генерация описания AI');
    return {
      rewrittenDescription: '',
      shortDescription: '',
      quality: { score: 0, errors: [error.message] }
    };
  }
}

/**
 * ФОРМИРОВАНИЕ ПРОМПТА ДЛЯ КОПИРАЙТЕРА
 */
function buildDescriptionPrompt(productData) {
  const parts = [];

  parts.push('Напиши уникальное SEO-описание товара для интернет-магазина.');
  parts.push('');
  parts.push(`ТОВАР: ${productData.productName || 'без названия'}`);
  parts.push(`АРТИКУЛ: ${productData.article || '-'}`);

  if (productData.brand) {
    parts.push(`БРЕНД: ${productData.brand}`);
  }

  if (productData.categories) {
    parts.push(`КАТЕГОРИЯ: ${productData.categories}`);
  }

  if (productData.description) {
    parts.push('');
    parts.push('ИСХОДНОЕ ОПИСАНИЕ ОТ ПОСТАВЩИКА:');
    parts.push(productData.description);
  }

  if (productData.specifications && Object.keys(productData.specifications).length > 0) {
    parts.push('');
    parts.push('ХАРАКТЕРИСТИКИ:');
    for (const [key, value] of Object.entries(productData.specifications)) {
      parts.push(`- ${key}: ${value}`);
    }
  }

  parts.push('');
  parts.push('ТРЕБОВАНИЯ:');
  parts.push('1. Описание должно быть уникальным (не копировать исходное)');
  parts.push('2. Длина 800-1200 символов');
  parts.push('3. Естественный стиль, без академичности и штампов');
  parts.push('4. Включить ключевые характеристики');
  parts.push('5. Подчеркнуть преимущества и применение товара');
  parts.push('6. Без markdown-разметки, только чистый текст');
  parts.push('7. Без слов "инновационный", "передовой", "революционный"');

  return parts.join('\n');
}

/**
 * РЕДАКТИРОВАНИЕ ОПИСАНИЯ
 */
function editDescription(rawDescription, productData) {
  try {
    logInfo('✏️ Редактируем описание через EDITOR');

    const editorPrompt = `Отредактируй это описание товара, сделай его более естественным и уникальным.

ОПИСАНИЕ:
${rawDescription}

ТОВАР: ${productData.productName}

ТРЕБОВАНИЯ:
1. Убери штампы и академичность
2. Сделай текст более живым и продающим
3. Сохрани ключевые характеристики
4. Длина 800-1200 символов
5. Без markdown-разметки
6. Проверь на "тошноту" - избегай повторов одних и тех же слов`;

    const editedText = callOpenAIAssistant(
      editorPrompt,
      AI_ASSISTANTS.EDITOR,
      getOpenAIKey()
    );

    return cleanMarkdownFromResult(editedText);

  } catch (error) {
    logWarning(`⚠️ Ошибка редактирования, возвращаем исходное: ${error.message}`);
    return rawDescription;
  }
}

/**
 * ГЕНЕРАЦИЯ КРАТКОГО ОПИСАНИЯ
 */
function generateShortDescription(fullDescription) {
  try {
    if (fullDescription.length <= 250) {
      return fullDescription;
    }

    // Берем первые 2-3 предложения
    const sentences = fullDescription.match(/[^.!?]+[.!?]+/g) || [];
    let short = '';

    for (const sentence of sentences) {
      if ((short + sentence).length <= 250) {
        short += sentence;
      } else {
        break;
      }
    }

    // Если получилось слишком коротко, берем первые 250 символов
    if (short.length < 100) {
      short = fullDescription.substring(0, 247) + '...';
    }

    return short.trim();

  } catch (error) {
    logWarning('⚠️ Ошибка генерации краткого описания');
    return fullDescription.substring(0, 250);
  }
}

/**
 * ОЦЕНКА КАЧЕСТВА ОПИСАНИЯ
 */
function assessDescriptionQuality(description) {
  const errors = [];
  const warnings = [];
  let score = 100;

  // Проверка длины
  if (description.length < 500) {
    errors.push('Описание слишком короткое (< 500 символов)');
    score -= 30;
  } else if (description.length < 800) {
    warnings.push('Описание коротковато (< 800 символов)');
    score -= 10;
  }

  if (description.length > 2000) {
    warnings.push('Описание слишком длинное (> 2000 символов)');
    score -= 10;
  }

  // Проверка на штампы
  const stamps = [
    'инновационный', 'передовой', 'революционный', 'уникальный в своем роде',
    'не имеет аналогов', 'лидер рынка', 'беспрецедентный'
  ];

  for (const stamp of stamps) {
    if (description.toLowerCase().includes(stamp)) {
      warnings.push(`Найден штамп: "${stamp}"`);
      score -= 5;
    }
  }

  // Проверка на markdown
  if (description.includes('**') || description.includes('##') || description.includes('```')) {
    warnings.push('Обнаружена markdown-разметка');
    score -= 15;
  }

  // Проверка на тошноту (частота слов)
  const wordFreq = calculateWordFrequency(description);
  for (const [word, freq] of Object.entries(wordFreq)) {
    if (word.length > 4 && freq > 5) {
      warnings.push(`Частое повторение слова "${word}" (${freq} раз)`);
      score -= 3;
    }
  }

  return {
    score: Math.max(0, score),
    errors: errors,
    warnings: warnings,
    length: description.length
  };
}

/**
 * РАСЧЕТ ЧАСТОТЫ СЛОВ
 */
function calculateWordFrequency(text) {
  const words = text
    .toLowerCase()
    .replace(/[.,!?;:()]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 3);

  const freq = {};
  for (const word of words) {
    freq[word] = (freq[word] || 0) + 1;
  }

  return freq;
}

// =============================================================================
// РАБОТА С OPENAI ASSISTANTS API
// =============================================================================

/**
 * ВЫЗОВ OPENAI ASSISTANT
 *
 * @param {string} userContent - Промпт для ассистента
 * @param {string} assistantId - ID ассистента
 * @param {string} apiKey - OpenAI API ключ
 * @returns {string} Ответ от ассистента
 */
function callOpenAIAssistant(userContent, assistantId, apiKey) {
  const timestamp = new Date().toISOString();
  const assistantName = ASSISTANT_NAMES[assistantId] || assistantId;

  logInfo(`🤖 Запрос к ${assistantName}`);

  // Проверка параметров
  if (!userContent || !assistantId || !apiKey) {
    throw new Error('Некорректные параметры для OpenAI');
  }

  // Повторы при ошибках
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      logInfo(`📡 OpenAI запрос (попытка ${attempt}/${maxRetries})`);

      const rawResult = executeOpenAIRequest(userContent, assistantId, apiKey);
      const cleanedResult = cleanMarkdownFromResult(rawResult);

      logInfo(`✅ OpenAI ответ получен: ${cleanedResult.length} символов`);

      return cleanedResult;

    } catch (error) {
      logError(`❌ OpenAI ошибка на попытке ${attempt}`, error);

      if (attempt === maxRetries) {
        throw new Error(`OpenAI failed after ${maxRetries} attempts: ${error.message}`);
      }

      // Пауза перед повтором
      const delay = attempt * 3000;
      logInfo(`⏳ Пауза ${delay / 1000} секунд перед попыткой ${attempt + 1}...`);
      Utilities.sleep(delay);
    }
  }
}

/**
 * ВЫПОЛНЕНИЕ ЗАПРОСА К OPENAI ASSISTANTS API
 */
function executeOpenAIRequest(userContent, assistantId, apiKey) {
  let threadId;

  // 1. Создание треда
  try {
    const threadPayload = {
      messages: [{
        role: 'user',
        content: userContent
      }]
    };

    const threadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      payload: JSON.stringify(threadPayload),
      muteHttpExceptions: true
    });

    const threadStatusCode = threadResponse.getResponseCode();
    if (threadStatusCode >= 400) {
      throw new Error(`Ошибка создания треда: ${threadStatusCode} - ${threadResponse.getContentText()}`);
    }

    const threadData = JSON.parse(threadResponse.getContentText());
    threadId = threadData.id;

  } catch (error) {
    throw new Error(`Ошибка создания thread: ${error.message}`);
  }

  // 2. Запуск выполнения
  let runId;
  try {
    const runPayload = { assistant_id: assistantId };

    const runResponse = UrlFetchApp.fetch(`https://api.openai.com/v1/threads/${threadId}/runs`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      payload: JSON.stringify(runPayload),
      muteHttpExceptions: true
    });

    const runStatusCode = runResponse.getResponseCode();
    if (runStatusCode >= 400) {
      throw new Error(`Ошибка запуска run: ${runStatusCode} - ${runResponse.getContentText()}`);
    }

    const runData = JSON.parse(runResponse.getContentText());
    runId = runData.id;

  } catch (error) {
    throw new Error(`Ошибка запуска run: ${error.message}`);
  }

  // 3. Ожидание завершения
  let status = 'queued';
  let attempts = 0;
  const maxStatusChecks = 20;

  while ((status !== 'completed' && status !== 'failed' && status !== 'cancelled') && attempts < maxStatusChecks) {
    Utilities.sleep(3000); // 3 секунды
    attempts++;

    try {
      const statusResponse = UrlFetchApp.fetch(`https://api.openai.com/v1/threads/${threadId}/runs/${runId}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'OpenAI-Beta': 'assistants=v2'
        },
        muteHttpExceptions: true
      });

      if (statusResponse.getResponseCode() >= 400) {
        continue;
      }

      const statusData = JSON.parse(statusResponse.getContentText());
      status = statusData.status;

    } catch (error) {
      continue;
    }
  }

  if (status === 'failed') {
    throw new Error('Ассистент завершился с ошибкой');
  }

  if (status === 'cancelled') {
    throw new Error('Выполнение было отменено');
  }

  if (status !== 'completed') {
    throw new Error(`Ассистент не завершил выполнение за ${maxStatusChecks * 3} секунд`);
  }

  // 4. Получение результата
  try {
    const messagesResponse = UrlFetchApp.fetch(`https://api.openai.com/v1/threads/${threadId}/messages`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'OpenAI-Beta': 'assistants=v2'
      },
      muteHttpExceptions: true
    });

    if (messagesResponse.getResponseCode() >= 400) {
      throw new Error(`Ошибка получения сообщений: ${messagesResponse.getContentText()}`);
    }

    const messagesData = JSON.parse(messagesResponse.getContentText());
    const messages = messagesData.data;

    for (let message of messages) {
      if (message.role === 'assistant' &&
        message.content &&
        Array.isArray(message.content) &&
        message.content[0]?.type === 'text' &&
        message.content[0]?.text?.value) {

        return message.content[0].text.value;
      }
    }

    throw new Error('Не найдено сообщение от ассистента');

  } catch (error) {
    throw new Error(`Ошибка при получении сообщений: ${error.message}`);
  }
}

/**
 * ОЧИСТКА MARKDOWN ИЗ РЕЗУЛЬТАТА
 */
function cleanMarkdownFromResult(text) {
  if (!text || typeof text !== 'string') {
    return text;
  }

  let cleaned = text.trim();

  // Удаляем markdown блоки
  cleaned = cleaned.replace(/^```html\s*/i, '');
  cleaned = cleaned.replace(/^```json\s*/i, '');
  cleaned = cleaned.replace(/^```\s*/i, '');
  cleaned = cleaned.replace(/\s*```\s*$/i, '');

  // Удаляем префиксы ассистентов
  cleaned = cleaned.replace(/^json\s*/i, '');
  cleaned = cleaned.replace(/^response:\s*/i, '');
  cleaned = cleaned.replace(/^result:\s*/i, '');

  return cleaned.trim();
}

// =============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * ПОЛУЧЕНИЕ OPENAI API КЛЮЧА
 */
function getOpenAIKey() {
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty(SCRIPT_PROPERTIES_KEYS.OPENAI_API_KEY);

  if (!apiKey) {
    throw new Error('OpenAI API ключ не настроен в Script Properties');
  }

  return apiKey;
}

/**
 * ПАКЕТНАЯ ГЕНЕРАЦИЯ ОПИСАНИЙ
 *
 * Обрабатывает отмеченные товары и генерирует для них описания
 */
function batchGenerateDescriptions() {
  try {
    logInfo('🚀 Запуск пакетной генерации описаний');

    const products = readSelectedProducts();

    if (products.length === 0) {
      logWarning('⚠️ Нет отмеченных товаров для обработки');
      return;
    }

    logInfo(`📦 Обрабатываем ${products.length} товаров`);

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < products.length; i++) {
      const product = products[i];

      try {
        logInfo(`[${i + 1}/${products.length}] Обработка ${product.article}`);

        // Проверяем наличие исходных данных
        if (!product[IMAGES_COLUMNS.DESCRIPTION - 1]) {
          logWarning(`⚠️ У товара ${product.article} нет описания от поставщика, пропускаем`);
          continue;
        }

        // Формируем данные для генерации
        const productData = {
          article: product[IMAGES_COLUMNS.ARTICLE - 1],
          productName: product[IMAGES_COLUMNS.PRODUCT_NAME - 1],
          description: product[IMAGES_COLUMNS.DESCRIPTION - 1],
          specifications: parseSpecifications(product[IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED - 1]),
          brand: product[IMAGES_COLUMNS.BRAND - 1],
          categories: product[IMAGES_COLUMNS.CATEGORIES - 1]
        };

        // Генерируем описание
        const result = generateProductDescription(productData);

        if (result.rewrittenDescription) {
          // Записываем результат в таблицу
          updateProductField(
            productData.article,
            IMAGES_COLUMNS.DESCRIPTION_REWRITTEN,
            result.rewrittenDescription
          );

          updateProductField(
            productData.article,
            IMAGES_COLUMNS.SHORT_DESCRIPTION,
            result.shortDescription
          );

          logInfo(`✅ [${i + 1}/${products.length}] ${product.article}: описание сгенерировано (${result.quality.score}%)`);
          successCount++;

        } else {
          logError(`❌ [${i + 1}/${products.length}] ${product.article}: не удалось сгенерировать описание`);
          errorCount++;
        }

        // Пауза между товарами
        if (i < products.length - 1) {
          Utilities.sleep(2000);
        }

      } catch (error) {
        handleError(error, `Обработка товара ${product.article}`);
        errorCount++;
      }
    }

    logInfo(`✅ Пакетная генерация завершена: успешно ${successCount}, ошибок ${errorCount}`);

  } catch (error) {
    handleError(error, 'Пакетная генерация описаний');
  }
}

/**
 * ПАРСИНГ ХАРАКТЕРИСТИК ИЗ JSON
 */
function parseSpecifications(specsJson) {
  try {
    if (!specsJson) return {};

    if (typeof specsJson === 'string') {
      return JSON.parse(specsJson);
    }

    return specsJson;

  } catch (error) {
    logWarning('⚠️ Ошибка парсинга характеристик');
    return {};
  }
}

/**
 * ТЕСТ ГЕНЕРАЦИИ ОПИСАНИЯ
 */
function testDescriptionGeneration() {
  logInfo('🧪 Тестируем генерацию описаний');

  const testProduct = {
    article: 'TEST-001',
    productName: 'Бинокль Veber 10x42',
    description: 'Высококачественный бинокль с увеличением 10x и диаметром объектива 42 мм. Призмы ROOF, стекло BaK-4.',
    specifications: {
      'Параметр: Кратность увеличения, крат': '10',
      'Параметр: Диаметр объектива, мм': '42',
      'Параметр: Призменная схема': 'ROOF',
      'Параметр: Марка стекла': 'BaK-4'
    },
    brand: 'Veber',
    categories: 'Бинокли'
  };

  const result = generateProductDescription(testProduct);

  logInfo('=== РЕЗУЛЬТАТ ТЕСТА ===');
  logInfo('Полное описание:', result.rewrittenDescription);
  logInfo('Краткое описание:', result.shortDescription);
  logInfo('Качество:', JSON.stringify(result.quality, null, 2));
}
