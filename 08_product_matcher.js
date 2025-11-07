/**
 * ========================================
 * МОДУЛЬ 08: СОПОСТАВЛЕНИЕ И ПРОВЕРКА ДУБЛИКАТОВ
 * ========================================
 *
 * Проверяет наличие товаров в InSales по артикулу и названию
 * Предотвращает создание дубликатов при импорте
 */

// =============================================================================
// КОНСТАНТЫ
// =============================================================================

const MATCH_STATUS = {
  NOT_CHECKED: 'Не проверен',
  NO_MATCH: 'Совпадений нет',
  EXACT_MATCH: 'Точное совпадение',
  PARTIAL_MATCH: 'Частичное совпадение',
  DUPLICATE: 'Дубликат'
};

const MATCH_THRESHOLDS = {
  EXACT: 100,        // 100% совпадение - точный дубликат
  HIGH: 85,          // 85%+ - очень похожий товар
  MEDIUM: 70,        // 70%+ - возможно дубликат
  LOW: 50            // 50%+ - слабое сходство
};

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * ПРОВЕРКА ТОВАРА НА ДУБЛИКАТЫ
 *
 * @param {string} article - Артикул товара
 * @param {string} productName - Название товара
 * @returns {Object} Результат сопоставления
 */
function checkProductDuplicate(article, productName) {
  try {
    logInfo(`🔍 Проверяем дубликаты для ${article}`);

    const result = {
      article: article,
      matchStatus: MATCH_STATUS.NOT_CHECKED,
      confidence: 0,
      matches: [],
      recommendation: ''
    };

    // 1. Проверка по точному артикулу
    const exactMatch = findProductByArticle(article);

    if (exactMatch) {
      result.matchStatus = MATCH_STATUS.EXACT_MATCH;
      result.confidence = 100;
      result.matches.push({
        type: 'exact_article',
        product: exactMatch,
        similarity: 100
      });
      result.recommendation = 'ДУБЛИКАТ! Товар с таким артикулом уже существует';

      logWarning(`⚠️ Найден точный дубликат по артикулу: ${exactMatch.title}`);
      return result;
    }

    // 2. Нечеткий поиск по названию
    if (productName) {
      const fuzzyMatches = findSimilarProductsByName(productName);

      if (fuzzyMatches.length > 0) {
        const topMatch = fuzzyMatches[0];

        result.matches = fuzzyMatches;

        if (topMatch.similarity >= MATCH_THRESHOLDS.EXACT) {
          result.matchStatus = MATCH_STATUS.DUPLICATE;
          result.confidence = topMatch.similarity;
          result.recommendation = 'Вероятный ДУБЛИКАТ! Проверьте вручную';

        } else if (topMatch.similarity >= MATCH_THRESHOLDS.HIGH) {
          result.matchStatus = MATCH_STATUS.PARTIAL_MATCH;
          result.confidence = topMatch.similarity;
          result.recommendation = 'Похожий товар найден. Рекомендуется проверка';

        } else if (topMatch.similarity >= MATCH_THRESHOLDS.MEDIUM) {
          result.matchStatus = MATCH_STATUS.PARTIAL_MATCH;
          result.confidence = topMatch.similarity;
          result.recommendation = 'Возможно похожий товар. Требует проверки';

        } else {
          result.matchStatus = MATCH_STATUS.NO_MATCH;
          result.confidence = topMatch.similarity;
          result.recommendation = 'Совпадений не найдено, можно создавать';
        }

        logInfo(`🔎 Найдено ${fuzzyMatches.length} похожих товаров, макс. сходство: ${topMatch.similarity}%`);
      } else {
        result.matchStatus = MATCH_STATUS.NO_MATCH;
        result.confidence = 0;
        result.recommendation = 'Совпадений не найдено, можно создавать';
        logInfo('✅ Дубликатов не найдено');
      }
    }

    return result;

  } catch (error) {
    handleError(error, 'Проверка дубликатов');
    return {
      article: article,
      matchStatus: MATCH_STATUS.NOT_CHECKED,
      confidence: 0,
      matches: [],
      recommendation: `Ошибка проверки: ${error.message}`
    };
  }
}

/**
 * ПОИСК ТОВАРА ПО АРТИКУЛУ
 */
function findProductByArticle(article) {
  try {
    // Ищем по всем вариантам товара
    const response = makeInsalesRequest(
      'GET',
      '/admin/products.json',
      null,
      {
        per_page: 250,
        updated_since: '2020-01-01' // Ищем только среди относительно свежих
      }
    );

    if (!response || !response.products) {
      return null;
    }

    // Проверяем все товары
    for (const product of response.products) {
      // Проверяем основной SKU
      if (product.sku && product.sku.toLowerCase() === article.toLowerCase()) {
        return product;
      }

      // Проверяем варианты
      if (product.variants && product.variants.length > 0) {
        for (const variant of product.variants) {
          if (variant.sku && variant.sku.toLowerCase() === article.toLowerCase()) {
            return product;
          }
        }
      }
    }

    return null;

  } catch (error) {
    logError('Ошибка поиска по артикулу', error);
    return null;
  }
}

/**
 * НЕЧЕТКИЙ ПОИСК ПО НАЗВАНИЮ
 *
 * @param {string} productName - Название для поиска
 * @param {number} limit - Максимум результатов (по умолчанию 5)
 * @returns {Array} Массив похожих товаров с оценкой сходства
 */
function findSimilarProductsByName(productName, limit = 5) {
  try {
    const normalizedQuery = normalizeProductName(productName);

    // Получаем все товары из InSales (или используем кеш)
    const allProducts = loadAllInSalesProducts();

    if (!allProducts || allProducts.length === 0) {
      return [];
    }

    // Рассчитываем сходство для каждого товара
    const matches = [];

    for (const product of allProducts) {
      const normalizedTitle = normalizeProductName(product.title);
      const similarity = calculateStringSimilarity(normalizedQuery, normalizedTitle);

      if (similarity >= MATCH_THRESHOLDS.LOW) {
        matches.push({
          product: product,
          similarity: similarity,
          title: product.title,
          id: product.id
        });
      }
    }

    // Сортируем по убыванию сходства
    matches.sort((a, b) => b.similarity - a.similarity);

    // Возвращаем топ-N результатов
    return matches.slice(0, limit);

  } catch (error) {
    logError('Ошибка нечеткого поиска', error);
    return [];
  }
}

/**
 * НОРМАЛИЗАЦИЯ НАЗВАНИЯ ТОВАРА
 *
 * Приводит название к единому виду для сравнения
 */
function normalizeProductName(name) {
  if (!name) return '';

  let normalized = name.toLowerCase();

  // Удаляем артикулы в скобках
  normalized = normalized.replace(/\s*\([^)]*\)\s*/g, ' ');

  // Удаляем спецсимволы
  normalized = normalized.replace(/[^а-яёa-z0-9\s]/gi, ' ');

  // Убираем множественные пробелы
  normalized = normalized.replace(/\s+/g, ' ').trim();

  return normalized;
}

/**
 * РАСЧЕТ СХОДСТВА СТРОК (алгоритм Левенштейна + токенизация)
 *
 * @returns {number} Процент сходства (0-100)
 */
function calculateStringSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  if (str1 === str2) return 100;

  // Токенизация (разбиение на слова)
  const tokens1 = new Set(str1.split(/\s+/));
  const tokens2 = new Set(str2.split(/\s+/));

  // Пересечение токенов
  const intersection = new Set([...tokens1].filter(x => tokens2.has(x)));
  const union = new Set([...tokens1, ...tokens2]);

  // Jaccard similarity для токенов
  const tokenSimilarity = (intersection.size / union.size) * 100;

  // Левенштейн для полных строк
  const levenshteinSimilarity = calculateLevenshteinSimilarity(str1, str2);

  // Комбинированная оценка (70% токены, 30% Левенштейн)
  const combined = (tokenSimilarity * 0.7) + (levenshteinSimilarity * 0.3);

  return Math.round(combined);
}

/**
 * РАССТОЯНИЕ ЛЕВЕНШТЕЙНА
 */
function calculateLevenshteinSimilarity(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;

  if (len1 === 0) return 0;
  if (len2 === 0) return 0;

  // Матрица расстояний
  const matrix = Array(len1 + 1).fill(null).map(() => Array(len2 + 1).fill(0));

  for (let i = 0; i <= len1; i++) matrix[i][0] = i;
  for (let j = 0; j <= len2; j++) matrix[0][j] = j;

  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,      // Удаление
        matrix[i][j - 1] + 1,      // Вставка
        matrix[i - 1][j - 1] + cost // Замена
      );
    }
  }

  const distance = matrix[len1][len2];
  const maxLen = Math.max(len1, len2);
  const similarity = ((maxLen - distance) / maxLen) * 100;

  return Math.round(similarity);
}

/**
 * ЗАГРУЗКА ВСЕХ ТОВАРОВ ИЗ INSALES
 *
 * Использует кеш для оптимизации
 */
function loadAllInSalesProducts() {
  try {
    const cacheKey = 'insales_products_cache';
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);

    if (cached) {
      logInfo('📦 Используем кешированные товары InSales');
      return JSON.parse(cached);
    }

    logInfo('🔄 Загружаем все товары из InSales...');

    const allProducts = [];
    let page = 1;
    const perPage = 250;
    let hasMore = true;

    while (hasMore && page <= 10) { // Максимум 10 страниц (2500 товаров)
      const response = makeInsalesRequest(
        'GET',
        '/admin/products.json',
        null,
        { per_page: perPage, page: page }
      );

      if (response && response.products && response.products.length > 0) {
        allProducts.push(...response.products);
        logInfo(`  Загружено страница ${page}: ${response.products.length} товаров`);

        if (response.products.length < perPage) {
          hasMore = false;
        } else {
          page++;
          Utilities.sleep(500); // Пауза между запросами
        }
      } else {
        hasMore = false;
      }
    }

    logInfo(`✅ Загружено ${allProducts.length} товаров из InSales`);

    // Кешируем на 30 минут
    cache.put(cacheKey, JSON.stringify(allProducts), 1800);

    return allProducts;

  } catch (error) {
    handleError(error, 'Загрузка товаров InSales');
    return [];
  }
}

/**
 * ОЧИСТКА КЕША ТОВАРОВ
 */
function clearProductsCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('insales_products_cache');
  logInfo('🧹 Кеш товаров очищен');
}

// =============================================================================
// ПАКЕТНАЯ ПРОВЕРКА
// =============================================================================

/**
 * ПРОВЕРКА ВСЕХ ОТМЕЧЕННЫХ ТОВАРОВ НА ДУБЛИКАТЫ
 */
function batchCheckDuplicates() {
  try {
    logInfo('🚀 Запуск пакетной проверки дубликатов');

    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    let checkedCount = 0;
    let duplicatesFound = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const checkbox = row[IMAGES_COLUMNS.CHECKBOX - 1];

      if (!checkbox) continue; // Пропускаем неотмеченные

      const article = row[IMAGES_COLUMNS.ARTICLE - 1];
      const productName = row[IMAGES_COLUMNS.PRODUCT_NAME - 1];

      if (!article) continue;

      logInfo(`[${checkedCount + 1}] Проверяем ${article}`);

      const matchResult = checkProductDuplicate(article, productName);

      // Записываем результат в таблицу
      sheet.getRange(i + 1, IMAGES_COLUMNS.MATCH_STATUS).setValue(matchResult.matchStatus);
      sheet.getRange(i + 1, IMAGES_COLUMNS.MATCH_CONFIDENCE).setValue(matchResult.confidence);

      if (matchResult.matchStatus === MATCH_STATUS.EXACT_MATCH ||
        matchResult.matchStatus === MATCH_STATUS.DUPLICATE) {
        duplicatesFound++;
        logWarning(`⚠️ Дубликат: ${article} (${matchResult.confidence}%)`);
      } else {
        logInfo(`✅ ${article}: ${matchResult.matchStatus} (${matchResult.confidence}%)`);
      }

      checkedCount++;

      // Пауза между проверками
      if (i < data.length - 1) {
        Utilities.sleep(500);
      }
    }

    SpreadsheetApp.flush();

    logInfo(`✅ Проверка завершена: ${checkedCount} товаров, найдено дубликатов: ${duplicatesFound}`);

  } catch (error) {
    handleError(error, 'Пакетная проверка дубликатов');
  }
}

/**
 * ТЕСТ ПРОВЕРКИ ДУБЛИКАТОВ
 */
function testDuplicateCheck() {
  logInfo('🧪 Тестируем проверку дубликатов');

  // Тест 1: Проверка реального товара (должен найти точное совпадение)
  const test1 = checkProductDuplicate('EXISTING-SKU', 'Бинокль Veber 10x42');
  logInfo('Тест 1 (существующий товар):', JSON.stringify(test1, null, 2));

  // Тест 2: Проверка несуществующего товара
  const test2 = checkProductDuplicate('NEW-SKU-123', 'Новый тестовый товар');
  logInfo('Тест 2 (новый товар):', JSON.stringify(test2, null, 2));

  // Тест 3: Нечеткий поиск
  const similarMatches = findSimilarProductsByName('Бинокль Veber 10x42 Black');
  logInfo('Тест 3 (похожие товары):', JSON.stringify(similarMatches, null, 2));
}
