/**
 * ========================================
 * МОДУЛЬ 03: INSALES API ИНТЕГРАЦИЯ (ИСПРАВЛЕННЫЙ)
 * ========================================
 * 
 * Назначение: Полная интеграция с платформой InSales
 * - Подключение к InSales API с аутентификацией
 * - Загрузка структуры каталога (categories из каталога)
 * - Фильтрация и загрузка товаров с правильными артикулами из вариантов
 * - Синхронизация данных с Google Sheets через data_manager
 */


// ========================================
// КОНСТАНТЫ И НАСТРОЙКИ
// ========================================

// INSALES_ENDPOINTS теперь определен в 00_config.gs.js

// Дополнительные эндпоинты, которых нет в основной конфигурации
const LEGACY_ENDPOINTS = {
  ACCOUNT: '/admin/account.json',
  CATEGORIES: '/admin/categories.json',  // КАТАЛОГ - основная структура
  CATEGORY_BY_ID: '/admin/categories/{id}.json',
  PRODUCT_VARIANTS: '/admin/products/{product_id}/variants.json' // ВАРИАНТЫ с артикулами
};


const REQUEST_PARAMS = {
  CATEGORY_ID: 'category_id',
  COLLECTION_ID: 'collection_id',
  STATUS: 'status',
  LIMIT: 'per_page',
  PAGE: 'page',
  SORT_BY: 'sort_by',
  SORT_ORDER: 'sort_order'
};


// ========================================
// ОСНОВНЫЕ ФУНКЦИИ ПОДКЛЮЧЕНИЯ
// ========================================


/**
 * Тестирование подключения к InSales API
 */
async function testInsalesConnection() {
  try {
    logInfo('🔌 Начинаем тестирование подключения к InSales API');
    
    const credentials = await getInsalesCredentials();
    if (!credentials) {
      logError('❌ Не удалось получить учетные данные InSales');
      return false;
    }
    
    const response = await makeInsalesRequest('GET', LEGACY_ENDPOINTS.ACCOUNT);
    
    if (response && response.id) {
      logInfo('✅ Подключение к InSales успешно установлено', {
        shop: response.subdomain,
        organization: response.organization,
        accountId: response.id
      });
      
      await sendNotificationSafe('🎉 InSales API подключен успешно!');
      return true;
    } else {
      logError('❌ Ответ API не содержит информацию об аккаунте');
      return false;
    }
    
  } catch (error) {
    handleError(error, 'Тестирование подключения InSales');
    return false;
  }
}


/**
 * Получение учетных данных InSales
 */
function getInsalesCredentials() {
  try {
    const apiKey = getSetting('insalesApiKey') || getSetting('InSales_API_Key');
    const password = getSetting('insalesPassword') || getSetting('InSales_Password');
    const shop = getSetting('insalesShop') || getSetting('InSales_Shop');

    if (!apiKey || !password || !shop) {
      logError('❌ Отсутствуют обязательные настройки InSales');
      return null;
    }

    const baseUrl = `https://${shop}`;

    logInfo('✅ Учетные данные InSales получены', {
      shop: shop,
      baseUrl: baseUrl,
      hasApiKey: !!apiKey,
      hasPassword: !!password
    });

    return {
      apiKey: apiKey,
      password: password,
      shop: shop,
      baseUrl: baseUrl
    };

  } catch (error) {
    handleError(error, 'Получение учетных данных InSales');
    return null;
  }
}


/**
 * Универсальная функция для выполнения запросов к InSales API
 */
function makeInsalesRequest(method, endpoint, payload = null, params = null) {
  const context = `InSales API ${method} ${endpoint}`;
  const maxRetries = 5; // Увеличено с 3 до 5 для лучшей обработки Rate Limit

  try {
    const credentials = getInsalesCredentials();
    
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }
    
    let url = `${credentials.baseUrl}${endpoint}`;
    
    // Формируем URL с параметрами для GET запросов
    if (method === 'GET' && params && typeof params === 'object') {
      const paramPairs = [];
      
      Object.keys(params).forEach(key => {
        const value = params[key];
        if (value !== null && value !== undefined) {
          const encodedKey = encodeURIComponent(key);
          const encodedValue = encodeURIComponent(String(value));
          paramPairs.push(`${encodedKey}=${encodedValue}`);
        }
      });
      
      if (paramPairs.length > 0) {
        url += '?' + paramPairs.join('&');
      }
    }
    
    logInfo('🌐 Выполняем запрос к InSales API', { url: url, method: method });
    
    const options = {
      method: method,
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    // Добавляем payload для POST/PUT запросов
    if (payload && (method === 'POST' || method === 'PUT')) {
      options.payload = JSON.stringify(payload);
    }
    
    // Выполняем запрос с повторами
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode === 200 || responseCode === 201) {
          if (responseText && responseText.trim()) {
            try {
              const jsonResponse = JSON.parse(responseText);
              return jsonResponse;
            } catch (parseError) {
              logWarning('⚠️ Не удалось распарсить JSON ответ');
              return responseText;
            }
          } else {
            return null;
          }
        } else if (responseCode === 429) {
          const delay = Math.min(1000 * Math.pow(2, attempt), 10000);
          logWarning(`⏳ Rate limit (429), ждем ${delay}мс перед повтором (попытка ${attempt}/${maxRetries})`);

          if (attempt === maxRetries) {
            logError(`❌ Все ${maxRetries} попытки исчерпаны из-за Rate Limit 429`);
            throw new Error(`Rate limit исчерпан после ${maxRetries} попыток`);
          }

          Utilities.sleep(delay);
          continue;
        } else if (responseCode === 404) {
          logWarning(`❌ Ресурс не найден (404): ${endpoint}`);
          return null;
        } else if (responseCode === 422) {
          // 422 Unprocessable Entity - обычно означает проблему валидации (дубликат SKU и т.д.)
          logWarning(`⚠️ Ошибка валидации (422): ${responseText}`);
          throw new Error(`Ошибка валидации: ${responseText}`);
        } else {
          throw new Error(`HTTP ${responseCode}: ${responseText}`);
        }

      } catch (fetchError) {
        if (attempt === maxRetries) {
          throw fetchError;
        }
        logWarning(`⚠️ Ошибка запроса (попытка ${attempt}/${maxRetries}): ${fetchError.message}`);
        Utilities.sleep(1000 * attempt);
      }
    }

    // Если мы дошли сюда - все попытки провалились
    logError(`❌ Все ${maxRetries} попытки запроса провалились`);
    throw new Error(`Не удалось выполнить запрос после ${maxRetries} попыток`);

  } catch (error) {
    logError('❌ Ошибка в makeInsalesRequest:', error, context);
    throw error;
  }
}


// ========================================
// ФУНКЦИИ РАБОТЫ С КАТАЛОГОМ (ИСПРАВЛЕНО)
// ========================================


/**
 * Загрузка полной структуры каталога из InSales
 * ИСПРАВЛЕНО: использует /admin/categories.json (каталог), а не коллекции
 */
async function loadCatalogStructure() {
  try {
    logInfo('📁 Начинаем загрузку структуры каталога InSales');
    
    // Проверяем кэш
    const cacheKey = 'insales_catalog_structure';
    const cached = getCachedData(cacheKey, 30);
    
    if (cached) {
      logInfo('💾 Структура каталога загружена из кэша');
      return cached;
    }
    
    // Загружаем категории каталога через API
    const categories = await makeInsalesRequest('GET', INSALES_ENDPOINTS.COLLECTIONS);
    
    if (!categories || !Array.isArray(categories)) {
      throw new Error('Не удалось загрузить категории каталога из InSales');
    }
    
    logInfo(`📊 Загружено ${categories.length} категорий каталога из InSales`);
    
    // Строим иерархическую структуру
    const catalogStructure = buildCategoryHierarchy(categories);
    
    // Добавляем количество товаров для каждой категории
    await addProductCountsToCategories(catalogStructure);
    
    // Сохраняем в кэш
    setCachedData(cacheKey, catalogStructure);
    
    logInfo('✅ Структура каталога успешно загружена и закэширована');
    
    return catalogStructure;
    
  } catch (error) {
    handleError(error, 'Загрузка структуры каталога');
    return [];
  }
}


/**
 * Построение иерархической структуры категорий каталога
 */
function buildCategoryHierarchy(categories) {
  try {
    logInfo('🌳 Строим иерархическую структуру категорий каталога');
    
    const categoryMap = {};
    categories.forEach(cat => {
      categoryMap[cat.id] = {
        ...cat,
        children: [],
        productCount: 0
      };
    });
    
    const rootCategories = [];
    
    categories.forEach(category => {
      if (category.parent_id === null || category.parent_id === undefined) {
        rootCategories.push(categoryMap[category.id]);
      } else {
        const parent = categoryMap[category.parent_id];
        if (parent) {
          parent.children.push(categoryMap[category.id]);
        }
      }
    });
    
    // Сортируем категории по позиции
    const sortByPosition = (a, b) => (a.position || 0) - (b.position || 0);
    rootCategories.sort(sortByPosition);
    
    function sortChildren(categories) {
      categories.forEach(cat => {
        if (cat.children.length > 0) {
          cat.children.sort(sortByPosition);
          sortChildren(cat.children);
        }
      });
    }
    
    sortChildren(rootCategories);
    
    logInfo('✅ Иерархическая структура каталога построена');
    
    return rootCategories;
    
  } catch (error) {
    handleError(error, 'Построение иерархии категорий каталога');
    return [];
  }
}


/**
 * Добавление количества товаров к каждой категории каталога
 */
async function addProductCountsToCategories(categories) {
  const context = "Подсчет товаров в категориях каталога";
  
  try {
    logInfo('🚀 Начинаем быстрый подсчет товаров в каталоге', null, context);
    
    // Загружаем все товары с пагинацией
    const allProducts = await loadAllProductsForCounting();
    
    if (!allProducts || allProducts.length === 0) {
      logWarning('⚠️ Товары не загружены, устанавливаем счетчики в 0');
      setAllCategoryCounts(categories, 0);
      return;
    }
    
    // Группируем товары по категориям каталога
    const categoryProductCounts = {};
    
    for (const product of allProducts) {
      if (product.category_id) {
        const categoryId = product.category_id;
        categoryProductCounts[categoryId] = (categoryProductCounts[categoryId] || 0) + 1;
      }
    }
    
    logInfo('📊 Статистика по товарам в каталоге:', {
      totalProducts: allProducts.length,
      categoriesWithProducts: Object.keys(categoryProductCounts).length
    });
    
    // Применяем счетчики к категориям
    applyCategoryProductCounts(categories, categoryProductCounts);
    
    logInfo('✅ Быстрый подсчет товаров в каталоге завершен', null, context);
    
  } catch (error) {
    logError('❌ Ошибка быстрого подсчета', error, context);
    // Устанавливаем нулевые счетчики при ошибке
    setAllCategoryCounts(categories, 0);
  }
}


/**
 * Загружает все товары с пагинацией для подсчета в каталоге
 */
async function loadAllProductsForCounting() {
  const context = "Загрузка всех товаров каталога";
  const allProducts = [];
  let page = 1;
  const perPage = 250;
  
  try {
    while (true) {
      logInfo(`📦 Загружаем страницу ${page} каталога...`, null, context);
      
      // Загружаем ВСЕ товары из каталога без фильтров
      const products = await makeInsalesRequest('GET', INSALES_ENDPOINTS.PRODUCTS, null, {
        page: page,
        per_page: perPage
        // БЕЗ фильтра status - загружаем все товары каталога
      });
      
      if (!products || !Array.isArray(products) || products.length === 0) {
        break;
      }
      
      allProducts.push(...products);
      
      logInfo(`📊 Страница ${page}: ${products.length} товаров, всего: ${allProducts.length}`);
      
      if (products.length < perPage || page > 50) {
        break;
      }
      
      page++;
      Utilities.sleep(300);
    }
    
    logInfo(`✅ Загружено товаров из каталога: ${allProducts.length}`, null, context);
    return allProducts;
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров каталога', error, context);
    return allProducts;
  }
}


/**
 * Применяет счетчики товаров к иерархии категорий
 */
function applyCategoryProductCounts(categories, productCounts) {
  for (const category of categories) {
    category.productCount = productCounts[category.id] || 0;
    
    if (category.children && category.children.length > 0) {
      applyCategoryProductCounts(category.children, productCounts);
    }
  }
}


/**
 * Устанавливает одинаковое значение счетчика для всех категорий
 */
function setAllCategoryCounts(categories, count) {
  for (const category of categories) {
    category.productCount = count;
    
    if (category.children && category.children.length > 0) {
      setAllCategoryCounts(category.children, count);
    }
  }
}


// ========================================
// ФУНКЦИИ ЗАГРУЗКИ ТОВАРОВ (ИСПРАВЛЕНО)
// ========================================


/**
 * Основная функция загрузки товаров из InSales каталога
 */
async function loadProductsFromInSales() {
  const context = "Загрузка товаров из InSales каталога";
  
  try {
    logInfo('🚀 Начинаем загрузку товаров из InSales каталога', null, context);
    
    // 1. Показываем диалог выбора категорий
    const selectedCategories = await showCategorySelectionSafe();
    
    if (!selectedCategories || selectedCategories.length === 0) {
      logWarning('⚠️ Категории не выбраны, загрузка отменена');
      return {
        success: false,
        message: 'Категории не выбраны'
      };
    }
    
    logInfo(`✅ Выбрано категорий каталога: ${selectedCategories.length}`);
    
    // 2. Загружаем товары из выбранных категорий каталога
    const loadedProducts = await loadProductsByCategories(selectedCategories);
    
    if (!loadedProducts || loadedProducts.length === 0) {
      logWarning('⚠️ Товары не найдены в выбранных категориях каталога');
      return {
        success: false,
        message: 'Товары не найдены'
      };
    }
    
    logInfo(`✅ Загружено товаров: ${loadedProducts.length}`);
    
    // 3. Синхронизируем с Google Sheets
    const syncResult = await syncProductData(loadedProducts);
    
    logInfo('✅ Загрузка товаров из InSales каталога завершена', {
      categoriesSelected: selectedCategories.length,
      productsLoaded: loadedProducts.length,
      syncResult: syncResult
    }, context);
    
    return {
      success: true,
      categoriesSelected: selectedCategories.length,
      productsLoaded: loadedProducts.length,
      syncResult: syncResult
    };
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров из InSales каталога', error, context);
    throw error;
  }
}


/**
 * Безопасная версия диалога выбора категорий каталога
 */
async function showCategorySelectionSafe() {
  const context = "Безопасный выбор категорий каталога";
  
  try {
    logInfo('📋 Безопасный выбор категорий каталога', null, context);
    
    const catalogStructure = await loadCatalogStructure();
    
    if (!catalogStructure || catalogStructure.length === 0) {
      throw new Error('Не удалось загрузить структуру каталога');
    }
    
    // Собираем все категории каталога с товарами
    const categoriesWithProducts = [];
    
    function collectCategoriesWithProducts(categories) {
      for (const category of categories) {
        if (category.productCount && category.productCount > 0) {
          categoriesWithProducts.push({
            id: category.id,
            title: category.title,
            productCount: category.productCount
          });
        }
        
        if (category.children && category.children.length > 0) {
          collectCategoriesWithProducts(category.children);
        }
      }
    }
    
    collectCategoriesWithProducts(catalogStructure);
    
    logInfo(`📊 Найдено категорий каталога с товарами: ${categoriesWithProducts.length}`);
    
    // Возвращаем первые 5 категорий для тестирования
    const selectedCategories = categoriesWithProducts.slice(0, 5);
    
    logInfo(`✅ Автоматически выбрано ${selectedCategories.length} категорий каталога для загрузки`);
    
    return selectedCategories;
    
  } catch (error) {
    logError('❌ Ошибка безопасного выбора категорий каталога', error, context);
    throw error;
  }
}


/**
 * Загружает товары из выбранных категорий каталога
 */
async function loadProductsByCategories(selectedCategories) {
  const context = "Загрузка товаров по категориям каталога";
  const allProducts = [];
  const processedSKUs = new Set();
  
  try {
    logInfo(`🔄 Загружаем товары из ${selectedCategories.length} категорий каталога`, null, context);
    
    for (const category of selectedCategories) {
      try {
        logInfo(`📁 Обрабатываем категорию каталога: "${category.title}" (${category.productCount} товаров)`);
        
        const categoryProducts = await loadProductsFromCategory(category.id);
        
        if (categoryProducts && categoryProducts.length > 0) {
          const uniqueProducts = categoryProducts.filter(product => {
            const sku = product.sku || product.id;
            if (processedSKUs.has(sku)) {
              return false;
            }
            processedSKUs.add(sku);
            return true;
          });
          
          allProducts.push(...uniqueProducts);
          
          logInfo(`✅ Загружено ${uniqueProducts.length} уникальных товаров из категории каталога "${category.title}"`);
        }
        
        Utilities.sleep(300);
        
      } catch (categoryError) {
        logWarning(`⚠️ Ошибка загрузки товаров из категории каталога "${category.title}": ${categoryError.message}`);
        continue;
      }
    }
    
    logInfo(`✅ Общий результат: загружено ${allProducts.length} уникальных товаров из каталога`, null, context);
    
    return allProducts;
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров по категориям каталога', error, context);
    throw error;
  }
}


/**
 * Загружает товары из одной категории каталога с правильными артикулами
 * ИСПРАВЛЕНО: получает артикулы из вариантов товаров (variant.sku)
 */
async function loadProductsFromCategory(categoryId) {
  const context = `Загрузка товаров категории каталога ${categoryId}`;
  const allProducts = [];
  let page = 1;
  const perPage = 100;
  
  try {
    while (true) {
      logInfo(`📦 Загружаем страницу ${page} категории каталога ${categoryId}...`);
      
      // Загружаем товары из каталога с фильтром по категории
      const products = await makeInsalesRequest('GET', INSALES_ENDPOINTS.PRODUCTS, null, {
        collection_id: categoryId,
        page: page,
        per_page: perPage
      });
      
      if (!products || !Array.isArray(products) || products.length === 0) {
        break;
      }
      
      // ИСПРАВЛЕНО: Обрабатываем товары и получаем артикулы из вариантов
      for (const product of products) {
        try {
          // Загружаем варианты товара для получения правильных артикулов
          const variants = await loadProductVariants(product.id);
          
          if (variants && variants.length > 0) {
            // Добавляем каждый вариант как отдельный товар с правильным артикулом
            for (const variant of variants) {
              const variantProduct = {
                ...product,
                id: `${product.id}_${variant.id}`,
                sku: variant.sku, // ПРАВИЛЬНЫЙ артикул из варианта
                variant_id: variant.id,
                variant_sku: variant.sku, // Дублируем для надежности
                price: variant.price || product.price,
                title: `${product.title}`.trim(),
                original_product_id: product.id
              };
              
              await enrichProductWithImages(variantProduct);
              allProducts.push(variantProduct);
            }
          } else {
            // Если вариантов нет, добавляем основной товар
            logWarning(`⚠️ У товара ${product.id} нет вариантов, добавляем как есть`);
            product.sku = `PRODUCT_${product.id}`; // Фолбэк артикул
            await enrichProductWithImages(product);
            allProducts.push(product);
          }
        } catch (variantError) {
          logWarning(`⚠️ Не удалось загрузить варианты товара ${product.id}: ${variantError.message}`);
          // Добавляем основной товар с фолбэк артикулом
          product.sku = `PRODUCT_${product.id}`;
          await enrichProductWithImages(product);
          allProducts.push(product);
        }
      }
      
      logInfo(`📊 Страница ${page}: загружено ${products.length} товаров, всего: ${allProducts.length}`);
      
      if (products.length < perPage || page > 50) {
        break;
      }
      
      page++;
      Utilities.sleep(200);
    }
    
    logInfo(`✅ Из категории каталога ${categoryId} загружено ${allProducts.length} товаров с артикулами`, null, context);
    
    return allProducts;
    
  } catch (error) {
    logError(`❌ Ошибка загрузки товаров категории каталога ${categoryId}`, error, context);
    return [];
  }
}


/**
 * ИСПРАВЛЕНО: Загружает варианты товара для получения правильных артикулов
 */
function loadProductVariants(productId) {
  try {
    const variantsEndpoint = LEGACY_ENDPOINTS.PRODUCT_VARIANTS.replace('{product_id}', productId);
    const variants = makeInsalesRequest('GET', variantsEndpoint);

    if (variants && Array.isArray(variants)) {
      logDebug(`📦 Загружено ${variants.length} вариантов для товара ${productId}`);
      return variants;
    }

    logDebug(`⚠️ Варианты не найдены для товара ${productId}`);
    return [];
  } catch (error) {
    logWarning(`⚠️ Не удалось загрузить варианты для товара ${productId}: ${error.message}`);
    return [];
  }
}


/**
 * Обогащает товар изображениями
 */
async function enrichProductWithImages(product) {
  try {
    // Если у товара уже есть изображения, форматируем их
    if (product.images && Array.isArray(product.images) && product.images.length > 0) {
      product.imageUrls = product.images.map(img => 
        img.original_url || img.medium_url || img.small_url
      ).filter(Boolean);
      return product;
    }
    
    // Если нет изображений, пробуем загрузить отдельно
    const productId = product.original_product_id || product.id;
    const imagesEndpoint = INSALES_ENDPOINTS.PRODUCT_IMAGES.replace('{product_id}', productId);
    const productImages = await makeInsalesRequest('GET', imagesEndpoint);
    
    if (productImages && Array.isArray(productImages) && productImages.length > 0) {
      product.imageUrls = productImages.map(img => 
        img.original_url || img.medium_url || img.small_url
      ).filter(Boolean);
    } else {
      product.imageUrls = [];
    }
    
    return product;
    
  } catch (error) {
    logWarning(`⚠️ Не удалось загрузить изображения для товара ${product.id}`);
    product.imageUrls = [];
    return product;
  }
}


// ========================================
// ФУНКЦИИ СИНХРОНИЗАЦИИ С GOOGLE SHEETS
// ========================================


/**
 * Синхронизация данных товаров с Google Sheets
 */
async function syncProductData(products) {
  const context = "Синхронизация товаров";
  
  try {
    logInfo(`🔄 Начинаем синхронизацию ${products.length} товаров с правильными артикулами`, null, context);
    
    const results = {
      added: 0,
      updated: 0,
      errors: 0,
      processed: []
    };
    
    for (let i = 0; i < products.length; i++) {
      const product = products[i];
      
      try {
        if (i % 10 === 0) {
          logInfo(`📦 Обрабатываем товар ${i + 1}/${products.length}`);
        }
        
        // Конвертируем в правильный формат с правильными артикулами
        const sheetProduct = convertInsalesProductToSheetFormat(product);
        
        // Проверяем существование
        const existingProduct = findProductByArticleDirect(sheetProduct.article);
        
        if (existingProduct) {
          // Обновляем существующий
          updateProductDirect(sheetProduct.article, {
            title: sheetProduct.title,
            images: sheetProduct.images
          });
          results.updated++;
        } else {
          // Добавляем новый
          writeProductDirectly(sheetProduct);
          results.added++;
        }
        
        results.processed.push({
          article: sheetProduct.article,
          title: sheetProduct.title,
          action: existingProduct ? 'updated' : 'added'
        });
        
      } catch (productError) {
        logWarning(`⚠️ Ошибка обработки товара "${product.title}": ${productError.message}`);
        results.errors++;
      }
    }
    
    logInfo('✅ Синхронизация завершена', results, context);
    
    // Отправляем уведомление
    const message = `✅ Синхронизация товаров завершена!\n\n` +
                   `📊 Результаты:\n` +
                   `• Обработано: ${products.length}\n` +
                   `• Добавлено: ${results.added}\n` +
                   `• Обновлено: ${results.updated}\n` +
                   `• Ошибок: ${results.errors}`;
    
    try {
      if (typeof sendNotificationSafe === 'function') {
        await sendNotificationSafe(message);
      } else if (typeof sendNotification === 'function') {
        sendNotification(message);
      }
    } catch (e) {
      logWarning('⚠️ Не удалось отправить уведомление');
    }
    
    return results;
    
  } catch (error) {
    logError('❌ Критическая ошибка синхронизации', error, context);
    throw error;
  }
}


/**
 * ИСПРАВЛЕНО: Конвертация товара InSales в формат Google Sheets с правильными артикулами
 */
function convertInsalesProductToSheetFormat(product) {
  try {
    // ИСПРАВЛЕНО: Правильный приоритет артикулов на основе анализа API:
    // 1. variant.sku (правильный артикул из варианта)
    // 2. product.sku (если есть у основного товара)
    // 3. Фолбэк ID с префиксом
    
    let article = '';
    
    // Проверяем различные источники артикула
    if (product.variant_sku && product.variant_sku.trim()) {
      // ОСНОВНОЙ источник - артикул из варианта товара
      article = String(product.variant_sku).trim();
    } else if (product.sku && product.sku.trim()) {
      // Альтернативный источник - артикул товара
      article = String(product.sku).trim();
    } else if (product.variant_id) {
      // Для вариантов используем комбинацию ID товара и варианта
      article = `${product.original_product_id || product.id}_VAR${product.variant_id}`;
    } else {
      // Последний вариант - ID товара
      article = `PRODUCT_${product.original_product_id || product.id}`;
    }
    
    // Собираем URL изображений
    let imageUrls = [];
    
    if (product.imageUrls && Array.isArray(product.imageUrls)) {
      imageUrls = product.imageUrls;
    } else if (product.images && Array.isArray(product.images)) {
      imageUrls = product.images.map(img => 
        img.original_url || img.medium_url || img.small_url
      ).filter(Boolean);
    }
    
    // Формируем название товара
    let title = product.title || 'Без названия';
    if (product.variant_title && product.variant_title !== product.title) {
      title += ` ${product.variant_title}`;
    }
    
    return {
      article: article,
      title: title.trim(),
      insalesId: String(product.original_product_id || product.id),
      images: imageUrls.join('\n'),
      // Дополнительная информация
      originalSku: product.sku,
      variantId: product.variant_id,
      variantSku: product.variant_sku,
      categoryId: product.category_id,
      collectionsIds: product.collections_ids ? product.collections_ids.join(',') : '',
      price: product.price
    };
    
  } catch (error) {
    logError('❌ Ошибка конвертации товара', error);
    
    return {
      article: `ERROR_${product.id || Date.now()}`,
      title: product.title || 'Ошибка загрузки',
      insalesId: String(product.id || ''),
      images: ''
    };
  }
}


// ========================================
// ФУНКЦИИ ПРЯМОЙ РАБОТЫ С GOOGLE SHEETS
// ========================================


/**
 * Прямая запись товара в Google Sheets
 */
function writeProductDirectly(productData) {
  const context = "Прямая запись товара";
  
  try {
    logInfo(`📝 Прямая запись товара: "${productData.title}"`, null, context);
    
    const sheet = getImagesSheet();
    
    if (!sheet) {
      throw new Error('Не удалось получить лист Images');
    }
    
    // ПРАВИЛЬНАЯ структура колонок согласно IMAGES_COLUMNS из config:
    // A(1) - CHECKBOX
    // B(2) - ARTICLE 
    // C(3) - INSALES_ID
    // D(4) - PRODUCT_NAME
    // E(5) - ORIGINAL_IMAGES
    // F(6) - PROCESSED_IMAGES
    // G(7) - ALT_TAGS
    // H(8) - SEO_FILENAMES
    // I(9) - PROCESSING_STATUS
    // J(10) - INSALES_STATUS
    
    const rowData = [
      false,                              // A - Чекбокс
      productData.article || '',          // B - Артикул
      productData.insalesId || '',        // C - ID InSales
      productData.title || '',            // D - Название товара
      productData.images || '',           // E - Исходные изображения
      '',                                 // F - Парсинг Поставщика (пусто)
      '',                                 // G - Дополнительные фото (пусто)
      '',                                 // H - Обработанные изображения (пусто)
      '',                                 // I - Alt-теги (пусто)
      '',                                 // J - SEO имена файлов (пусто)
      'Не обработано',                    // K - Статус обработки
      'Не отправлено'                     // L - Статус InSales
    ];
    
    // Добавляем строку в конец листа
    sheet.appendRow(rowData);
    
    // Настраиваем чекбокс для новой строки
    const lastRow = sheet.getLastRow();
    const checkboxRange = sheet.getRange(lastRow, 1); // Колонка A
    checkboxRange.insertCheckboxes();
    
    logInfo(`✅ Товар записан на строку ${lastRow}`);
    
    return {
      success: true,
      row: lastRow,
      article: productData.article
    };
    
  } catch (error) {
    logError('❌ Ошибка прямой записи товара', error, context);
    throw error;
  }
}


/**
 * Получает лист Images
 */
function getImagesSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Проверяем конфигурацию из config модуля
    const sheetName = SHEET_NAMES?.IMAGES || 'Обработка изображений';
    
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      // Пробуем альтернативные названия
      const alternativeNames = ['Images', 'images', 'Изображения', 'Обработка изображений'];
      
      for (const name of alternativeNames) {
        sheet = spreadsheet.getSheetByName(name);
        if (sheet) {
          logInfo(`✅ Найден лист: ${name}`);
          break;
        }
      }
    }
    
    if (!sheet) {
      throw new Error('Лист для обработки изображений не найден');
    }
    
    return sheet;
    
  } catch (error) {
    logError('❌ Ошибка получения листа', error);
    throw error;
  }
}


/**
 * Проверяет, существует ли товар по артикулу
 */
function findProductByArticleDirect(article) {
  const context = "Поиск товара по артикулу";
  
  try {
    const sheet = getImagesSheet();
    if (!sheet) {
      throw new Error('Не удалось получить лист');
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return null;
    }
    
    // Артикул в колонке B (индекс 1)
    const ARTICLE_COLUMN = 1; // B колонка (0-based)
    
    for (let i = 1; i < data.length; i++) {
      const rowArticle = data[i][ARTICLE_COLUMN];
      
      if (rowArticle && String(rowArticle).trim() === String(article).trim()) {
        logInfo(`✅ Товар найден на строке ${i + 1}`);
        return {
          row: i + 1,
          article: rowArticle,
          title: data[i][3],  // D колонка - название
          insalesId: data[i][2]   // C колонка - ID InSales
        };
      }
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка поиска товара', error, context);
    return null;
  }
}


/**
 * Обновляет товар по артикулу
 */
function updateProductDirect(article, updateData) {
  const context = "Прямое обновление товара";
  
  try {
    const existingProduct = findProductByArticleDirect(article);
    
    if (!existingProduct) {
      throw new Error(`Товар с артикулом ${article} не найден`);
    }
    
    const sheet = getImagesSheet();
    const row = existingProduct.row;
    
    // Обновляем конкретные колонки согласно IMAGES_COLUMNS
    if (updateData.title) {
      sheet.getRange(row, 4).setValue(updateData.title); // D - PRODUCT_NAME
    }
    
    if (updateData.images) {
      sheet.getRange(row, 5).setValue(updateData.images); // E - ORIGINAL_IMAGES
    }
    
    if (updateData.lastUpdated) {
      // Можно добавить в комментарий к статусу
      const currentStatus = sheet.getRange(row, 9).getValue();
      sheet.getRange(row, 9).setValue(currentStatus + ` (обновлено ${updateData.lastUpdated})`);
    }
    
    logInfo(`✅ Товар обновлен на строке ${row}`);
    
    return {
      success: true,
      row: row,
      article: article
    };
    
  } catch (error) {
    logError('❌ Ошибка прямого обновления товара', error, context);
    throw error;
  }
}


/**
 * СОЗДАНИЕ НОВОГО ТОВАРА В INSALES
 *
 * @param {Object} productData - Данные для создания товара
 * @returns {Object} Созданный товар с ID
 */
async function createProductInInSales(productData) {
  try {
    logInfo(`🆕 Создаем товар в InSales: ${productData.article}`);

    // Валидация обязательных полей
    if (!productData.article) {
      throw new Error('Артикул обязателен для создания товара');
    }

    if (!productData.productName) {
      throw new Error('Название товара обязательно');
    }

    // Формируем payload для InSales API
    const productPayload = {
      product: {
        title: productData.productName,
        description: productData.descriptionRewritten || productData.description || '',
        short_description: productData.shortDescription || '',

        // Основные параметры
        available: true,
        is_hidden: true, // По умолчанию скрываем до проверки

        // Цена (обязательное поле)
        price: parseFloat(productData.price) || 0,

        // Характеристики
        characteristics: buildCharacteristicsForInSales(productData.specificationsNormalized),

        // Комплектация
        package_contents: productData.packageContents || null,

        // Бренд (как custom field)
        fields_values_attributes: buildCustomFields(productData),

        // Варианты товара с артикулом, весом и габаритами
        // ВАЖНО: вес и габариты относятся к варианту, а не к товару!
        variants_attributes: [
          buildVariantAttributes(productData)
        ]
      }
    };

    // Категории (если указаны)
    if (productData.categories) {
      const categoryIds = await findOrCreateCategories(productData.categories);
      if (categoryIds && categoryIds.length > 0) {
        productPayload.product.category_id = categoryIds[0]; // Основная категория
      }
    }

    logInfo('📦 Отправляем запрос на создание товара');

    // Создаем товар
    const response = await makeInsalesRequest(
      'POST',
      INSALES_ENDPOINTS.PRODUCTS,
      productPayload
    );

    if (!response || !response.id) {
      throw new Error('InSales не вернул ID созданного товара');
    }

    logInfo(`✅ Товар создан в InSales, ID: ${response.id}`);

    // Загружаем изображения (если есть)
    if (productData.supplierImages) {
      await uploadProductImages(response.id, productData.supplierImages);
    }

    // Обновляем таблицу
    updateProductField(productData.article, IMAGES_COLUMNS.INSALES_ID, response.id);
    updateProductField(productData.article, IMAGES_COLUMNS.IMPORT_STATUS, 'Создан в InSales');
    updateProductField(productData.article, IMAGES_COLUMNS.INSALES_STATUS, STATUS_VALUES.INSALES.SENT);

    return response;

  } catch (error) {
    handleError(error, 'Создание товара в InSales');

    // Обновляем статус ошибки
    if (productData.article) {
      updateProductField(
        productData.article,
        IMAGES_COLUMNS.IMPORT_STATUS,
        `Ошибка: ${error.message}`
      );
    }

    throw error;
  }
}


/**
 * ФОРМИРОВАНИЕ ХАРАКТЕРИСТИК ДЛЯ INSALES
 */
function buildCharacteristicsForInSales(specificationsJson) {
  try {
    if (!specificationsJson) return [];

    const specs = typeof specificationsJson === 'string'
      ? JSON.parse(specificationsJson)
      : specificationsJson;

    const characteristics = [];

    for (const [key, value] of Object.entries(specs)) {
      if (value) {
        // Убираем префикс "Параметр: " из названия
        const cleanKey = key.replace(/^Параметр:\s*/i, '');

        characteristics.push({
          title: cleanKey,
          value: String(value)
        });
      }
    }

    return characteristics;

  } catch (error) {
    logWarning('⚠️ Ошибка формирования характеристик', error);
    return [];
  }
}


/**
 * ФОРМИРОВАНИЕ КАСТОМНЫХ ПОЛЕЙ (БРЕНД, СЕРИЯ)
 */
function buildCustomFields(productData) {
  const fields = [];

  if (productData.brand) {
    fields.push({
      name: 'Бренд',
      value: productData.brand
    });
  }

  if (productData.series) {
    fields.push({
      name: 'Серия',
      value: productData.series
    });
  }

  return fields;
}


/**
 * ФОРМИРОВАНИЕ АТРИБУТОВ ВАРИАНТА С ВЕСОМ И ГАБАРИТАМИ
 *
 * InSales API требует передавать вес и габариты в variants_attributes:
 * - weight: вес в КИЛОГРАММАХ
 * - width, depth, height: габариты в САНТИМЕТРАХ
 *
 * Источники данных из справочника параметров:
 * - Вес: "Параметр: Вес в упаковке, кг" (синонимы: "Вес в упаковке", "Вес с упаковкой")
 * - Габариты: "Параметр: Размер упаковки (ДхШхВ), мм" (формат: ДлинаxШиринаxВысота в мм)
 *
 * @param {Object} productData - Данные товара
 * @returns {Object} Атрибуты варианта для InSales API
 */
function buildVariantAttributes(productData) {
  const variant = {
    sku: productData.article,
    price: parseFloat(productData.price) || 0,
    quantity: parseInt(productData.stock) || 0,
    available: true
  };

  // Извлекаем вес и габариты из нормализованных спецификаций
  let specs = {};
  if (productData.specificationsNormalized) {
    try {
      specs = typeof productData.specificationsNormalized === 'string'
        ? JSON.parse(productData.specificationsNormalized)
        : productData.specificationsNormalized;
    } catch (e) {
      logWarning('⚠️ Не удалось распарсить спецификации для веса/габаритов');
    }
  }

  // Обработка веса
  // Ищем "Вес упаковки" в нормализованных спецификациях
  // ВАЖНО: ключи в JSON БЕЗ префикса "Параметр:", например "Вес упаковки": "1.66"
  const weightValue = findSpecValue(specs, [
    'Вес упаковки',
    'Вес в упаковке',
    'Вес с упаковкой',
    'Параметр: Вес в упаковке, кг',
    'Параметр: Вес упаковки'
  ]);

  if (weightValue) {
    const weightKg = parseWeight(weightValue);
    if (weightKg > 0) {
      variant.weight = weightKg;
      logInfo(`📦 Вес для доставки: ${weightKg} кг`);
    }
  }

  // Обработка габаритов
  // InSales API принимает габариты как СТРОКУ в формате "ШхГхВ" (Ширина×Глубина×Высота)
  // Документация: https://www.insales.ru/collection/doc-prochee/product/javascript-api-oformleniya-zakaza-dlya-vneshnih-sposobov-dostavki
  // Пример: "dimensions": "10х20х15" (значения в СМ)
  const dimensionsValue = findSpecValue(specs, [
    'Размер упаковки (ДхШхВ)',
    'Размер упаковки (ДхШхВ), мм',
    'Размер упаковки',
    'Габариты упаковки',
    'Параметр: Размер упаковки (ДхШхВ), мм',
    'Параметр: Размер упаковки (ДхШхВ)'
  ]);

  if (dimensionsValue) {
    const dims = parseDimensionsMm(dimensionsValue);
    if (dims) {
      // Конвертируем мм → см и формируем строку в формате InSales: "ДхШхВ"
      // В справочнике данные в формате "ДхШхВ" (длина, ширина, высота) в мм
      // InSales ожидает "ДхШхВ" (длина, ширина, высота) в см
      const lengthCm = Math.round(dims.length / 10); // Д - длина (depth в InSales)
      const widthCm = Math.round(dims.width / 10);   // Ш - ширина
      const heightCm = Math.round(dims.height / 10); // В - высота

      // Используем кириллическую "х" как в документации InSales
      variant.dimensions = `${lengthCm}х${widthCm}х${heightCm}`;
      logInfo(`📐 Габариты для доставки: ${variant.dimensions} см (ДхШхВ)`);
    }
  }

  return variant;
}


/**
 * Поиск значения в спецификациях по списку возможных ключей
 */
function findSpecValue(specs, keys) {
  if (!specs || typeof specs !== 'object') return null;

  for (const key of keys) {
    if (specs[key] !== undefined && specs[key] !== null && specs[key] !== '') {
      return specs[key];
    }
  }

  // Поиск по частичному совпадению ключа
  const specKeys = Object.keys(specs);
  for (const key of keys) {
    const found = specKeys.find(k =>
      k.toLowerCase().includes(key.toLowerCase()) ||
      key.toLowerCase().includes(k.toLowerCase())
    );
    if (found && specs[found]) {
      return specs[found];
    }
  }

  return null;
}


/**
 * Парсинг веса из строки
 * Поддерживает форматы: "2.4", "2,4", "2.4 кг", "2400 г"
 * @returns {number} Вес в килограммах
 */
function parseWeight(value) {
  if (!value) return 0;

  const str = String(value).trim().toLowerCase();

  // Извлекаем число
  const match = str.match(/([\d.,]+)/);
  if (!match) return 0;

  let weight = parseFloat(match[1].replace(',', '.'));
  if (isNaN(weight)) return 0;

  // Проверяем единицы измерения
  if (str.includes('г') && !str.includes('кг')) {
    // Граммы → конвертируем в кг
    weight = weight / 1000;
  }

  return weight;
}


/**
 * Парсинг габаритов из строки (значения в мм)
 * Поддерживает форматы: "142x51x51", "142х51х51", "142*51*51"
 * @returns {Object|null} { length, width, height } в мм
 */
function parseDimensionsMm(value) {
  if (!value) return null;

  const str = String(value).trim();

  // Паттерн: три числа через разделитель (x, х, *, ×)
  const match = str.match(/(\d+(?:[.,]\d+)?)\s*[xх×\*]\s*(\d+(?:[.,]\d+)?)\s*[xх×\*]\s*(\d+(?:[.,]\d+)?)/i);

  if (match) {
    return {
      length: parseFloat(match[1].replace(',', '.')),  // Д - первое число
      width: parseFloat(match[2].replace(',', '.')),   // Ш - второе число
      height: parseFloat(match[3].replace(',', '.'))   // В - третье число
    };
  }

  logWarning(`⚠️ Не удалось распарсить габариты: "${value}"`);
  return null;
}


/**
 * ПОИСК ИЛИ СОЗДАНИЕ КАТЕГОРИЙ
 */
async function findOrCreateCategories(categoriesString) {
  try {
    if (!categoriesString) return [];

    // Разбиваем по > или запятой
    const categoryNames = categoriesString
      .split(/[>,]/)
      .map(c => c.trim())
      .filter(c => c);

    if (categoryNames.length === 0) return [];

    // Загружаем существующие категории
    const existingCategories = await loadCatalogStructure();

    const categoryIds = [];

    for (const categoryName of categoryNames) {
      // Ищем существующую категорию
      const existing = existingCategories.find(cat =>
        cat.title.toLowerCase() === categoryName.toLowerCase()
      );

      if (existing) {
        categoryIds.push(existing.id);
      } else {
        // Создаем новую категорию
        logInfo(`🆕 Создаем новую категорию: ${categoryName}`);

        const newCategory = await makeInsalesRequest(
          'POST',
          LEGACY_ENDPOINTS.CATEGORIES,
          {
            category: {
              title: categoryName,
              position: 999
            }
          }
        );

        if (newCategory && newCategory.id) {
          categoryIds.push(newCategory.id);
        }
      }
    }

    return categoryIds;

  } catch (error) {
    logError('Ошибка работы с категориями', error);
    return [];
  }
}


/**
 * ЗАГРУЗКА ИЗОБРАЖЕНИЙ ТОВАРА
 */
async function uploadProductImages(productId, imagesString) {
  try {
    if (!imagesString) return;

    const imageUrls = imagesString.split('\n').filter(url => url.trim().startsWith('http'));

    if (imageUrls.length === 0) {
      logInfo('📷 Нет изображений для загрузки');
      return;
    }

    logInfo(`📷 Загружаем ${imageUrls.length} изображений`);

    for (let i = 0; i < imageUrls.length; i++) {
      const imageUrl = imageUrls[i].trim();

      try {
        const imagePayload = {
          image: {
            src: imageUrl,
            position: i + 1
          }
        };

        const endpoint = INSALES_ENDPOINTS.PRODUCT_IMAGES.replace('{product_id}', productId);

        await makeInsalesRequest('POST', endpoint, imagePayload);

        logInfo(`✅ Изображение ${i + 1}/${imageUrls.length} загружено`);

        // Пауза между загрузками
        if (i < imageUrls.length - 1) {
          Utilities.sleep(1000);
        }

      } catch (imageError) {
        logWarning(`⚠️ Ошибка загрузки изображения ${i + 1}: ${imageError.message}`);
      }
    }

    logInfo(`✅ Загрузка изображений завершена`);

  } catch (error) {
    logError('Ошибка загрузки изображений', error);
  }
}


/**
 * ПАКЕТНОЕ СОЗДАНИЕ ТОВАРОВ ИЗ ТАБЛИЦЫ
 */
async function batchCreateProductsInInSales() {
  try {
    logInfo('🚀 Запуск пакетного создания товаров в InSales');

    const products = readSelectedProducts();

    if (products.length === 0) {
      logWarning('⚠️ Нет отмеченных товаров для создания');
      return;
    }

    logInfo(`📦 Создаем ${products.length} товаров`);

    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;

    for (let i = 0; i < products.length; i++) {
      const row = products[i];

      try {
        const article = row[IMAGES_COLUMNS.ARTICLE - 1];

        if (!article) {
          logWarning(`⚠️ Строка ${i + 1}: нет артикула, пропускаем`);
          skippedCount++;
          continue;
        }

        logInfo(`[${i + 1}/${products.length}] Создаем товар ${article}`);

        // Проверяем статус сопоставления
        const matchStatus = row[IMAGES_COLUMNS.MATCH_STATUS - 1];

        if (matchStatus === MATCH_STATUS.EXACT_MATCH || matchStatus === MATCH_STATUS.DUPLICATE) {
          logWarning(`⚠️ ${article}: товар уже существует (${matchStatus}), пропускаем`);
          skippedCount++;
          continue;
        }

        // Формируем данные товара
        const productData = {
          article: article,
          productName: row[IMAGES_COLUMNS.PRODUCT_NAME - 1],
          description: row[IMAGES_COLUMNS.DESCRIPTION - 1],
          descriptionRewritten: row[IMAGES_COLUMNS.DESCRIPTION_REWRITTEN - 1],
          shortDescription: row[IMAGES_COLUMNS.SHORT_DESCRIPTION - 1],
          specificationsNormalized: row[IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED - 1],
          price: row[IMAGES_COLUMNS.PRICE - 1],
          stock: row[IMAGES_COLUMNS.STOCK - 1],
          categories: row[IMAGES_COLUMNS.CATEGORIES - 1],
          brand: row[IMAGES_COLUMNS.BRAND - 1],
          series: row[IMAGES_COLUMNS.SERIES - 1],
          weight: row[IMAGES_COLUMNS.WEIGHT - 1],
          dimensions: row[IMAGES_COLUMNS.DIMENSIONS - 1],
          packageContents: row[IMAGES_COLUMNS.PACKAGE_CONTENTS - 1],
          supplierImages: row[IMAGES_COLUMNS.SUPPLIER_IMAGES - 1]
        };

        // Создаем товар
        await createProductInInSales(productData);

        logInfo(`✅ [${i + 1}/${products.length}] ${article}: товар создан`);
        successCount++;

        // Пауза между товарами
        if (i < products.length - 1) {
          Utilities.sleep(2000);
        }

      } catch (error) {
        logError(`❌ [${i + 1}/${products.length}] Ошибка создания товара`, error);
        errorCount++;
      }
    }

    logInfo(`✅ Пакетное создание завершено: успешно ${successCount}, ошибок ${errorCount}, пропущено ${skippedCount}`);

  } catch (error) {
    handleError(error, 'Пакетное создание товаров');
  }
}


// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ========================================


/**
 * Безопасная отправка уведомлений
 */
async function sendNotificationSafe(message) {
  try {
    // Проверяем функцию sendNotification
    if (typeof sendNotification === 'function') {
      await sendNotification(message);
    } else if (typeof sendTelegramNotification === 'function') {
      // Проверяем частоту отправки
      const lastNotification = getSetting('last_notification_time') || '0';
      const timeSinceLastNotification = Date.now() - parseInt(lastNotification);
      
      if (timeSinceLastNotification < 5000) {
        const delay = 5000 - timeSinceLastNotification;
        Utilities.sleep(delay);
      }
      
      await sendTelegramNotification(message);
      setSetting('last_notification_time', Date.now().toString());
    } else {
      logInfo(`📢 Уведомление: ${message}`);
    }
  } catch (error) {
    logWarning('⚠️ Ошибка отправки уведомления:', error.message);
  }
}


/**
 * Получение данных из кэша
 */
function getCachedData(key, maxAgeMinutes = 60) {
  try {
    const cacheData = getSetting(key + '_cache');
    if (!cacheData) return null;
    
    const parsed = JSON.parse(cacheData);
    const ageMinutes = (Date.now() - parsed.timestamp) / (1000 * 60);
    
    if (ageMinutes <= maxAgeMinutes) {
      return parsed.data;
    } else {
      setSetting(key + '_cache', null);
      return null;
    }
  } catch (error) {
    return null;
  }
}


/**
 * Сохранение данных в кэш
 */
function setCachedData(key, data) {
  try {
    const cacheData = {
      data: data,
      timestamp: Date.now()
    };
    
    const cacheString = JSON.stringify(cacheData);
    
    // Проверяем размер перед сохранением (лимит Properties ~9KB на значение)
    if (cacheString.length > 8000) {
      logWarning(`⚠️ Объект слишком большой для кэширования (${cacheString.length} символов)`);
      return false;
    }
    
    setSetting(key + '_cache', cacheString);
    return true;
    
  } catch (error) {
    logWarning('⚠️ Не удалось сохранить данные в кэш', {
      key: key,
      error: error.message
    });
    return false;
  }
}


/**
 * Очистка кэша структуры каталога
 */
function clearCatalogCache() {
  const context = "Очистка кэша";
  
  try {
    logInfo('🗑️ Очищаем кэш структуры каталога', null, context);
    
    const cacheKeys = [
      'insales_catalog_structure_cache',
      'catalog_structure_cache',
      'InSales_Catalog_Cache',
      'catalogStructureCache'
    ];
    
    for (const key of cacheKeys) {
      try {
        PropertiesService.getScriptProperties().deleteProperty(key);
      } catch (e) {
        // Игнорируем ошибки - ключ может не существовать
      }
    }
    
    logInfo('🧹 Все возможные ключи кэша очищены');
    return true;
    
  } catch (error) {
    logError('❌ Ошибка очистки кэша', error, context);
    return false;
  }
}


// ========================================
// ТЕСТОВЫЕ ФУНКЦИИ
// ========================================


/**
 * Простой тест подключения
 */
async function testConnection() {
  try {
    console.log('🧪 Тест подключения к InSales API');
    
    const result = await testInsalesConnection();
    
    if (result) {
      console.log('✅ Подключение работает!');
      return { success: true };
    } else {
      console.log('❌ Подключение не работает');
      return { success: false };
    }
    
  } catch (error) {
    console.error('❌ Ошибка теста:', error.message);
    return { success: false, error: error.message };
  }
}


/**
 * Тест загрузки товара с правильными артикулами
 */
async function testSingleProductWithVariants() {
  const context = "Тест товара с вариантами";
  
  try {
    logInfo('🧪 Тестируем загрузку товара с вариантами и правильными артикулами', null, context);
    
    const products = await makeInsalesRequest('GET', INSALES_ENDPOINTS.PRODUCTS, null, {
      per_page: 1
    });
    
    if (!products || products.length === 0) {
      throw new Error('Товары не найдены');
    }
    
    const product = products[0];
    logInfo(`🎯 Тестируем товар: "${product.title}" (ID: ${product.id})`);
    
    // Загружаем варианты товара
    const variants = await loadProductVariants(product.id);
    logInfo(`📦 Загружено вариантов: ${variants.length}`);
    
    if (variants.length > 0) {
      const firstVariant = variants[0];
      logInfo('🔍 Первый вариант:', {
        variantId: firstVariant.id,
        variantSku: firstVariant.sku,
        price: firstVariant.price
      });
      
      // Создаем товар с правильным артикулом
      const productWithVariant = {
        ...product,
        id: `${product.id}_${firstVariant.id}`,
        sku: firstVariant.sku,
        variant_id: firstVariant.id,
        variant_sku: firstVariant.sku,
        price: firstVariant.price,
        original_product_id: product.id
      };
      
      const convertedProduct = convertInsalesProductToSheetFormat(productWithVariant);
      logInfo('📝 Результат конвертации:', convertedProduct);
      
      const syncResult = await syncProductData([productWithVariant]);
      
      logInfo('✅ Тест завершен', syncResult, context);
      
      return {
        success: true,
        product: product,
        variants: variants,
        converted: convertedProduct,
        syncResult: syncResult
      };
    } else {
      logWarning('⚠️ У товара нет вариантов');
      return {
        success: false,
        message: 'У товара нет вариантов'
      };
    }
    
  } catch (error) {
    logError('❌ Ошибка теста', error, context);
    throw error;
  }
}


/**
 * Тест структуры каталога
 */
async function testCatalogStructure() {
  try {
    logInfo('📁 Тестирование загрузки структуры каталога');
    
    const catalogStructure = await loadCatalogStructure();
    
    if (!catalogStructure || catalogStructure.length === 0) {
      throw new Error('Не удалось загрузить структуру каталога');
    }
    
    // Анализируем структуру
    let totalCategories = 0;
    let categoriesWithProducts = 0;
    let totalProducts = 0;
    
    function analyzeCategory(category) {
      totalCategories++;
      if (category.productCount && category.productCount > 0) {
        categoriesWithProducts++;
        totalProducts += category.productCount;
      }
      
      if (category.children) {
        category.children.forEach(analyzeCategory);
      }
    }
    
    catalogStructure.forEach(analyzeCategory);
    
    const stats = {
      totalCategories: totalCategories,
      categoriesWithProducts: categoriesWithProducts,
      totalProducts: totalProducts
    };
    
    logInfo('✅ Анализ структуры каталога:', stats);
    
    return {
      success: true,
      structure: catalogStructure,
      stats: stats
    };
    
  } catch (error) {
    logError('❌ Ошибка тестирования структуры каталога', error);
    throw error;
  }
}


/**
 * Полный тест загрузки товаров с правильными артикулами
 */
async function testFullWorkflowWithCorrectSKU() {
  const context = "Полный тест workflow с правильными артикулами";
  
  try {
    logInfo('🚀 Запускаем полный тест загрузки товаров с правильными артикулами', null, context);
    
    const result = await loadProductsFromInSales();
    
    if (result.success) {
      logInfo('✅ Полный тест завершен успешно', result, context);
      
      console.log(`🎉 РЕЗУЛЬТАТЫ ПОЛНОГО ТЕСТА С ПРАВИЛЬНЫМИ АРТИКУЛАМИ:
📁 Выбрано категорий каталога: ${result.categoriesSelected}
📦 Загружено товаров: ${result.productsLoaded}
➕ Добавлено: ${result.syncResult.added}
🔄 Обновлено: ${result.syncResult.updated}
❌ Ошибок: ${result.syncResult.errors}`);
    } else {
      logWarning('⚠️ Тест завершен с предупреждением', result, context);
    }
    
    return result;
    
  } catch (error) {
    logError('❌ Ошибка полного теста', error, context);
    throw error;
  }
}


/**
 * Проверка структуры листа
 */
function checkSheetStructure() {
  const context = "Проверка структуры листа";
  
  try {
    logInfo('🔍 Проверяем структуру листа Images', null, context);
    
    const sheet = getImagesSheet();
    
    if (!sheet) {
      throw new Error('Лист не найден');
    }
    
    const info = {
      name: sheet.getName(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      maxRows: sheet.getMaxRows(),
      maxColumns: sheet.getMaxColumns()
    };
    
    logInfo(`📋 Информация о листе:`, info);
    
    if (sheet.getLastRow() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      logInfo('📝 Заголовки листа:', headers);
    }
    
    return {
      success: true,
      sheetInfo: info
    };
    
  } catch (error) {
    logError('❌ Ошибка проверки структуры листа', error, context);
    throw error;
  }
}

/**
 * ========================================
 * ИСПРАВЛЕНИЯ ПРОИЗВОДИТЕЛЬНОСТИ
 * ========================================
 * 
 * Добавьте эти функции в конец файла 03_insales_api.gs
 * Заменяют проблемные функции для решения таймаутов и кэширования
 */


/**
 * ИСПРАВЛЕННАЯ функция кэширования - разбивает большие объекты
 */
function setCachedDataFixed(key, data) {
  try {
    // Для больших структур каталога не кэшируем - слишком много данных
    if (key.includes('catalog_structure')) {
      logInfo('ℹ️ Структура каталога слишком большая для кэширования, пропускаем');
      return false;
    }
    
    const cacheData = {
      data: data,
      timestamp: Date.now()
    };
    
    const cacheString = JSON.stringify(cacheData);
    
    // Лимит Properties API - примерно 9KB на значение
    if (cacheString.length > 8000) {
      logWarning(`⚠️ Объект слишком большой для кэширования (${cacheString.length} символов), пропускаем`);
      return false;
    }
    
    setSetting(key + '_cache', cacheString);
    logInfo(`✅ Данные закэшированы`, { key: key, size: cacheString.length });
    return true;
    
  } catch (error) {
    logWarning('⚠️ Не удалось сохранить данные в кэш', {
      key: key,
      error: error.message
    });
    return false;
  }
}


/**
 * ИСПРАВЛЕННАЯ функция загрузки структуры каталога - без кэширования
 */
async function loadCatalogStructureFast() {
  try {
    logInfo('📁 Начинаем быструю загрузку структуры каталога InSales');
    
    // НЕ используем кэш - загружаем напрямую
    logInfo('🚀 Загружаем категории каталога напрямую (без кэша)');
    
    // Загружаем только основные категории без подсчета товаров
    const categories = await makeInsalesRequest('GET', INSALES_ENDPOINTS.COLLECTIONS, null, {
      per_page: 100 // Ограничиваем для скорости
    });
    
    if (!categories || !Array.isArray(categories)) {
      throw new Error('Не удалось загрузить категории каталога из InSales');
    }
    
    logInfo(`📊 Загружено ${categories.length} категорий каталога из InSales`);
    
    // Строим простую структуру БЕЗ подсчета товаров
    const catalogStructure = buildSimpleCategoryHierarchy(categories);
    
    logInfo('✅ Быстрая структура каталога загружена');
    
    return catalogStructure;
    
  } catch (error) {
    handleError(error, 'Быстрая загрузка структуры каталога');
    return [];
  }
}


/**
 * Простое построение иерархии БЕЗ подсчета товаров
 */
function buildSimpleCategoryHierarchy(categories) {
  try {
    logInfo('🌳 Строим простую иерархию категорий каталога');
    
    const categoryMap = {};
    categories.forEach(cat => {
      categoryMap[cat.id] = {
        id: cat.id,
        title: cat.title,
        parent_id: cat.parent_id,
        position: cat.position,
        children: [],
        productCount: 0 // Устанавливаем в 0 для скорости
      };
    });
    
    const rootCategories = [];
    
    categories.forEach(category => {
      if (category.parent_id === null || category.parent_id === undefined) {
        rootCategories.push(categoryMap[category.id]);
      } else {
        const parent = categoryMap[category.parent_id];
        if (parent) {
          parent.children.push(categoryMap[category.id]);
        }
      }
    });
    
    // Сортируем по позиции
    const sortByPosition = (a, b) => (a.position || 0) - (b.position || 0);
    rootCategories.sort(sortByPosition);
    
    function sortChildren(categories) {
      categories.forEach(cat => {
        if (cat.children.length > 0) {
          cat.children.sort(sortByPosition);
          sortChildren(cat.children);
        }
      });
    }
    
    sortChildren(rootCategories);
    
    logInfo('✅ Простая иерархическая структура каталога построена');
    
    return rootCategories;
    
  } catch (error) {
    handleError(error, 'Построение простой иерархии категорий каталога');
    return [];
  }
}


/**
 * БЫСТРАЯ версия выбора категорий - без диалога
 */
async function showCategorySelectionFast() {
  const context = "Быстрый выбор категорий каталога";
  
  try {
    logInfo('⚡ Быстрый выбор категорий каталога (без диалога)', null, context);
    
    const catalogStructure = await loadCatalogStructureFast();
    
    if (!catalogStructure || catalogStructure.length === 0) {
      throw new Error('Не удалось загрузить структуру каталога');
    }
    
    // Собираем ВСЕ категории каталога (без проверки товаров)
    const allCategories = [];
    
    function collectAllCategories(categories) {
      for (const category of categories) {
        allCategories.push({
          id: category.id,
          title: category.title,
          productCount: 0 // Не считаем для скорости
        });
        
        if (category.children && category.children.length > 0) {
          collectAllCategories(category.children);
        }
      }
    }
    
    collectAllCategories(catalogStructure);
    
    logInfo(`📊 Найдено всего категорий каталога: ${allCategories.length}`);
    
    // Возвращаем первые 3 категории для быстрого тестирования
    const selectedCategories = allCategories.slice(0, 3);
    
    logInfo(`✅ Автоматически выбрано ${selectedCategories.length} категорий каталога для быстрой загрузки`);
    
    return selectedCategories;
    
  } catch (error) {
    logError('❌ Ошибка быстрого выбора категорий каталога', error, context);
    throw error;
  }
}


/**
 * БЫСТРАЯ версия загрузки товаров - с ограничениями
 */
async function loadProductsFromInSalesFast() {
  const context = "Быстрая загрузка товаров из InSales каталога";
  
  try {
    logInfo('⚡ Начинаем БЫСТРУЮ загрузку товаров из InSales каталога', null, context);
    
    // 1. Быстрый выбор категорий без диалога
    const selectedCategories = await showCategorySelectionFast();
    
    if (!selectedCategories || selectedCategories.length === 0) {
      logWarning('⚠️ Категории не выбраны, загрузка отменена');
      return {
        success: false,
        message: 'Категории не выбраны'
      };
    }
    
    logInfo(`✅ Выбрано категорий каталога: ${selectedCategories.length}`);
    
    // 2. Загружаем товары с ограничениями для скорости
    const loadedProducts = await loadProductsByCategoriesFast(selectedCategories);
    
    if (!loadedProducts || loadedProducts.length === 0) {
      logWarning('⚠️ Товары не найдены в выбранных категориях каталога');
      return {
        success: false,
        message: 'Товары не найдены'
      };
    }
    
    logInfo(`✅ Загружено товаров: ${loadedProducts.length}`);
    
    // 3. Синхронизируем с Google Sheets
    const syncResult = await syncProductData(loadedProducts);
    
    logInfo('✅ БЫСТРАЯ загрузка товаров из InSales каталога завершена', {
      categoriesSelected: selectedCategories.length,
      productsLoaded: loadedProducts.length,
      syncResult: syncResult
    }, context);
    
    return {
      success: true,
      categoriesSelected: selectedCategories.length,
      productsLoaded: loadedProducts.length,
      syncResult: syncResult
    };
    
  } catch (error) {
    logError('❌ Ошибка быстрой загрузки товаров из InSales каталога', error, context);
    throw error;
  }
}


/**
 * БЫСТРАЯ загрузка товаров с ограничениями
 */
async function loadProductsByCategoriesFast(selectedCategories) {
  const context = "Быстрая загрузка товаров по категориям каталога";
  const allProducts = [];
  const processedSKUs = new Set();
  const maxProductsPerCategory = 10; // ОГРАНИЧЕНИЕ для скорости
  
  try {
    logInfo(`⚡ БЫСТРО загружаем товары из ${selectedCategories.length} категорий каталога (макс ${maxProductsPerCategory} с каждой)`, null, context);
    
    for (const category of selectedCategories) {
      try {
        logInfo(`📁 Быстро обрабатываем категорию каталога: "${category.title}"`);
        
        const categoryProducts = await loadProductsFromCategoryFast(category.id, maxProductsPerCategory);
        
        if (categoryProducts && categoryProducts.length > 0) {
          const uniqueProducts = categoryProducts.filter(product => {
            const sku = product.sku || product.id;
            if (processedSKUs.has(sku)) {
              return false;
            }
            processedSKUs.add(sku);
            return true;
          });
          
          allProducts.push(...uniqueProducts);
          
          logInfo(`✅ Быстро загружено ${uniqueProducts.length} товаров из категории каталога "${category.title}"`);
        }
        
        Utilities.sleep(100); // Меньше задержка
        
      } catch (categoryError) {
        logWarning(`⚠️ Ошибка быстрой загрузки товаров из категории каталога "${category.title}": ${categoryError.message}`);
        continue;
      }
    }
    
    logInfo(`✅ БЫСТРЫЙ результат: загружено ${allProducts.length} товаров из каталога`, null, context);
    
    return allProducts;
    
  } catch (error) {
    logError('❌ Ошибка быстрой загрузки товаров по категориям каталога', error, context);
    throw error;
  }
}


/**
 * БЫСТРАЯ загрузка товаров из категории с ограничениями
 */
async function loadProductsFromCategoryFast(categoryId, maxProducts = 10) {
  const context = `Быстрая загрузка товаров категории каталога ${categoryId}`;
  const allProducts = [];
  const perPage = Math.min(maxProducts, 25); // Не больше 25 за раз
  
  try {
    logInfo(`⚡ БЫСТРО загружаем товары категории каталога ${categoryId} (макс ${maxProducts})...`);
    
    // Загружаем только одну страницу для скорости
    const products = await makeInsalesRequest('GET', INSALES_ENDPOINTS.PRODUCTS, null, {
      collection_id: categoryId,
      page: 1,
      per_page: perPage
    });
    
    if (!products || !Array.isArray(products) || products.length === 0) {
      logInfo(`📭 Товары не найдены в категории каталога ${categoryId}`);
      return [];
    }
    
    // Обрабатываем только первые несколько товаров
    const productsToProcess = products.slice(0, maxProducts);
    
    for (const product of productsToProcess) {
      try {
        // Загружаем варианты товара для получения правильных артикулов
        const variants = await loadProductVariants(product.id);
        
        if (variants && variants.length > 0) {
          // Берем только первый вариант для скорости
          const firstVariant = variants[0];
          const variantProduct = {
            ...product,
            id: `${product.id}_${firstVariant.id}`,
            sku: firstVariant.sku,
            variant_id: firstVariant.id,
            variant_sku: firstVariant.sku,
            price: firstVariant.price || product.price,
            title: `${product.title}`.trim(),
            original_product_id: product.id
          };
          
          // НЕ загружаем изображения для скорости
          variantProduct.imageUrls = [];
          allProducts.push(variantProduct);
        } else {
          // Если вариантов нет, добавляем основной товар
          product.sku = `PRODUCT_${product.id}`;
          product.imageUrls = [];
          allProducts.push(product);
        }
      } catch (variantError) {
        logWarning(`⚠️ Пропускаем товар ${product.id}: ${variantError.message}`);
        continue;
      }
    }
    
    logInfo(`✅ БЫСТРО из категории каталога ${categoryId} загружено ${allProducts.length} товаров`, null, context);
    
    return allProducts;
    
  } catch (error) {
    logError(`❌ Ошибка быстрой загрузки товаров категории каталога ${categoryId}`, error, context);
    return [];
  }
}


/**
 * БЫСТРЫЙ тест полного workflow
 */
async function testFastWorkflow() {
  const context = "Быстрый тест workflow";
  
  try {
    logInfo('⚡ Запускаем БЫСТРЫЙ тест загрузки товаров', null, context);
    
    const result = await loadProductsFromInSalesFast();
    
    if (result.success) {
      logInfo('✅ Быстрый тест завершен успешно', result, context);
      
      console.log(`🎉 РЕЗУЛЬТАТЫ БЫСТРОГО ТЕСТА:
📁 Выбрано категорий каталога: ${result.categoriesSelected}
📦 Загружено товаров: ${result.productsLoaded}
➕ Добавлено: ${result.syncResult.added}
🔄 Обновлено: ${result.syncResult.updated}
❌ Ошибок: ${result.syncResult.errors}`);
      
      return result;
    } else {
      logWarning('⚠️ Быстрый тест завершен с предупреждением', result, context);
      return result;
    }
    
  } catch (error) {
    logError('❌ Ошибка быстрого теста', error, context);
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * Переопределяем функцию setCachedData для исправления кэширования
 */
function setCachedData(key, data) {
  return setCachedDataFixed(key, data);
}


/**
 * ===================================================================
 * 💡 ОСНОВНЫЕ ИСПРАВЛЕНИЯ В МОДУЛЕ
 * ===================================================================
 * 
 * ✅ ИСПРАВЛЕНО: Загрузка категорий каталога
 * - Используется /admin/categories.json (структура каталога)
 * - Правильная иерархия категорий с parent_id
 * - Подсчет товаров по категориям каталога
 * 
 * ✅ ИСПРАВЛЕНО: Правильные артикулы товаров
 * - Артикулы берутся из variant.sku (правильное поле)
 * - Загружаются варианты товаров через /admin/products/{id}/variants.json
 * - Каждый вариант создается как отдельный товар с уникальным артикулом
 * 
 * ✅ ИСПРАВЛЕНО: Структура конвертации товаров
 * - Приоритет: variant.sku > product.sku > фолбэк ID
 * - Сохранение связей товар-вариант-категория
 * - Правильное формирование данных для Google Sheets
 * 
 * ✅ ДОБАВЛЕНО: Новые endpoints для вариантов
 * - PRODUCT_VARIANTS: '/admin/products/{product_id}/variants.json'
 * - Функция loadProductVariants() для загрузки вариантов
 * - Обогащение товаров правильными артикулами
 * 
 * ИСПОЛЬЗОВАНИЕ:
 * 1. testConnection() - тест подключения
 * 2. testSingleProductWithVariants() - тест товара с вариантами
 * 3. testCatalogStructure() - тест структуры каталога
 * 4. testFullWorkflowWithCorrectSKU() - полный тест с правильными артикулами
 * 5. loadProductsFromInSales() - основная функция загрузки
 * 
 * ===================================================================
 */
/**
 * УПРОЩЕННАЯ загрузка товаров без диалога выбора категорий
 * Загружает товары из первых N родительских категорий
 */
async function loadProductsSimplified() {
  const context = "Упрощенная загрузка товаров";
  
  try {
    logInfo('🚀 Начинаем упрощенную загрузку товаров', null, context);
    
    // Параметры загрузки
    const MAX_CATEGORIES = 5; // Количество категорий для загрузки
    const MAX_PRODUCTS_PER_CATEGORY = 20; // Максимум товаров из каждой категории
    
    // 1. Загружаем только структуру категорий БЕЗ подсчета товаров
    const categories = await loadCategoriesWithoutCount();
    
    if (!categories || categories.length === 0) {
      throw new Error('Не удалось загрузить категории');
    }
    
    // 2. Берем только родительские категории (без parent_id)
    const rootCategories = categories.filter(cat => !cat.parent_id);
    
    logInfo(`📊 Найдено ${rootCategories.length} родительских категорий`);
    
    // 3. Выбираем первые N категорий
    const selectedCategories = rootCategories.slice(0, MAX_CATEGORIES);
    
    logInfo(`✅ Выбрано ${selectedCategories.length} категорий для загрузки`);
    
    // 4. Загружаем товары с ограничениями
    const loadedProducts = await loadProductsWithLimit(selectedCategories, MAX_PRODUCTS_PER_CATEGORY);
    
    if (!loadedProducts || loadedProducts.length === 0) {
      return {
        success: false,
        message: 'Товары не найдены'
      };
    }
    
    logInfo(`✅ Загружено товаров: ${loadedProducts.length}`);
    
    // 5. Синхронизируем с Google Sheets
    const syncResult = await syncProductData(loadedProducts);
    
    return {
      success: true,
      categoriesSelected: selectedCategories.length,
      productsLoaded: loadedProducts.length,
      syncResult: syncResult
    };
    
  } catch (error) {
    logError('❌ Ошибка упрощенной загрузки', error, context);
    throw error;
  }
}

/**
 * Загружает категории БЕЗ подсчета товаров
 */
async function loadCategoriesWithoutCount() {
  try {
    logInfo('📁 Загружаем категории БЕЗ подсчета товаров');
    
    // Загружаем все категории одним запросом
    const allCategories = [];
    let page = 1;
    const perPage = 250; // Максимум за раз
    
    while (true) {
      const categories = await makeInsalesRequest('GET', INSALES_ENDPOINTS.COLLECTIONS, null, {
        page: page,
        per_page: perPage
      });
      
      if (!categories || !Array.isArray(categories) || categories.length === 0) {
        break;
      }
      
      allCategories.push(...categories);
      
      if (categories.length < perPage) {
        break;
      }
      
      page++;
      Utilities.sleep(200); // Небольшая задержка
    }
    
    logInfo(`✅ Загружено ${allCategories.length} категорий`);
    
    // Добавляем нулевой счетчик товаров (не загружаем реальное количество)
    return allCategories.map(cat => ({
      ...cat,
      productCount: 0 // Заглушка вместо реального подсчета
    }));
    
  } catch (error) {
    logError('❌ Ошибка загрузки категорий', error);
    return [];
  }
}

/**
 * Загружает товары с ограничением по количеству
 */
async function loadProductsWithLimit(categories, maxPerCategory) {
  const allProducts = [];
  const processedSKUs = new Set();
  
  try {
    for (const category of categories) {
      logInfo(`📦 Загружаем товары из категории "${category.title}" (максимум ${maxPerCategory})`);
      
      // Загружаем только первую страницу с ограничением
      const products = await makeInsalesRequest('GET', INSALES_ENDPOINTS.PRODUCTS, null, {
        collection_id: category.id,
        page: 1,
        per_page: Math.min(maxPerCategory, 100) // Не больше 100 за раз
      });
      
      if (products && Array.isArray(products)) {
        // Обрабатываем товары и получаем артикулы из вариантов
        for (const product of products.slice(0, maxPerCategory)) {
          try {
            // Быстрая проверка на дубликаты
            if (processedSKUs.has(product.id)) {
              continue;
            }
            processedSKUs.add(product.id);
            
            // Загружаем варианты для получения артикулов
            const variants = await loadProductVariants(product.id);
            
            if (variants && variants.length > 0) {
              // Берем только первый вариант для скорости
              const variant = variants[0];
              const productWithSKU = {
                ...product,
                sku: variant.sku,
                variant_id: variant.id,
                variant_sku: variant.sku,
                imageUrls: [] // Пропускаем изображения для скорости
              };
              
              allProducts.push(productWithSKU);
            }
            
          } catch (err) {
            logWarning(`⚠️ Пропускаем товар ${product.id}: ${err.message}`);
          }
        }
      }
      
      Utilities.sleep(300); // Задержка между категориями
    }
    
    return allProducts;
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров с ограничением', error);
    return allProducts;
  }
}

/**
 * Главная функция для меню - загрузка товаров из InSales
 */
function loadProductsFromInSalesMenu() {
  const context = "Загрузка товаров из InSales";
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Загружаем список категорий...',
      '⏳ Подождите',
      -1
    );
    
    // 1. Загружаем категории синхронно
    const categories = loadRootCategoriesOnlySync();

    // ОТЛАДКА: проверяем что реально загрузилось
    console.log('🔍 ОТЛАДКА: Загружено категорий:', categories.length);
    categories.forEach((cat, index) => {
      if (index < 10) { // Показываем первые 10
        console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id})`);
      }
    });
    
    if (!categories || categories.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Не удалось загрузить категории',
        '❌ Ошибка',
        5
      );
      return {
        success: false,
        message: 'Не удалось загрузить категории'
      };
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Загружено ${categories.length} родительских категорий`,
      '✅ Готово',
      2
    );
    
    // 2. Показываем диалог выбора категорий
    showCategorySearchDialog();
    return; // Выходим, так как обработка теперь в новой функции
    
    if (!selectedCategories || selectedCategories.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Категории не выбраны',
        '⚠️ Отменено',
        3
      );
      return {
        success: false,
        message: 'Категории не выбраны'
      };
    }
    
    // 3. Спрашиваем количество товаров
    const maxProducts = askProductsLimit();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Загружаем товары из ${selectedCategories.length} категорий...`,
      '⏳ Загрузка',
      -1
    );
    
    // 4. Загружаем товары с изображениями
    const allProducts = [];
    const processedSKUs = new Set();
    
    for (let i = 0; i < selectedCategories.length; i++) {
      const category = selectedCategories[i];
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Загружаем категорию ${i + 1}/${selectedCategories.length}: ${category.title}`,
        '⏳ Загрузка',
        -1
      );
      
      const categoryProducts = loadProductsFromCategoryWithImages(category.id, maxProducts);
      
      // Фильтруем дубликаты
      for (const product of categoryProducts) {
        const key = product.sku || product.id;
        if (!processedSKUs.has(key)) {
          processedSKUs.add(key);
          allProducts.push(product);
        }
      }
      
      // Небольшая задержка между категориями
      Utilities.sleep(300);
    }
    
    if (allProducts.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Товары не найдены в выбранных категориях',
        '⚠️ Предупреждение',
        5
      );
      return {
        success: false,
        message: 'Товары не найдены'
      };
    }
    
    logInfo(`✅ Загружено ${allProducts.length} уникальных товаров`);
    
    // 5. Синхронизируем с Google Sheets
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Сохраняем ${allProducts.length} товаров в таблицу...`,
      '💾 Сохранение',
      -1
    );
    
    const syncResult = syncProductData(allProducts);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Загружено ${allProducts.length} товаров\nДобавлено: ${syncResult.added}\nОбновлено: ${syncResult.updated}`,
      '✅ Готово',
      10
    );
    
    return {
      success: true,
      categoriesSelected: selectedCategories.length,
      productsLoaded: allProducts.length,
      syncResult: syncResult
    };
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Ошибка: ${error.message}`,
      '❌ Ошибка',
      10
    );
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Синхронная загрузка родительских категорий
 */
function loadRootCategoriesOnlySync() {
  try {
    logInfo('📁 Загружаем коллекции каталога (синхронно)');
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const allCollections = [];
    let page = 1;
    const perPage = 250;
    
    // Загружаем все коллекции постранично
    while (true) {
      // ИСПРАВЛЕНО: используем /admin/collections.json вместо /admin/categories.json
      const url = `${credentials.baseUrl}${INSALES_ENDPOINTS.COLLECTIONS}?per_page=${perPage}&page=${page}&is_hidden=false`;
      
      console.log('🔍 Запрос к URL:', url);
      
      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 429) {
        // Rate limit - ждем и повторяем
        Utilities.sleep(2000);
        continue;
      }
      
      if (responseCode !== 200) {
        throw new Error(`Ошибка API: ${responseCode} - ${response.getContentText()}`);
      }
      
      const collections = JSON.parse(response.getContentText());
      
      console.log('📦 Получено коллекций на странице:', collections.length);
      collections.forEach((col, index) => {
        console.log(`  ${index + 1}. "${col.title}" (ID: ${col.id}, parent_id: ${col.parent_id || 'null'})`);
      });
      
      if (!collections || collections.length === 0) {
        break;
      }
      
      allCollections.push(...collections);
      
      if (collections.length < perPage || page >= 10) {
        break;
      }
      
      page++;
      Utilities.sleep(200);
    }
    
    console.log('🗂️ ВСЕ КОЛЛЕКЦИИ ПЕРЕД ФИЛЬТРАЦИЕЙ:');
    allCollections.forEach((col, index) => {
      console.log(`  ${index + 1}. "${col.title}" (ID: ${col.id}, parent_id: ${col.parent_id || 'null'})`);
    });
    
    // Фильтруем только родительские коллекции
    const rootCollections = allCollections.filter(col => !col.parent_id || col.parent_id === null);
    
    console.log('🌳 РОДИТЕЛЬСКИЕ КОЛЛЕКЦИИ ПОСЛЕ ФИЛЬТРАЦИИ:');
    rootCollections.forEach((col, index) => {
      console.log(`  ${index + 1}. "${col.title}" (ID: ${col.id})`);
    });
    
    logInfo(`✅ Найдено ${rootCollections.length} родительских коллекций из ${allCollections.length} всего`);
    
    // Сортируем по названию
    rootCollections.sort((a, b) => a.title.localeCompare(b.title));
    
    return rootCollections;
    
  } catch (error) {
    logError('❌ Ошибка загрузки коллекций', error);
    throw error;
  }
}

/**
 * Получение учетных данных синхронно
 */
function getInsalesCredentialsSync() {
  const apiKey = getSetting('insalesApiKey') || getSetting('InSales_API_Key');
  const password = getSetting('insalesPassword') || getSetting('InSales_Password');
  const shop = getSetting('insalesShop') || getSetting('InSales_Shop');
  
  if (!apiKey || !password || !shop) {
    logError('❌ Отсутствуют учетные данные InSales');
    return null;
  }
  
  return {
    apiKey: apiKey,
    password: password,
    shop: shop,
    baseUrl: `https://${shop}`
  };
}

/**
 * Показывает диалог выбора с готовыми данными
 */
function showCategorySelectionDialogWithData(categories) {
  try {
    // Сохраняем категории для диалога
    const cache = CacheService.getUserCache();
    cache.put('dialogCategories', JSON.stringify(categories), 300); // 5 минут
    
    // Создаем HTML
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { 
              font-family: Arial, sans-serif; 
              padding: 20px;
              margin: 0;
            }
            h3 { 
              margin-top: 0;
              color: #333;
            }
            .info {
              color: #666;
              margin-bottom: 15px;
            }
            .category-list { 
              max-height: 400px; 
              overflow-y: auto; 
              border: 1px solid #ddd; 
              padding: 10px;
              margin: 10px 0;
              background: #f9f9f9;
            }
            .category-item { 
              padding: 8px 5px; 
              display: flex;
              align-items: center;
              border-bottom: 1px solid #eee;
            }
            .category-item:last-child {
              border-bottom: none;
            }
            .category-item:hover {
              background: #f0f0f0;
            }
            .category-item input { 
              margin-right: 10px; 
            }
            .category-item label {
              cursor: pointer;
              flex: 1;
            }
            .buttons { 
              margin-top: 20px; 
              text-align: right;
            }
            button { 
              padding: 10px 20px; 
              margin-left: 10px;
              cursor: pointer;
              border: 1px solid #ddd;
              background: #fff;
              border-radius: 4px;
              font-size: 14px;
            }
            button:hover {
              background: #f5f5f5;
            }
            .select-all { 
              margin: 10px 0;
              padding: 10px;
              background: #e8f0fe;
              border-radius: 4px;
            }
            .primary { 
              background: #4285f4; 
              color: white; 
              border: none;
            }
            .primary:hover {
              background: #3367d6;
            }
            .search-box {
              width: 100%;
              padding: 10px;
              margin-bottom: 10px;
              border: 1px solid #ddd;
              border-radius: 4px;
              box-sizing: border-box;
              font-size: 14px;
            }
            .loading {
              text-align: center;
              padding: 40px;
              color: #666;
            }
            .spinner {
              border: 3px solid #f3f3f3;
              border-top: 3px solid #4285f4;
              border-radius: 50%;
              width: 40px;
              height: 40px;
              animation: spin 1s linear infinite;
              margin: 20px auto;
            }
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          </style>
        </head>
        <body>
          <h3>Выберите категории для загрузки товаров</h3>
          
          <div id="content">
            <div class="loading">
              <div class="spinner"></div>
              <p>Загрузка категорий...</p>
            </div>
          </div>
          
          <script>
            let allCategories = [];
            
            // Загружаем категории при открытии диалога
            window.onload = function() {
              google.script.run
                .withSuccessHandler(loadCategories)
                .withFailureHandler(showError)
                .getCategoriesForDialog();
            };
            
            function loadCategories(categories) {
              if (!categories || categories.length === 0) {
                showError({ message: 'Категории не найдены' });
                return;
              }
              
              allCategories = categories;
              
              const content = document.getElementById('content');
              content.innerHTML = \`
                <p class="info">Всего родительских категорий: \${categories.length}</p>
                
                <input type="text" 
                       class="search-box" 
                       id="searchBox" 
                       placeholder="Поиск категорий..." 
                       onkeyup="filterCategories()">
                
                <div class="select-all">
                  <label>
                    <input type="checkbox" id="selectAll" onchange="toggleAll()">
                    <strong>Выбрать все</strong>
                  </label>
                </div>
                
                <div class="category-list" id="categoryList">
                  \${categories.map((cat, index) => \`
                    <div class="category-item" data-title="\${cat.title.toLowerCase()}">
                      <input type="checkbox" 
                             id="cat_\${index}" 
                             value="\${cat.id}" 
                             data-title="\${cat.title}">
                      <label for="cat_\${index}">\${cat.title}</label>
                    </div>
                  \`).join('')}
                </div>
                
                <div class="buttons">
                  <button onclick="google.script.host.close()">Отмена</button>
                  <button class="primary" onclick="submitSelection()">
                    Загрузить выбранные
                  </button>
                </div>
              \`;
            }
            
            function showError(error) {
              document.getElementById('content').innerHTML = 
                '<div class="loading">Ошибка загрузки категорий: ' + error.message + '</div>';
            }
            
            function toggleAll() {
              const selectAll = document.getElementById('selectAll').checked;
              const checkboxes = document.querySelectorAll('.category-item:not([style*="display: none"]) input[type="checkbox"]');
              checkboxes.forEach(cb => {
                if (cb.id !== 'selectAll') {
                  cb.checked = selectAll;
                }
              });
            }
            
            function filterCategories() {
              const searchText = document.getElementById('searchBox').value.toLowerCase();
              const items = document.querySelectorAll('.category-item');
              
              items.forEach(item => {
                const title = item.getAttribute('data-title');
                if (title.includes(searchText)) {
                  item.style.display = 'flex';
                } else {
                  item.style.display = 'none';
                }
              });
              
              // Сбрасываем "выбрать все" при фильтрации
              document.getElementById('selectAll').checked = false;
            }
            
            function submitSelection() {
              const selected = [];
              const checkboxes = document.querySelectorAll('.category-item input[type="checkbox"]:checked');
              
              checkboxes.forEach(cb => {
                if (cb.id !== 'selectAll') {
                  selected.push({
                    id: parseInt(cb.value),
                    title: cb.getAttribute('data-title')
                  });
                }
              });
              
              if (selected.length === 0) {
                alert('Пожалуйста, выберите хотя бы одну категорию');
                return;
              }
              
              // Отправляем выбор и закрываем диалог
              google.script.run
                .withSuccessHandler(() => google.script.host.close())
                .withFailureHandler((error) => {
                  alert('Ошибка: ' + error.message);
                })
                .saveSelectedCategories(selected);
            }
          </script>
        </body>
      </html>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(500)
      .setHeight(650);
    
    // Показываем диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Выбор категорий InSales');
    
    // Ждем результата выбора
    const userProperties = PropertiesService.getUserProperties();
    const startTime = new Date().getTime();
    const timeout = 300000; // 5 минут
    
    while (new Date().getTime() - startTime < timeout) {
      Utilities.sleep(500);
      
      const selected = userProperties.getProperty('selectedCategories');
      if (selected) {
        userProperties.deleteProperty('selectedCategories');
        return JSON.parse(selected);
      }
      
      // Проверяем, не закрыт ли диалог
      const dialogClosed = userProperties.getProperty('dialogClosed');
      if (dialogClosed) {
        userProperties.deleteProperty('dialogClosed');
        return null;
      }
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка показа диалога выбора', error);
    throw error;
  }
}

/**
 * Функция для получения категорий в диалоге
 */
function getCategoriesForDialog() {
  try {
    // Всегда загружаем свежие данные без кэша
    logInfo('🔄 Загружаем свежие категории для диалога');
    return loadRootCategoriesOnlySync();
  } catch (error) {
    logError('❌ Ошибка загрузки категорий для диалога', error);
    throw new Error('Не удалось загрузить категории: ' + error.message);
  }
}

/**
 * Загружает товары из категории с изображениями (синхронно)
 */
function loadProductsFromCategoryWithImages(categoryId, maxProducts) {
  const products = [];
  
  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }
    
    let page = 1;
    let loaded = 0;
    const perPage = Math.min(maxProducts, 100);
    
    while (loaded < maxProducts) {
      const url = `${credentials.baseUrl}${INSALES_ENDPOINTS.PRODUCTS}` +
                  `?collection_id=${categoryId}` +
                  `&per_page=${perPage}` +
                  `&page=${page}`;
      
      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 429) {
        // Rate limit
        Utilities.sleep(2000);
        continue;
      }
      
      if (response.getResponseCode() !== 200) {
        logWarning(`⚠️ Ошибка загрузки товаров категории ${categoryId}: ${response.getResponseCode()}`);
        break;
      }
      
      const rawProducts = JSON.parse(response.getContentText());
      
      if (!rawProducts || rawProducts.length === 0) {
        break;
      }
      
      // Обрабатываем каждый товар
      for (const product of rawProducts) {
        if (loaded >= maxProducts) break;
        
        try {
          // Загружаем варианты для получения артикула
          const variants = loadProductVariantsSync(product.id);
          
          if (variants && variants.length > 0) {
            const variant = variants[0];
            
            // Форматируем изображения
            let imageUrls = [];
            if (product.images && Array.isArray(product.images)) {
              imageUrls = product.images.map(img => 
                img.original_url || img.medium_url || img.small_url || img.url
              ).filter(Boolean);
            }
            
            products.push({
              ...product,
              sku: variant.sku || `VAR_${variant.id}`,
              variant_id: variant.id,
              variant_sku: variant.sku,
              price: variant.price || product.price,
              imageUrls: imageUrls,
              images: imageUrls.join('\n'), // Для колонки E
              original_product_id: product.id
            });
            
            loaded++;
          } else {
            // Если нет вариантов, используем сам товар
            let imageUrls = [];
            if (product.images && Array.isArray(product.images)) {
              imageUrls = product.images.map(img => 
                img.original_url || img.medium_url || img.small_url || img.url
              ).filter(Boolean);
            }
            
            products.push({
              ...product,
              sku: product.sku || `PRODUCT_${product.id}`,
              imageUrls: imageUrls,
              images: imageUrls.join('\n') // Для колонки E
            });
            
            loaded++;
          }
          
        } catch (err) {
          logWarning(`⚠️ Ошибка обработки товара ${product.id}: ${err.message}`);
        }
      }
      
      if (rawProducts.length < perPage) {
        break;
      }
      
      page++;
      Utilities.sleep(300);
    }
    
    logInfo(`✅ Загружено ${products.length} товаров из категории ${categoryId}`);
    return products;
    
  } catch (error) {
    logError(`❌ Ошибка загрузки товаров категории ${categoryId}`, error);
    return products;
  }
}

/**
 * Синхронная загрузка вариантов товара
 */
function loadProductVariantsSync(productId) {
  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) return [];
    
    const url = `${credentials.baseUrl}/admin/products/${productId}/variants.json`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    
    return [];
    
  } catch (error) {
    logDebug(`Не удалось загрузить варианты для товара ${productId}`);
    return [];
  }
}

/**
 * Синхронная загрузка родительских категорий
 */
function loadRootCategoriesOnlySync() {
  try {
    logInfo('📁 Загружаем родительские категории (синхронно)');
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }
    
    // Загружаем первую страницу категорий
    const url = `${credentials.baseUrl}${INSALES_ENDPOINTS.COLLECTIONS}?per_page=${perPage}&page=${page}&is_hidden=false`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`Ошибка API: ${responseCode}`);
    }
    
    const allCategories = JSON.parse(response.getContentText());

    // ОТЛАДКА: что реально загрузилось
    console.log('🔍 ОТЛАДКА allCategories.length:', allCategories.length);
    if (allCategories.length > 0) {
      console.log('Первые 3 коллекции:');
      allCategories.slice(0, 3).forEach((cat, index) => {
        console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id}, parent_id: ${cat.parent_id})`);
      });
    }
    
    // Фильтруем только родительские категории
    const rootCategories = allCategories.filter(cat => {
      // Основные товарные категории по названиям
      const mainCategories = [
        'Бинокли', 'Телескоп', 'Микроскоп', 'Зрительные трубы', 
        'Монокуляр', 'Аксессуары для биноклей', 'Аксессуары для зрительных труб',
        'Аксессуары для микроскопов', 'Прицел'
      ];
      
      return mainCategories.some(mainCat => cat.title.includes(mainCat));
    });

    logInfo(`✅ Найдено ${rootCategories.length} родительских категорий`);
    
    // Сортируем по названию
    rootCategories.sort((a, b) => a.title.localeCompare(b.title));
    
    // ОТЛАДКА: что после фильтрации
      console.log('🔍 ОТЛАДКА rootCategories.length:', rootCategories.length);
      if (rootCategories.length > 0) {
        console.log('Корневые коллекции:');
        rootCategories.slice(0, 5).forEach((cat, index) => {
          console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id})`);
        });
      }

    return rootCategories;
    
  } catch (error) {
    logError('❌ Ошибка загрузки категорий', error);
    return [];
  }
}

/**
 * Получение учетных данных синхронно
 */
function getInsalesCredentialsSync() {
  const apiKey = getSetting('insalesApiKey') || getSetting('InSales_API_Key');
  const password = getSetting('insalesPassword') || getSetting('InSales_Password');
  const shop = getSetting('insalesShop') || getSetting('InSales_Shop');
  
  if (!apiKey || !password || !shop) {
    return null;
  }
  
  return {
    apiKey: apiKey,
    password: password,
    shop: shop,
    baseUrl: `https://${shop}`
  };
}

/**
 * Показывает диалог выбора с переданными данными
 */
function showCategorySelectionDialogWithData(categories) {
  try {
    // Сохраняем категории для диалога
    const cache = CacheService.getUserCache();
    cache.put('dialogCategories', JSON.stringify(categories), 300); // 5 минут
    
    // Создаем HTML шаблон
    const template = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            h3 { margin-top: 0; }
            .category-list { 
              max-height: 400px; 
              overflow-y: auto; 
              border: 1px solid #ddd; 
              padding: 10px;
              margin: 10px 0;
            }
            .category-item { 
              padding: 5px 0; 
              display: flex;
              align-items: center;
            }
            .category-item input { 
              margin-right: 10px; 
            }
            .buttons { 
              margin-top: 20px; 
              text-align: right;
            }
            button { 
              padding: 8px 16px; 
              margin-left: 10px;
              cursor: pointer;
            }
            .select-all { 
              margin: 10px 0;
            }
            .primary { 
              background: #4285f4; 
              color: white; 
              border: none;
            }
            .search-box {
              width: 100%;
              padding: 8px;
              margin-bottom: 10px;
              border: 1px solid #ddd;
              box-sizing: border-box;
            }
            .loading {
              text-align: center;
              padding: 20px;
              color: #666;
            }
          </style>
        </head>
        <body>
          <h3>Выберите категории для загрузки товаров</h3>
          
          <div id="content">
            <div class="loading">Загрузка категорий...</div>
          </div>
          
          <script>
            let allCategories = [];
            
            // Загружаем категории при открытии диалога
            google.script.run
              .withSuccessHandler(loadCategories)
              .withFailureHandler(showError)
              .getCategoriesForDialog();
            
            function loadCategories(categories) {
              allCategories = categories;
              
              const content = document.getElementById('content');
              content.innerHTML = \`
                <p>Всего родительских категорий: \${categories.length}</p>
                
                <input type="text" class="search-box" id="searchBox" placeholder="Поиск категорий..." onkeyup="filterCategories()">
                
                <div class="select-all">
                  <label>
                    <input type="checkbox" id="selectAll" onchange="toggleAll()">
                    Выбрать все
                  </label>
                </div>
                
                <div class="category-list" id="categoryList">
                  \${categories.map((cat, index) => \`
                    <div class="category-item" data-title="\${cat.title.toLowerCase()}">
                      <input type="checkbox" id="cat_\${index}" value="\${cat.id}" data-title="\${cat.title}">
                      <label for="cat_\${index}">\${cat.title}</label>
                    </div>
                  \`).join('')}
                </div>
                
                <div class="buttons">
                  <button onclick="google.script.host.close()">Отмена</button>
                  <button class="primary" onclick="submitSelection()">Загрузить выбранные</button>
                </div>
              \`;
            }
            
            function showError(error) {
              document.getElementById('content').innerHTML = 
                '<div class="loading">Ошибка загрузки категорий: ' + error.message + '</div>';
            }
            
            function toggleAll() {
              const selectAll = document.getElementById('selectAll').checked;
              const checkboxes = document.querySelectorAll('.category-item:not([style*="display: none"]) input[type="checkbox"]');
              checkboxes.forEach(cb => {
                if (cb.id !== 'selectAll') {
                  cb.checked = selectAll;
                }
              });
            }
            
            function filterCategories() {
              const searchText = document.getElementById('searchBox').value.toLowerCase();
              const items = document.querySelectorAll('.category-item');
              
              items.forEach(item => {
                const title = item.getAttribute('data-title');
                if (title.includes(searchText)) {
                  item.style.display = 'flex';
                } else {
                  item.style.display = 'none';
                }
              });
            }
            
            function submitSelection() {
              const selected = [];
              const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
              
              checkboxes.forEach(cb => {
                if (cb.id !== 'selectAll') {
                  selected.push({
                    id: parseInt(cb.value),
                    title: cb.getAttribute('data-title')
                  });
                }
              });
              
              if (selected.length === 0) {
                alert('Пожалуйста, выберите хотя бы одну категорию');
                return;
              }
              
              google.script.run
                .withSuccessHandler(() => google.script.host.close())
                .saveSelectedCategories(selected);
            }
          </script>
        </body>
      </html>
    `);
    
    const htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(600);
    
    // Показываем диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Выбор категорий InSales');
    
    // Ждем результата
    const userProperties = PropertiesService.getUserProperties();
    let attempts = 0;
    
    while (attempts < 60) { // Ждем до 60 секунд
      Utilities.sleep(1000);
      const selected = userProperties.getProperty('selectedCategories');
      
      if (selected) {
        userProperties.deleteProperty('selectedCategories');
        return JSON.parse(selected);
      }
      
      attempts++;
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка показа диалога', error);
    return null;
  }
}

/**
 * Функция для получения категорий в диалоге
 */
function getCategoriesForDialog() {
  try {
    // Всегда загружаем свежие данные без кэша
    logInfo('🔄 Загружаем свежие категории для диалога');
    return loadRootCategoriesOnlySync();
  } catch (error) {
    logError('❌ Ошибка загрузки категорий для диалога', error);
    throw new Error('Не удалось загрузить категории: ' + error.message);
  }
}

/**
 * Загружает товары из выбранных категорий
 */
function loadProductsFromSelectedCategories(selectedCategories, maxProducts) {
  try {
    const allProducts = [];
    const processedSKUs = new Set();
    
    for (const category of selectedCategories) {
      logInfo(`📦 Загружаем товары из категории "${category.title}"`);
      
      const categoryProducts = loadProductsFromCategoryWithImages(category.id, maxProducts);
      
      for (const product of categoryProducts) {
        if (!processedSKUs.has(product.id)) {
          processedSKUs.add(product.id);
          allProducts.push(product);
        }
      }
    }
    
    // Синхронизируем с Google Sheets
    const syncResult = syncProductData(allProducts);
    
    return {
      success: true,
      categoriesSelected: selectedCategories.length,
      productsLoaded: allProducts.length,
      syncResult: syncResult
    };
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Загружает товары с изображениями синхронно
 */
function loadProductsFromCategoryWithImages(categoryId, maxProducts) {
  const products = [];
  
  try {
    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}${INSALES_ENDPOINTS.PRODUCTS}?collection_id=${categoryId}&per_page=${Math.min(maxProducts, 100)}&page=1`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const rawProducts = JSON.parse(response.getContentText());
      
      for (const product of rawProducts.slice(0, maxProducts)) {
        // Загружаем варианты
        const variants = loadProductVariantsSync(product.id);
        
        if (variants && variants.length > 0) {
          const variant = variants[0];
          
          // Форматируем изображения
          let imageUrls = [];
          if (product.images && Array.isArray(product.images)) {
            imageUrls = product.images.map(img => 
              img.original_url || img.medium_url || img.small_url
            ).filter(Boolean);
          }
          
          products.push({
            ...product,
            sku: variant.sku,
            variant_id: variant.id,
            variant_sku: variant.sku,
            imageUrls: imageUrls,
            images: imageUrls.join('\n')
          });
        }
      }
    }
    
    return products;
    
  } catch (error) {
    logError(`❌ Ошибка загрузки товаров категории ${categoryId}`, error);
    return products;
  }
}

/**
 * Синхронная загрузка вариантов
 */
function loadProductVariantsSync(productId) {
  try {
    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/products/${productId}/variants.json`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
    
    return [];
    
  } catch (error) {
    return [];
  }
}

/**
 * Сохраняет выбранные категории (вызывается из диалога)
 */
function saveSelectedCategories(selected) {
  PropertiesService.getUserProperties().setProperty('selectedCategories', JSON.stringify(selected));
}

/**
 * Спрашивает у пользователя лимит товаров
 */
function askProductsLimit() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Количество товаров',
    'Сколько товаров загрузить из каждой категории? (максимум)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const limit = parseInt(response.getResponseText());
    if (!isNaN(limit) && limit > 0) {
      return Math.min(limit, 500); // Максимум 500
    }
  }
  
  return 100; // По умолчанию 100
}

function loadRootCategoriesOnlySync() {
  try {
    logInfo('📁 Загружаем родительские категории (синхронно)');
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const allCategories = [];
    let page = 1;
    const perPage = 250;
    
    // Загружаем все категории постранично
    while (true) {
      const url = `${credentials.baseUrl}${INSALES_ENDPOINTS.COLLECTIONS}?per_page=${perPage}&page=${page}&is_hidden=false`;
      
      // ОТЛАДКА: выводим URL запроса
      console.log('🔍 Запрос к URL:', url);
      
      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 429) {
        // Rate limit - ждем и повторяем
        Utilities.sleep(2000);
        continue;
      }
      
      if (responseCode !== 200) {
        throw new Error(`Ошибка API: ${responseCode} - ${response.getContentText()}`);
      }
      
      const categories = JSON.parse(response.getContentText());
      
      // ОТЛАДКА: выводим все полученные категории
      console.log('📦 Получено категорий на странице:', categories.length);
      categories.forEach((cat, index) => {
        console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id}, parent_id: ${cat.parent_id || 'null'})`);
      });
      
      if (!categories || categories.length === 0) {
        break;
      }
      
      allCategories.push(...categories);
      
      if (categories.length < perPage || page >= 10) {
        break;
      }
      
      page++;
      Utilities.sleep(200);
    }
    
    // ОТЛАДКА: показываем все категории перед фильтрацией
    console.log('🗂️ ВСЕ КАТЕГОРИИ ПЕРЕД ФИЛЬТРАЦИЕЙ:');
    allCategories.forEach((cat, index) => {
      console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id}, parent_id: ${cat.parent_id || 'null'})`);
    });
    
    // Фильтруем только родительские категории
    const rootCategories = allCategories.filter(cat => {
      // Основные товарные категории по названиям
      const mainCategories = [
        'Бинокли', 'Телескоп', 'Микроскоп', 'Зрительные трубы', 
        'Монокуляр', 'Аксессуары для биноклей', 'Аксессуары для зрительных труб',
        'Аксессуары для микроскопов', 'Прицел'
      ];
      
      return mainCategories.some(mainCat => cat.title.includes(mainCat));
    });
    
    // ОТЛАДКА: показываем результат фильтрации
    console.log('🌳 РОДИТЕЛЬСКИЕ КАТЕГОРИИ ПОСЛЕ ФИЛЬТРАЦИИ:');
    rootCategories.forEach((cat, index) => {
      console.log(`  ${index + 1}. "${cat.title}" (ID: ${cat.id})`);
    });
    
    logInfo(`✅ Найдено ${rootCategories.length} родительских категорий из ${allCategories.length} всего`);
    
    // Сортируем по названию
    rootCategories.sort((a, b) => a.title.localeCompare(b.title));
    
    return rootCategories;
    
  } catch (error) {
    logError('❌ Ошибка загрузки категорий', error);
    throw error;
  }
}

function showCategorySearchDialog() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .search-field { 
            width: 100%; 
            padding: 10px; 
            margin: 10px 0; 
            border: 1px solid #ddd; 
            border-radius: 4px;
          }
          .buttons { margin-top: 20px; text-align: right; }
          button { padding: 10px 20px; margin-left: 10px; }
          .primary { background: #4285f4; color: white; border: none; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h3>Поиск категорий для загрузки</h3>
        <label>Введите названия категорий (через запятую):</label>
        <textarea id="categoryNames" class="search-field" rows="4" 
                  placeholder="Например: Бинокли 10x25, Монокуляры, Аксессуары"></textarea>
        <div class="buttons">
          <button onclick="google.script.host.close()">Отмена</button>
          <button class="primary" onclick="searchAndLoad()">Найти и загрузить</button>
        </div>
        <script>
          function searchAndLoad() {
            const input = document.getElementById('categoryNames').value.trim();
            if (!input) {
              alert('Введите названия категорий');
              return;
            }
            const searchTerms = input.split(',').map(term => term.trim());
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .loadProductsBySearchTerms(searchTerms);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(500).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Поиск категорий InSales');
}

function loadProductsBySearchTerms(searchTerms) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Ищем категории...', 'Поиск', -1);
    
    const allCategories = loadAllCategoriesQuick();
    const foundCategories = allCategories.filter(cat => {
      return searchTerms.some(term => 
        cat.title.toLowerCase().includes(term.toLowerCase())
      );
    });
    
    if (foundCategories.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Категории не найдены', 'Предупреждение', 5);
      return;
    }
    
    // Показываем диалог множественного выбора
    showCategorySelectionFromFound(foundCategories);
    
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка: ' + error.message, 'Ошибка', 10);
  }
}

function loadAllCategoriesQuick() {
  const credentials = getInsalesCredentialsSync();
  const allCategories = [];
  let page = 1;
  
  while (page <= 4) { // Ограничим для скорости
    const url = `${credentials.baseUrl}/admin/collections.json?per_page=250&page=${page}&is_hidden=false`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) break;
    
    const categories = JSON.parse(response.getContentText());
    if (!categories || categories.length === 0) break;
    
    allCategories.push(...categories);
    if (categories.length < 250) break;
    page++;
  }
  
  return allCategories;
}

function getFoundCategoriesForDialog() {
  const cache = CacheService.getUserCache();
  const categoriesJson = cache.get('foundCategories');
  return categoriesJson ? JSON.parse(categoriesJson) : [];
}

function loadProductsFromSelectedCategories(selectedCategories) {
  try {
    const maxProducts = askProductsLimit();
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Загружаем товары...', 'Загрузка', -1);
    
    const allProducts = [];
    for (const category of selectedCategories) {
      const categoryProducts = loadProductsFromCategoryWithImages(category.id, maxProducts);
      allProducts.push(...categoryProducts);
    }
    
    const syncResult = syncProductData(allProducts);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Загружено ${allProducts.length} товаров из ${selectedCategories.length} категорий`,
      'Готово',
      10
    );
    
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка: ' + error.message, 'Ошибка', 10);
  }
}

function showCategorySelectionFromFound(foundCategories) {
  // Сохраняем найденные категории для диалога
  const cache = CacheService.getUserCache();
  cache.put('foundCategories', JSON.stringify(foundCategories), 300);
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          h3 { margin-top: 0; color: #333; }
          .info { color: #666; margin-bottom: 15px; }
          .category-list { 
            max-height: 400px; 
            overflow-y: auto; 
            border: 1px solid #ddd; 
            padding: 10px;
            margin: 10px 0;
            background: #f9f9f9;
          }
          .category-item { 
            padding: 8px 5px; 
            display: flex;
            align-items: center;
            border-bottom: 1px solid #eee;
          }
          .category-item:last-child { border-bottom: none; }
          .category-item:hover { background: #f0f0f0; }
          .category-item input { margin-right: 10px; }
          .category-item label { cursor: pointer; flex: 1; }
          .buttons { margin-top: 20px; text-align: right; }
          button { 
            padding: 10px 20px; 
            margin-left: 10px;
            cursor: pointer;
            border: 1px solid #ddd;
            background: #fff;
            border-radius: 4px;
          }
          button:hover { background: #f5f5f5; }
          .select-all { 
            margin: 10px 0;
            padding: 10px;
            background: #e8f0fe;
            border-radius: 4px;
          }
          .primary { 
            background: #4285f4; 
            color: white; 
            border: none;
          }
          .primary:hover { background: #3367d6; }
        </style>
      </head>
      <body>
        <h3>Выберите категории для загрузки</h3>
        
        <div id="content">
          <div>Загрузка найденных категорий...</div>
        </div>
        
        <script>
          let foundCategories = [];
          
          window.onload = function() {
            google.script.run
              .withSuccessHandler(loadFoundCategories)
              .withFailureHandler(showError)
              .getFoundCategoriesForDialog();
          };
          
          function loadFoundCategories(categories) {
            foundCategories = categories;
            
            const content = document.getElementById('content');
            content.innerHTML = \`
              <p class="info">Найдено категорий: \${categories.length}</p>
              
              <div class="select-all">
                <label>
                  <input type="checkbox" id="selectAll" onchange="toggleAll()">
                  <strong>Выбрать все</strong>
                </label>
              </div>
              
              <div class="category-list">
                \${categories.map((cat, index) => \`
                  <div class="category-item">
                    <input type="checkbox" 
                           id="cat_\${index}" 
                           value="\${cat.id}" 
                           data-title="\${cat.title}">
                    <label for="cat_\${index}">\${cat.title}</label>
                  </div>
                \`).join('')}
              </div>
              
              <div class="buttons">
                <button onclick="google.script.host.close()">Отмена</button>
                <button class="primary" onclick="loadSelectedCategories()">
                  Загрузить выбранные
                </button>
              </div>
            \`;
          }
          
          function showError(error) {
            document.getElementById('content').innerHTML = 
              '<div>Ошибка загрузки категорий: ' + error.message + '</div>';
          }
          
          function toggleAll() {
            const selectAll = document.getElementById('selectAll').checked;
            const checkboxes = document.querySelectorAll('.category-item input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = selectAll);
          }
          
          function loadSelectedCategories() {
            const selected = [];
            const checkboxes = document.querySelectorAll('.category-item input[type="checkbox"]:checked');
            
            checkboxes.forEach(cb => {
              selected.push({
                id: parseInt(cb.value),
                title: cb.getAttribute('data-title')
              });
            });
            
            if (selected.length === 0) {
              alert('Выберите хотя бы одну категорию');
              return;
            }
            
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler((error) => alert('Ошибка: ' + error.message))
              .loadProductsFromSelectedCategories(selected);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(650);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Выбор найденных категорий');
}

// ========================================
// ОТПРАВКА ОБРАБОТАННЫХ ИЗОБРАЖЕНИЙ В INSALES
// ========================================

// Кеш SKU -> Product ID для ускорения повторных поисков
let _skuToProductIdCache = null;

/**
 * Построение кеша SKU -> Product ID
 * Загружает все товары один раз и строит индекс
 */
function buildSkuCache() {
  const context = "Построение кеша SKU";

  if (_skuToProductIdCache !== null) {
    logInfo(`📦 Используем существующий кеш (${Object.keys(_skuToProductIdCache).length} артикулов)`);
    return _skuToProductIdCache;
  }

  logInfo(`🔄 Строим кеш SKU -> Product ID...`);
  _skuToProductIdCache = {};

  let page = 1;
  const perPage = 250; // Максимум для InSales
  let totalProducts = 0;

  while (true) {
    const products = makeInsalesRequest('GET', '/admin/products.json', null, {
      per_page: perPage,
      page: page
    });

    if (!products || products.length === 0) {
      break;
    }

    for (const product of products) {
      if (product.variants && product.variants.length > 0) {
        for (const variant of product.variants) {
          if (variant.sku) {
            const sku = String(variant.sku).trim();
            _skuToProductIdCache[sku] = product.id;
          }
        }
      }
    }

    totalProducts += products.length;

    if (products.length < perPage) {
      break;
    }

    page++;
    Utilities.sleep(200);
  }

  logInfo(`✅ Кеш построен: ${Object.keys(_skuToProductIdCache).length} артикулов из ${totalProducts} товаров`);
  return _skuToProductIdCache;
}

/**
 * Сброс кеша SKU (вызывать после создания новых товаров)
 */
function clearSkuCache() {
  _skuToProductIdCache = null;
  logInfo('🗑️ Кеш SKU очищен');
}

/**
 * Поиск товара в InSales по артикулу (SKU)
 * Использует кеш для быстрого поиска
 * @param {string} article - Артикул товара
 * @returns {number|null} - ID товара в InSales или null если не найден
 */
function findProductIdByArticle(article) {
  const context = "Поиск товара по артикулу в InSales";

  try {
    if (!article) {
      logWarning('⚠️ Артикул не указан для поиска');
      return null;
    }

    const searchArticle = String(article).trim();
    logInfo(`🔍 Ищем товар с артикулом: ${searchArticle}`);

    // Используем кеш для быстрого поиска
    const cache = buildSkuCache();

    if (cache[searchArticle]) {
      const productId = cache[searchArticle];
      logInfo(`✅ Найден товар ID ${productId} с артикулом ${searchArticle} (из кеша)`);
      return productId;
    }

    logInfo(`ℹ️ Товар с артикулом ${searchArticle} не найден в InSales`);
    return null;

  } catch (error) {
    logError(`❌ Ошибка поиска товара по артикулу: ${article}`, error, context);
    return null;
  }
}

/**
 * Отправка выбранных товаров с обработанными изображениями в InSales
 * Читает данные из Google Sheets и обновляет карточки товаров
 */
function sendProcessedImagesToInSales() {
  const context = "Отправка изображений в InSales";
  
  try {
    logInfo('🚀 Начинаем отправку обработанных изображений в InSales', null, context);
    
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    const readyToSend = [];
    
    // ИСПРАВЛЕННАЯ проверка с правильными константами
    for (let i = 1; i < data.length; i++) {
      const checkbox = data[i][IMAGES_COLUMNS.CHECKBOX - 1];
      const processingStatus = data[i][IMAGES_COLUMNS.PROCESSING_STATUS - 1];
      const processedImages = data[i][IMAGES_COLUMNS.PROCESSED_IMAGES - 1];
      const insalesStatus = data[i][IMAGES_COLUMNS.INSALES_STATUS - 1];
      
      // ИСПРАВЛЕНО: сравниваем со строковыми значениями из config
      if (checkbox === true && 
          processingStatus === STATUS_VALUES.PROCESSING.COMPLETED &&  // 'Обработано' 
          processedImages && 
          processedImages.trim() !== '' &&
          insalesStatus !== STATUS_VALUES.INSALES.SENT) {  // != 'Отправлено ✅'
        
        readyToSend.push({
          rowIndex: i + 1,
          article: data[i][IMAGES_COLUMNS.ARTICLE - 1],
          insalesId: data[i][IMAGES_COLUMNS.INSALES_ID - 1],
          productName: data[i][IMAGES_COLUMNS.PRODUCT_NAME - 1],
          processedImages: processedImages,
          altTags: data[i][IMAGES_COLUMNS.ALT_TAGS - 1] || '',
          seoFilenames: data[i][IMAGES_COLUMNS.SEO_FILENAMES - 1] || ''
        });
      }
    }
    
    if (readyToSend.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Нет товаров готовых к отправке.\n\nПроверьте:\n1. Отмечены чекбоксы\n2. Статус "Обработано"\n3. Есть обработанные изображения',
        '⚠️ Нет данных',
        8
      );
      return;
    }
    
    // Подтверждение отправки
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Отправка в InSales',
      `Готово к отправке: ${readyToSend.length} товаров\n\n` +
      'ВНИМАНИЕ: Старые изображения будут заменены!\n' +
      'Alt-теги загружаются автоматически вместе с изображениями.\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    // Обрабатываем каждый товар
    for (let i = 0; i < readyToSend.length; i++) {
      const item = readyToSend[i];
      
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Товар ${i + 1}/${readyToSend.length}: ${item.productName}`,
          '📤 Отправка',
          3
        );
        
        // Обновляем статус
        sheet.getRange(item.rowIndex, IMAGES_COLUMNS.INSALES_STATUS)
             .setValue(STATUS_VALUES.INSALES.SENDING);  // 'Отправка...'
        
        // ИСПРАВЛЕННАЯ подготовка изображений - как в старом проекте
        const imageUrls = [];
        if (item.processedImages) {
          // Разделяем по символу | (как в ваших данных)
        const rawUrls = item.processedImages.split(/[|,\n]/)
            .map(url => url.trim())
            .filter(url => url && url.startsWith('http'));
          
          imageUrls.push(...rawUrls);
        }
        
        // Alt-теги разделяются ТОЛЬКО переносом строки (не запятой - она часть текста!)
        const altTags = item.altTags ?
          item.altTags.split(/[\n]/).map(alt => alt.trim()).filter(alt => alt) :
          [];
        
        if (imageUrls.length === 0) {
          throw new Error('Нет валидных URL изображений');
        }

        // Если InSales ID отсутствует - ищем товар по артикулу
        let productId = item.insalesId;
        if (!productId) {
          logInfo(`⚠️ InSales ID пустой, ищем товар по артикулу: ${item.article}`);
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Поиск товара по артикулу ${item.article}...`,
            '🔍 Поиск в InSales',
            5
          );

          productId = findProductIdByArticle(item.article);

          if (productId) {
            // Сохраняем найденный ID в таблицу для будущих операций
            sheet.getRange(item.rowIndex, IMAGES_COLUMNS.INSALES_ID).setValue(productId);
            logInfo(`✅ Найден и сохранён InSales ID: ${productId}`);
          } else {
            throw new Error(`Товар с артикулом ${item.article} не найден в InSales. Сначала создайте товар.`);
          }
        }

        logInfo(`📤 Товар: ${item.productName} (InSales ID: ${productId})`);
        logInfo(`📷 Изображений: ${imageUrls.length}, Alt-тегов: ${altTags.length}`);

        // Отправляем в InSales (как в старом проекте)
        const seoFilenames = item.seoFilenames ?
          item.seoFilenames.split(/[|,\n]/).map(name => name.trim()).filter(name => name) :
          [];

        const success = updateProductInInSalesWorking(productId, imageUrls, altTags, seoFilenames);
        
        if (success) {
          sheet.getRange(item.rowIndex, IMAGES_COLUMNS.INSALES_STATUS)
               .setValue(STATUS_VALUES.INSALES.SENT);  // 'Отправлено ✅'
          successCount++;
        } else {
          sheet.getRange(item.rowIndex, IMAGES_COLUMNS.INSALES_STATUS)
               .setValue(STATUS_VALUES.INSALES.ERROR);  // 'Ошибка отправки'
          errorCount++;
        }
        
        // Пауза между запросами (как в старом проекте)
        Utilities.sleep(5000);
        
      } catch (error) {
        sheet.getRange(item.rowIndex, IMAGES_COLUMNS.INSALES_STATUS)
             .setValue(STATUS_VALUES.INSALES.ERROR);
        errorCount++;
        logError(`❌ Ошибка товара ${item.insalesId}`, error, context);
      }
    }
    
    const resultMessage = `Отправка завершена!\n\n✅ Успешно: ${successCount}\n❌ Ошибок: ${errorCount}\n\n📷 Alt-теги загружены вместе с изображениями`;
    SpreadsheetApp.getActiveSpreadsheet().toast(resultMessage, '🎉 Готово', 15);
    
    return { success: true, total: readyToSend.length, sent: successCount, errors: errorCount };
    
  } catch (error) {
    logError('❌ Критическая ошибка отправки', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка: ${error.message}`, '❌ Ошибка', 10);
    throw error;
  }
}

/**
 * Обновление товара в InSales через base64 (проверенный метод)
 * Заменяет ВСЕ изображения товара на новые
 */
function updateProductInInSalesWorking(productId, imageUrls, altTags, seoFilenames = []) {
  const context = `Обновление товара ${productId}`;
  
  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    logInfo(`🔄 Обновляем товар ${productId} через base64`, null, context);
    
    const authString = Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`);
    
    // ЭТАП 1: Получаем текущие изображения для удаления
    const getCurrentImages = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + authString,
        'Content-Type': 'application/json'
      }
    };
    
    const currentResponse = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products/${productId}.json`, getCurrentImages);
    
    if (currentResponse.getResponseCode() !== 200) {
      throw new Error(`Не удалось получить данные товара: ${currentResponse.getResponseCode()}`);
    }
    
    const currentProduct = JSON.parse(currentResponse.getContentText());
    const currentImages = currentProduct.images || [];
    
    logInfo(`📊 Текущих изображений: ${currentImages.length}, новых: ${imageUrls.length}`);
    
    // ЭТАП 2: Удаляем старые изображения
    const imagesToDelete = currentImages.map(img => ({
      id: img.id,
      _destroy: true
    }));
    
    // ЭТАП 3: Подготавливаем новые изображения как base64
    const imagesToAdd = [];
    
    for (let index = 0; index < imageUrls.length; index++) {
      const url = imageUrls[index];
      
      try {
        logInfo(`📥 Подготавливаем изображение ${index + 1}/${imageUrls.length}`);
        
        // Проверяем и скачиваем изображение
        logInfo(`📥 Проверяем URL: ${url}`);

        if (!url || !url.startsWith('http')) {
          throw new Error(`Неверный URL: ${url}`);
        }

        logInfo(`📤 Добавляем прямую ссылку: ${url}`);

        // title в InSales API используется как alt-тег на фронтенде
        const altTag = altTags[index] || '';
        // Очищаем имя файла: убираем расширение и точки (InSales добавит расширение сам)
        let filename = seoFilenames[index] || `processed-image-${index + 1}`;
        // Убираем расширение .webp/.jpg и т.д. если есть
        filename = filename.replace(/\.(webp|jpg|jpeg|png|gif)$/i, '');
        // Убираем точки в конце
        filename = filename.replace(/\.+$/, '').trim();

        // Явно добавляем расширение .webp чтобы InSales не добавлял своё
        // (InSales может некорректно добавлять расширение, создавая двойную точку)
        filename = filename + '.webp';

        logInfo(`🔧 SEO filename после очистки: "${filename}"`);

        imagesToAdd.push({
          src: url,
          filename: filename,
          title: altTag,  // InSales использует title как alt-тег
          position: index + 1
        });

        logInfo(`✅ Изображение ${index + 1} готово: ${filename}${altTag ? `, alt: "${altTag.substring(0, 30)}..."` : ''}`);
        
      } catch (imageError) {
        logError(`❌ Ошибка подготовки изображения ${index + 1}`, imageError, context);
        continue;
      }
    }
    
    if (imagesToAdd.length === 0) {
      throw new Error('Не удалось подготовить ни одного изображения');
    }
    
    // ЭТАП 4: Сначала удаляем старые изображения (отдельный запрос)
    if (imagesToDelete.length > 0) {
      logInfo(`🗑️ Удаляем ${imagesToDelete.length} старых изображений...`);

      const deleteData = {
        product: {
          id: productId,
          images_attributes: imagesToDelete
        }
      };

      const deleteResponse = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products/${productId}.json`, {
        method: 'PUT',
        headers: {
          'Authorization': 'Basic ' + authString,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(deleteData),
        muteHttpExceptions: true
      });

      if (deleteResponse.getResponseCode() !== 200) {
        logError(`⚠️ Ошибка удаления старых изображений: ${deleteResponse.getResponseCode()}`, null, context);
        // Продолжаем добавлять новые, даже если удаление не сработало
      } else {
        logInfo(`✅ Старые изображения удалены`);
      }

      Utilities.sleep(2000); // Пауза между запросами
    }

    // ЭТАП 5: Добавляем новые изображения порциями по 3 штуки
    const BATCH_SIZE = 3;
    let addedCount = 0;

    for (let i = 0; i < imagesToAdd.length; i += BATCH_SIZE) {
      const batch = imagesToAdd.slice(i, i + BATCH_SIZE);
      const batchNum = Math.floor(i / BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(imagesToAdd.length / BATCH_SIZE);

      logInfo(`📤 Добавляем изображения: пакет ${batchNum}/${totalBatches} (${batch.length} шт.)`);

      // Дебаг: логируем что отправляем
      batch.forEach((img, idx) => {
        logInfo(`   📎 [${idx}] filename: "${img.filename}"`);
      });

      const addData = {
        product: {
          id: productId,
          images_attributes: batch
        }
      };

      let addResponse;
      try {
        addResponse = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products/${productId}.json`, {
          method: 'PUT',
          headers: {
            'Authorization': 'Basic ' + authString,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(addData),
          muteHttpExceptions: true
        });
      } catch (fetchError) {
        logError(`❌ Ошибка HTTP запроса (пакет ${batchNum})`, fetchError, context);
        continue; // Пробуем следующий пакет
      }

      const responseCode = addResponse.getResponseCode();

      if (responseCode === 200) {
        addedCount += batch.length;
        logInfo(`✅ Пакет ${batchNum} загружен успешно`);
      } else {
        const responseText = addResponse.getContentText();
        logError(`❌ Ошибка пакета ${batchNum}: HTTP ${responseCode}`, null, context);
        logError(`Ответ InSales: ${responseText.substring(0, 300)}`, null, context);
      }

      // Пауза между пакетами (кроме последнего)
      if (i + BATCH_SIZE < imagesToAdd.length) {
        Utilities.sleep(3000);
      }
    }

    if (addedCount > 0) {
      logInfo(`✅ Товар ${productId} обновлен: ${addedCount}/${imagesToAdd.length} изображений загружено`);
      return true;
    } else {
      logError(`❌ Не удалось загрузить ни одного изображения для товара ${productId}`, null, context);
      return false;
    }
    
  } catch (error) {
    logError(`❌ Критическая ошибка обновления товара ${productId}`, error, context);
    return false;
  }
}

/**
 * Синхронное получение учетных данных InSales
 */
function getInsalesCredentialsSync() {
  const apiKey = getSetting('insalesApiKey') || getSetting('InSales_API_Key');
  const password = getSetting('insalesPassword') || getSetting('InSales_Password');
  const shop = getSetting('insalesShop') || getSetting('InSales_Shop');
  
  if (!apiKey || !password || !shop) {
    return null;
  }
  
  return {
    apiKey: apiKey,
    password: password,
    shop: shop,
    baseUrl: `https://${shop}`
  };
}

// ========================================
// ПОМОЩНИК ALT-ТЕГОВ
// ========================================
/**
 * ЗАМЕНА для функции createAltTagCopyHelper в 03_insales_api.txt
 * Улучшенная версия с отслеживанием скопированных тегов
 */
function createAltTagCopyHelper() {
  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    const sentProducts = [];
    
    // Находим отправленные товары с alt-тегами
    for (let i = 1; i < data.length; i++) {
      const insalesStatus = data[i][IMAGES_COLUMNS.INSALES_STATUS - 1];
      const altTags = data[i][IMAGES_COLUMNS.ALT_TAGS - 1];
      const insalesId = data[i][IMAGES_COLUMNS.INSALES_ID - 1];
      const productName = data[i][IMAGES_COLUMNS.PRODUCT_NAME - 1];
      const supplierImages = data[i][IMAGES_COLUMNS.SUPPLIER_IMAGES - 1];

      const isSelected = data[i][IMAGES_COLUMNS.CHECKBOX - 1];
      if (isSelected && insalesStatus === STATUS_VALUES.INSALES.SENT && altTags && 
          (data[i][IMAGES_COLUMNS.ORIGINAL_IMAGES - 1] || supplierImages)) {
        const altTagsArray = altTags.split(/[|\n]/).map(alt => alt.trim()).filter(alt => alt);
        
        sentProducts.push({
          id: insalesId,
          name: productName,
          altTags: altTagsArray
        });
      }
    }
    
    if (sentProducts.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Нет данных',
        'Не найдено отправленных товаров с alt-тегами.\n\nСначала отправьте товары в InSales.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Создаем HTML-помощник с улучшенным отслеживанием
    const htmlContent = createImprovedAltTagHelperHTML(sentProducts);
    
    // Показываем в боковой панели
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setTitle('Помощник копирования Alt-тегов (улучшенный)');

    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    
    logInfo(`✅ Улучшенный HTML-помощник создан для ${sentProducts.length} товаров`);
    
  } catch (error) {
    logError('Ошибка создания помощника alt-тегов', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Ошибка: ${error.message}`,
      '❌ Ошибка',
      5
    );
  }
}

/**
 * Создание улучшенного HTML-контента с отслеживанием скопированных тегов
 */
function createImprovedAltTagHelperHTML(products) {
  let htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
  body { 
    font-family: Arial, sans-serif; 
    margin: 10px; 
    font-size: 12px;
    background: #f9f9f9;
  }
  
  h1 { 
    color: #1976d2; 
    text-align: center;
    margin: 10px 0;
    font-size: 16px;
  }
  
  .instructions { 
    background: #e3f2fd; 
    padding: 8px; 
    margin: 8px 0; 
    border-radius: 4px;
    font-size: 10px;
    border-left: 3px solid #1976d2;
  }
  
  .product { 
    background: white;
    border: 1px solid #ddd; 
    margin: 8px 0; 
    padding: 10px; 
    border-radius: 4px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.1);
  }
  
  .product-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
    padding-bottom: 5px;
    border-bottom: 1px solid #eee;
  }
  
  .product-name { 
    font-weight: bold; 
    color: #333; 
    font-size: 11px;
    flex: 1;
    margin-right: 5px;
  }
  
  .admin-link { 
    background: #4caf50;
    color: white;
    padding: 4px 8px;
    text-decoration: none;
    border-radius: 3px;
    font-size: 9px;
    white-space: nowrap;
  }
  
  .alt-tag { 
    background: #f8f9fa; 
    border: 1px solid #e0e0e0; 
    padding: 6px; 
    margin: 4px 0; 
    border-radius: 3px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    transition: background-color 0.3s;
  }
  
  .alt-tag.copied {
    background: #e8f5e8;
    border-color: #4caf50;
  }
  
  .alt-text {
    flex: 1;
    font-size: 10px;
    color: #333;
    margin-right: 5px;
  }
  
  .copy-btn { 
    background: #1976d2; 
    color: white; 
    border: none; 
    padding: 4px 8px; 
    border-radius: 3px; 
    cursor: pointer;
    font-size: 9px;
    transition: background-color 0.3s;
  }
  
  .copy-btn.copied {
    background: #4caf50;
  }
  
  .copy-btn:hover {
    opacity: 0.9;
  }
  
  .progress-bar {
    background: #e0e0e0;
    height: 20px;
    border-radius: 10px;
    margin: 10px 0;
    overflow: hidden;
  }
  
  .progress-fill {
    background: linear-gradient(90deg, #4caf50, #66bb6a);
    height: 100%;
    width: 0%;
    transition: width 0.5s ease;
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-size: 10px;
    font-weight: bold;
  }
  
  .stats {
    text-align: center;
    margin: 10px 0;
    font-size: 11px;
    color: #666;
  }
  
  .reset-btn {
    background: #ff9800;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 10px;
    margin: 5px;
  }
</style>
</head>
<body>
  <h1>🏷️ Помощник копирования Alt-тегов v2.0</h1>
  
  <div style="text-align: center; margin: 10px 0;">
    <button onclick="openInNewWindow()" style="background: #ff9800; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 11px;">
      📱 Открыть в отдельном окне
    </button>
    <button class="reset-btn" onclick="resetProgress()">
      🔄 Сбросить прогресс
    </button>
  </div>

  <div class="progress-bar">
    <div class="progress-fill" id="progressFill">0%</div>
  </div>
  
  <div class="stats" id="stats">
    <strong>📊 Товаров: ${products.length} | Скопировано: <span id="copiedCount">0</span></strong>
  </div>
  
  <div class="instructions">
    <strong>Быстрая инструкция:</strong> Нажмите "Открыть админку" → Медиа → Кликните на изображение → Копируйте alt-тег → Вставьте в поле описания<br>
    <strong>Новое:</strong> Теги остаются помеченными как скопированные до сброса прогресса
  </div>
`;

  // Добавляем товары
  products.forEach((product, index) => {
    const adminLink = `https://binokl.shop/admin/products/${product.id}`;
    
    htmlContent += `
  <div class="product">
    <div class="product-header">
      <div class="product-name">${product.name}</div>
      <a href="${adminLink}" target="_blank" class="admin-link">
        📝 Открыть в админке InSales
      </a>
    </div>
    
    <div class="alt-tags">`;
    
    product.altTags.forEach((altTag, altIndex) => {
      if (altTag) {
        const tagId = `tag_${index}_${altIndex}`;
        const btnId = `btn_${index}_${altIndex}`;
        
        htmlContent += `
      <div class="alt-tag" id="${tagId}">
        <div class="alt-text">
          <strong>Изображение ${altIndex + 1}:</strong> ${altTag}
        </div>
        <button class="copy-btn" id="${btnId}" onclick="copyToClipboardImproved('${altTag.replace(/'/g, "\\'")}', '${tagId}', '${btnId}')">
          📋 Копировать
        </button>
      </div>`;
      }
    });
    
    htmlContent += `
    </div>
  </div>`;
  });

  // Добавляем улучшенный JavaScript
  htmlContent += `
  
  <div style="text-align: center; margin: 30px 0; color: #666;">
    <p>🎯 <strong>Цель:</strong> Улучшить SEO и доступность товаров через правильные alt-теги</p>
    <p>⏱️ <strong>Время:</strong> ~2-3 минуты на товар</p>
    <p>🆕 <strong>Новое:</strong> Отслеживание прогресса и постоянные отметки о копировании</p>
  </div>

  <script>
    // Хранилище скопированных тегов в localStorage
    let copiedTags = JSON.parse(localStorage.getItem('copiedAltTags') || '{}');
    let totalTags = 0;
    
    // Инициализация при загрузке
    window.onload = function() {
      // Подсчитываем общее количество тегов
      totalTags = document.querySelectorAll('.alt-tag').length;
      
      // Восстанавливаем состояние скопированных тегов
      restoreCopiedState();
      
      // Обновляем прогресс
      updateProgress();
      
      console.log('🏷️ Помощник alt-тегов v2.0 готов к работе!');
      console.log('💡 Теги теперь остаются помеченными как скопированные');
    };
    
    function copyToClipboardImproved(text, tagId, btnId) {
      // Используем современный Clipboard API
      if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(function() {
          markAsCopied(tagId, btnId, text);
        }).catch(function(err) {
          fallbackCopyToClipboard(text, tagId, btnId);
        });
      } else {
        fallbackCopyToClipboard(text, tagId, btnId);
      }
    }
    
    function fallbackCopyToClipboard(text, tagId, btnId) {
      const textArea = document.createElement('textarea');
      textArea.value = text;
      textArea.style.position = 'fixed';
      textArea.style.left = '-999999px';
      textArea.style.top = '-999999px';
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      
      try {
        document.execCommand('copy');
        markAsCopied(tagId, btnId, text);
      } catch (err) {
        alert('Не удалось скопировать. Выделите текст вручную: ' + text);
      }
      
      document.body.removeChild(textArea);
    }
    
    function markAsCopied(tagId, btnId, text) {
      const tagElement = document.getElementById(tagId);
      const btnElement = document.getElementById(btnId);
      
      if (tagElement && btnElement) {
        // Визуально помечаем как скопированный
        tagElement.classList.add('copied');
        btnElement.classList.add('copied');
        btnElement.textContent = '✅ Скопирован!';
        
        // Сохраняем в localStorage
        copiedTags[tagId] = {
          text: text,
          timestamp: Date.now()
        };
        localStorage.setItem('copiedAltTags', JSON.stringify(copiedTags));
        
        // Обновляем прогресс
        updateProgress();
        
        console.log('Alt-тег ПОСТОЯННО помечен как скопированный:', text);
        
        // Кратковременное уведомление
        setTimeout(() => {
          if (btnElement.textContent === '✅ Скопирован!') {
            btnElement.textContent = '✅ Готово';
          }
        }, 2000);
      }
    }
    
    function restoreCopiedState() {
      Object.keys(copiedTags).forEach(tagId => {
        const tagElement = document.getElementById(tagId);
        const btnId = tagId.replace('tag_', 'btn_');
        const btnElement = document.getElementById(btnId);
        
        if (tagElement && btnElement) {
          tagElement.classList.add('copied');
          btnElement.classList.add('copied');
          btnElement.textContent = '✅ Готово';
        }
      });
    }
    
    function updateProgress() {
      const copiedCount = Object.keys(copiedTags).length;
      const percentage = totalTags > 0 ? Math.round((copiedCount / totalTags) * 100) : 0;
      
      const progressFill = document.getElementById('progressFill');
      const copiedCountElement = document.getElementById('copiedCount');
      
      if (progressFill) {
        progressFill.style.width = percentage + '%';
        progressFill.textContent = percentage + '%';
      }
      
      if (copiedCountElement) {
        copiedCountElement.textContent = copiedCount + '/' + totalTags;
      }
      
      // Празднуем завершение
      if (percentage === 100 && copiedCount > 0) {
        setTimeout(() => {
          alert('🎉 Поздравляем! Все alt-теги скопированы!\\n\\nТеперь можно перейти к следующему товару.');
        }, 500);
      }
    }
    
    function resetProgress() {
      if (confirm('Сбросить прогресс копирования?\\n\\nВсе отметки о скопированных тегах будут удалены.')) {
        // Очищаем localStorage
        copiedTags = {};
        localStorage.removeItem('copiedAltTags');
        
        // Сбрасываем визуальное состояние
        document.querySelectorAll('.alt-tag.copied').forEach(tag => {
          tag.classList.remove('copied');
        });
        
        document.querySelectorAll('.copy-btn.copied').forEach(btn => {
          btn.classList.remove('copied');
          btn.textContent = '📋 Копировать';
        });
        
        // Обновляем прогресс
        updateProgress();
        
        console.log('🔄 Прогресс копирования сброшен');
      }
    }
    
    function openInNewWindow() {
      const currentContent = document.documentElement.outerHTML;
      const newWindow = window.open('', '_blank', 'width=450,height=700,scrollbars=yes,resizable=yes');
      newWindow.document.write(currentContent);
      newWindow.document.close();
      newWindow.focus();
    }
  </script>

</body>
</html>`;

  return htmlContent;
}


// ========================================
// ФУНКЦИИ СОЗДАНИЯ ТОВАРОВ В INSALES
// ========================================

/**
 * Проверка существования товара по артикулу (SKU) через лист "Выгрузка"
 *
 * ОПТИМИЗАЦИЯ: Вместо API запросов к InSales (медленно, 1000+ товаров),
 * проверяем наличие артикула в листе "Выгрузка" (колонка X - Артикул).
 * Лист "Выгрузка" содержит актуальный каталог магазина.
 *
 * @param {string} sku - Артикул товара
 * @returns {Object|null} Информация о товаре, если найден, иначе null
 */
function checkProductExists(sku) {
  try {
    const normalizedSku = String(sku).trim();
    logInfo(`🔍 Проверяем существование товара с артикулом: ${normalizedSku}`);

    // Читаем артикулы из листа "Выгрузка" (колонка X = 24)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const exportSheet = ss.getSheetByName(SHEET_NAMES.EXPORT);

    if (!exportSheet) {
      logWarning('⚠️ Лист "Выгрузка" не найден, пропускаем проверку дубликатов');
      return null;
    }

    // Читаем колонки A (ID товара), B (Название) и X (Артикул)
    const lastRow = exportSheet.getLastRow();
    if (lastRow < 2) {
      logInfo('ℹ️ Лист "Выгрузка" пуст');
      return null;
    }

    // Читаем только нужные колонки для экономии памяти
    const skuColumn = exportSheet.getRange(2, 24, lastRow - 1, 1).getValues(); // Колонка X (Артикул)
    const idColumn = exportSheet.getRange(2, 1, lastRow - 1, 1).getValues();   // Колонка A (ID товара)
    const nameColumn = exportSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Колонка B (Название)

    // Ищем артикул
    for (let i = 0; i < skuColumn.length; i++) {
      const existingSku = String(skuColumn[i][0] || '').trim();
      if (existingSku === normalizedSku) {
        const productId = idColumn[i][0];
        const productName = nameColumn[i][0];
        logInfo(`✅ Товар найден в каталоге: "${productName}" (ID: ${productId})`);
        return {
          id: productId,
          title: productName,
          sku: existingSku
        };
      }
    }

    logInfo(`ℹ️ Товар с артикулом "${normalizedSku}" не найден в каталоге`);
    return null;

  } catch (error) {
    handleError(error, `Проверка существования товара ${sku}`);
    return null;
  }
}


/**
 * Создание товара в InSales
 *
 * @param {Object} productData - Данные товара для создания
 * @returns {Object|null} Созданный товар или null при ошибке
 */
function createInsalesProduct(productData) {
  try {
    const {
      categoryId = DEFAULT_CATEGORY_ID,
      collectionIds,  // Массив ID категорий для множественного выбора
      title,
      description,
      shortDescription,
      sku,
      price,
      quantity,
      properties,
      specificationsNormalized,  // Для извлечения веса и габаритов
      isHidden = IMPORT_SETTINGS.CREATE_HIDDEN
    } = productData;

    // Валидация обязательных полей
    if (!title) {
      throw new Error('Отсутствует название товара');
    }

    if (!sku) {
      throw new Error('Отсутствует артикул товара');
    }

    if (!price || price <= 0) {
      throw new Error('Отсутствует или некорректная цена товара');
    }

    // Проверка дубликата по артикулу через лист "Выгрузка"
    const existingProduct = checkProductExists(sku);
    if (existingProduct) {
      logWarning(`⚠️ Товар с артикулом "${sku}" уже существует в каталоге: "${existingProduct.title}" (ID: ${existingProduct.id})`);
      throw new Error(`Товар с артикулом "${sku}" уже существует в InSales (ID: ${existingProduct.id})`);
    }

    logInfo(`📦 Создаем товар в InSales: "${title}" (${sku})`);

    // Формируем payload для API
    // ВАЖНО: поля collection_id и collections_ids НЕ РАБОТАЮТ при создании товара!
    // Категории добавляются ПОСЛЕ создания через API collects
    const payload = {
      product: {
        'title': title,
        'is-hidden': isHidden,
        'available': true
      }
    };

    // Добавляем описания
    if (description) {
      payload.product.description = description;
    }

    if (shortDescription) {
      payload.product['short-description'] = shortDescription;
    }

    // Создаем вариант товара с SKU, ценой, весом и габаритами
    // ВАЖНО: InSales API использует подчёркивания, а не дефисы!
    // Вес и габариты извлекаются из specificationsNormalized через buildVariantAttributes
    const variantData = buildVariantAttributes({
      article: sku,
      price: price,
      stock: quantity,
      specificationsNormalized: specificationsNormalized
    });
    payload.product['variants_attributes'] = [variantData];

    // Добавляем характеристики (properties)
    if (properties && Array.isArray(properties) && properties.length > 0) {
      payload.product['properties_attributes'] = properties.map(prop => ({
        'title': prop.title || prop.name,
        'value': prop.value
      }));
    }

    logInfo('🌐 Отправляем запрос на создание товара', {
      categoryId,
      title,
      sku,
      price,
      quantity,
      propertiesCount: properties ? properties.length : 0,
      payloadKeys: Object.keys(payload.product).join(', ')
    });

    // Выполняем запрос к API
    const response = makeInsalesRequest(
      'POST',
      INSALES_ENDPOINTS.PRODUCTS,
      payload
    );

    // Логируем полный ответ для отладки
    logInfo('📥 Ответ API при создании товара', {
      hasResponse: !!response,
      responseType: typeof response,
      hasId: response ? !!response.id : false,
      responseKeys: response ? Object.keys(response).join(', ') : 'null'
    });

    if (!response || !response.id) {
      logError('❌ API вернул некорректный ответ', {
        response: JSON.stringify(response)
      });
      throw new Error('API не вернул ID созданного товара');
    }

    logInfo(`✅ Товар успешно создан в InSales`, {
      productId: response.id,
      title: response.title,
      permalink: response.permalink,
      categoryId: response.category_id,
      isHidden: response.is_hidden
    });

    // Добавляем товар в выбранные категории (если указаны)
    if (collectionIds && Array.isArray(collectionIds) && collectionIds.length > 0) {
      logInfo(`📂 Добавляем товар в ${collectionIds.length} категорий: ${collectionIds.join(', ')}`);
      const addedCount = addProductToCollections(response.id, collectionIds);

      if (addedCount > 0) {
        logInfo(`✅ Товар добавлен в ${addedCount} категорий`);
      } else {
        logWarning(`⚠️ Не удалось добавить товар в категории`);
      }
    } else if (categoryId) {
      // Если передана одна категория через categoryId (старая логика)
      logInfo(`📂 Добавляем товар в категорию ${categoryId}`);
      addProductToCollection(response.id, categoryId);
    }

    return response;

  } catch (error) {
    handleError(error, 'Создание товара в InSales', {
      title: productData.title,
      sku: productData.sku
    });
    return null;
  }
}


/**
 * Добавляет товар в категорию через API collects
 *
 * @param {number} productId - ID товара в InSales
 * @param {number} collectionId - ID категории (collection)
 * @returns {Object|null} Созданная связь или null при ошибке
 */
function addProductToCollection(productId, collectionId) {
  const context = 'addProductToCollection';

  try {
    logInfo(`📂 Добавляем товар ${productId} в категорию ${collectionId}`, null, context);

    const payload = {
      collect: {
        'product_id': productId,
        'collection_id': collectionId
      }
    };

    const response = makeInsalesRequest(
      'POST',
      INSALES_ENDPOINTS.COLLECTS,
      payload
    );

    if (response && response.id) {
      logInfo(`✅ Товар успешно добавлен в категорию (collect ID: ${response.id})`, null, context);
      return response;
    } else {
      logWarning(`⚠️ Не удалось добавить товар в категорию ${collectionId}`, null, context);
      return null;
    }

  } catch (error) {
    handleError(error, `Добавление товара ${productId} в категорию ${collectionId}`, {
      productId,
      collectionId
    });
    return null;
  }
}


/**
 * Добавляет товар в несколько категорий
 *
 * @param {number} productId - ID товара в InSales
 * @param {Array<number>} collectionIds - Массив ID категорий
 * @returns {number} Количество успешно добавленных связей
 */
function addProductToCollections(productId, collectionIds) {
  const context = 'addProductToCollections';

  try {
    if (!collectionIds || !Array.isArray(collectionIds) || collectionIds.length === 0) {
      logWarning('⚠️ Не указаны категории для добавления товара', null, context);
      return 0;
    }

    logInfo(`📂 Добавляем товар ${productId} в ${collectionIds.length} категорий: ${collectionIds.join(', ')}`, null, context);

    let successCount = 0;

    for (let i = 0; i < collectionIds.length; i++) {
      const collectionId = collectionIds[i];

      // Добавляем товар в категорию
      const collect = addProductToCollection(productId, collectionId);

      if (collect) {
        successCount++;
      }

      // Пауза между запросами для избежания rate limit
      if (i < collectionIds.length - 1) {
        Utilities.sleep(500); // 0.5 секунды между запросами
      }
    }

    logInfo(`✅ Товар добавлен в ${successCount} из ${collectionIds.length} категорий`, null, context);

    return successCount;

  } catch (error) {
    handleError(error, 'Добавление товара в несколько категорий', {
      productId,
      collectionIds
    });
    return 0;
  }
}


/**
 * Добавление изображения к товару
 *
 * @param {number} productId - ID товара в InSales
 * @param {string} imageUrl - URL изображения
 * @param {number} position - Позиция изображения (1 = главное фото)
 * @returns {Object|null} Добавленное изображение или null при ошибке
 */
function addProductImage(productId, imageUrl, position = 1) {
  try {
    if (!imageUrl || !imageUrl.startsWith('http')) {
      logWarning(`⚠️ Некорректный URL изображения: ${imageUrl}`);
      return null;
    }

    logInfo(`🖼️ Добавляем изображение к товару ${productId} (позиция ${position})`);

    const endpoint = INSALES_ENDPOINTS.PRODUCT_IMAGES.replace('{product_id}', productId);

    const payload = {
      image: {
        src: imageUrl,
        position: position
      }
    };

    const response = makeInsalesRequest('POST', endpoint, payload);

    if (response && response.id) {
      logInfo(`✅ Изображение успешно добавлено (ID: ${response.id})`);
      return response;
    } else {
      logWarning(`⚠️ Не удалось добавить изображение к товару ${productId}`);
      return null;
    }

  } catch (error) {
    handleError(error, `Добавление изображения к товару ${productId}`, {
      imageUrl,
      position
    });
    return null;
  }
}


/**
 * Добавление нескольких изображений к товару
 *
 * @param {number} productId - ID товара в InSales
 * @param {Array<string>} imageUrls - Массив URL изображений
 * @returns {number} Количество успешно добавленных изображений
 */
function addProductImages(productId, imageUrls) {
  try {
    if (!imageUrls || !Array.isArray(imageUrls) || imageUrls.length === 0) {
      logWarning('⚠️ Массив изображений пуст');
      return 0;
    }

    // Ограничиваем количество изображений
    const limitedUrls = imageUrls.slice(0, IMPORT_SETTINGS.MAX_IMAGES_PER_PRODUCT);

    logInfo(`📸 Добавляем ${limitedUrls.length} изображений к товару ${productId}`);

    let successCount = 0;

    for (let i = 0; i < limitedUrls.length; i++) {
      const imageUrl = limitedUrls[i];
      const position = i + 1;

      const result = addProductImage(productId, imageUrl, position);

      if (result) {
        successCount++;
      }

      // Пауза между запросами для избежания rate limit
      if (i < limitedUrls.length - 1) {
        Utilities.sleep(IMPORT_SETTINGS.API_DELAY_MS);
      }
    }

    logInfo(`✅ Добавлено ${successCount} из ${limitedUrls.length} изображений`);
    return successCount;

  } catch (error) {
    handleError(error, `Добавление изображений к товару ${productId}`);
    return 0;
  }
}