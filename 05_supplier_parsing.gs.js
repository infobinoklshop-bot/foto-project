/**
 * ========================================
 * МОДУЛЬ 05: ПАРСИНГ ПОСТАВЩИКОВ
 * ========================================
 */

function logError(message, error) {
  console.error(message, error);
  if (error && error.stack) {
    console.error('Stack trace:', error.stack);
  }
}

function logInfo(message) {
  console.log(message);
}

function logWarning(message) {
  console.warn(message);
}

const SUPPLIERS_CONFIG = {
  VEBER: {
    name: 'Veber',
    baseUrl: 'https://veber.ru',
    searchPath: '/search?q=',
    enabled: true,
    parser: 'parseVeberImages'
  },
  GLAZA_4: {
    name: '4glaza',
    baseUrl: 'https://4glaza.ru',
    searchPath: '/search/?q=',
    enabled: true,  // ИЗМЕНИ ЭТО ЗНАЧЕНИЕ
    parser: 'parse4glazaImages'
  },

  LEVENHUK_OPT: {
    name: 'Levenhuk-opt',
    baseUrl: 'https://www.levenhuk-opt.ru',
    searchPath: '/search/?q=',
    enabled: true,
    parser: 'parseLevenhukOptImages'
  },

  QUARTA: {
    name: 'Quarta Hunt',
    baseUrl: 'https://quarta-hunt.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseQuartaImages'
  },
  STURMAN: {
    name: 'Sturman',
    baseUrl: 'https://sturman.ru',
    searchPath: '/search/?q=',
    enabled: true,
    parser: 'parseSturmanImages'
  },
  ZOOMA: {
    name: 'Zooma',
    baseUrl: 'https://www.zooma.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseZoomaImages'
  },
  ARTELV: {
    name: 'Artelv',
    baseUrl: 'https://artelv.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseArtelvImages'
  },
  MIRZRENIYA: {
    name: 'Mir Zreniya',
    baseUrl: 'https://mirzreniya.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseMirzreniyaImages'
  },
  SFH: {
    name: 'SFH',
    baseUrl: 'https://sfh.ltd',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseSfhImages'
  },
  SUNTC: {
    name: 'Suntc',
    baseUrl: 'https://suntc.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseSuntcImages'
  },
  GEARY: {
    name: 'Geary',
    baseUrl: 'https://geary.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseGearyImages'
  },
  OPTIC4U: {
    name: 'Optic4u',
    baseUrl: 'https://www.optic4u.ru',
    searchPath: '/search/?q=',
    enabled: false,
    parser: 'parseOptic4uImages'
  }
};

function showSupplierParsingDialog() {
  try {
    const selectedProducts = getSelectedProductsForParsing();
    
    if (selectedProducts.length === 0) {
      showNotification('Отметьте товары чекбоксами', 'warning');
      return;
    }
    
    const html = HtmlService.createTemplateFromFile('SupplierSelectionDialog');
    html.suppliers = getEnabledSuppliers();
    html.products = selectedProducts;
    
    SpreadsheetApp.getUi().showModalDialog(
      html.evaluate().setWidth(600).setHeight(500),
      'Выбор поставщиков'
    );
  } catch (error) {
    logError('Ошибка диалога', error);
  }
}

function getEnabledSuppliers() {
  return Object.entries(SUPPLIERS_CONFIG)
    .filter(([k, c]) => c.enabled)
    .map(([k, c]) => ({ key: k, name: c.name, baseUrl: c.baseUrl }));
}

function getSelectedProductsForParsing() {
  const sheet = getImagesSheet();
  const data = sheet.getDataRange().getValues();
  const products = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      products.push({
        row: i + 1,
        article: data[i][1],
        name: data[i][3]
      });
    }
  }
  return products;
}

function runSupplierParsing(selectedSuppliers, articlesMap) {
  try {
    logInfo(`Запуск парсинга: ${selectedSuppliers.length} поставщиков`);
   
    const allResults = [];
   
    for (let s = 0; s < selectedSuppliers.length; s++) {
      const supplierKey = selectedSuppliers[s];
      const config = SUPPLIERS_CONFIG[supplierKey];
     
      if (!config || !config.enabled) continue;
     
      for (const productArticle in articlesMap) {
        const articles = articlesMap[productArticle].split(',').map(a => a.trim());
       
        for (let a = 0; a < articles.length; a++) {
          const article = articles[a];
          
          // ИСПРАВЛЕНО: передаем только артикул, парсер сам формирует URL
          let searchQuery = article;
          
          // Для Sturman - используем только артикул без URL
          if (supplierKey === 'STURMAN') {
            searchQuery = article; // Просто артикул - парсер сам формирует правильный URL
          } else {
            // Для других поставщиков - формируем полный URL как раньше
            searchQuery = `${config.baseUrl}${config.searchPath}${encodeURIComponent(article)}`;
          }
         
          try {
            const images = executeSupplierParser(supplierKey, searchQuery);
           
            if (images && images.length > 0) {
              allResults.push({
                productArticle: productArticle,
                supplier: config.name,
                images: images
              });
             
              logInfo(`Найдено ${images.length} изображений для ${productArticle}`);
            }
          } catch (error) {
            logError(`Ошибка парсинга ${config.name}:`, error);
          }
         
          Utilities.sleep(1000);
        }
      }
    }
   
    return allResults;
   
  } catch (error) {
    logError('Критическая ошибка парсинга', error);
    throw error;
  }
}

function executeSupplierParser(supplierKey, searchUrl) {
  const config = SUPPLIERS_CONFIG[supplierKey];
  
  switch(config.parser) {
    case 'parseVeberImages':
      return parseVeberImages(searchUrl);
    case 'parse4glazaImages':
      return parse4glazaImages(searchUrl);
    case 'parseLevenhukOptImages':
      return parseLevenhukOptImages(searchUrl);  
    case 'parseSturmanImages':
      return parseSturmanImages(searchUrl);
    default:
      throw new Error(`Парсер ${config.parser} не реализован`);
  }
}

function parseVeberImages(urlOrArticle) {
  try {
    let productUrl = urlOrArticle;
    
    // Если это не прямой URL - ищем товар
    if (!urlOrArticle.includes('/product/')) {
      const searchUrl = urlOrArticle.includes('search?q=') ? 
        urlOrArticle : `https://veber.ru/search?q=${encodeURIComponent(urlOrArticle)}`;
      
      logInfo(`Ищем товар: ${searchUrl}`);
      
      const searchResponse = UrlFetchApp.fetch(searchUrl, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      if (searchResponse.getResponseCode() === 200) {
        const searchHtml = searchResponse.getContentText();
        const linkMatch = searchHtml.match(/<a[^>]+href="([^"]*\/product\/[^"]*)"[^>]*>/i);
        
        if (linkMatch) {
          productUrl = linkMatch[1].startsWith('/') ? 
            'https://veber.ru' + linkMatch[1] : linkMatch[1];
          logInfo(`Найдена страница: ${productUrl}`);
        } else {
          logWarning('Товар не найден в поиске');
          return [];
        }
      } else {
        return [];
      }
    }
    
    // Теперь парсим найденную страницу товара
    const response = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.getResponseCode() !== 200) {
      return [];
    }
    
    const html = response.getContentText();
    const imageUrls = new Set();
    
    const allHrefMatches = html.match(/href="(\/upload\/iblock\/[^"]+\.jpg)"/gi);
    
    if (allHrefMatches) {
      allHrefMatches.forEach(match => {
        const urlMatch = match.match(/href="([^"]+)"/);
        if (urlMatch) {
          const imageUrl = 'https://veber.ru' + urlMatch[1];
          
          const excludePatterns = [
            'logo', 'banner', 'icon', 'button', 
            'vomz.jpg', 'yukon.jpg', 'levenhuk.jpg', 'bresser', 
            'nikon.jpg', 'pentax.jpg', 'olympus.jpg', 'fuginon.jpg',
            'selestron.jpg', 'meade.jpg', 'komz.jpg', 'micromed.jpg',
            'Alekat.jpg', 'Falke.jpg', 'Brite2.jpg', 'warne.jpg', 'est.jpg',
            'iray.jpg', 'ToupTek.jpg', 'EASTCOLIGHT.jpg'
          ];
          
          if (!excludePatterns.some(p => imageUrl.toLowerCase().includes(p.toLowerCase()))) {
            imageUrls.add(imageUrl);
          }
        }
      });
    }
    
    const uniqueImages = Array.from(imageUrls);
    logInfo(`Найдено товарных изображений: ${uniqueImages.length}`);
    
    return uniqueImages;
    
  } catch (error) {
    logError('Ошибка парсинга Veber', error);
    return [];
  }
}

function parse4glazaImages(urlOrArticle) {
  try {
    let productUrl = urlOrArticle;
    
    // Обрабатываем разные форматы входных данных
    if (urlOrArticle.includes('4glaza.ru')) {
      // URL поиска с закодированным товарным URL
      if (urlOrArticle.includes('/search/?q=') && urlOrArticle.includes('products%2F')) {
        const decodedUrl = decodeURIComponent(urlOrArticle);
        const productMatch = decodedUrl.match(/(https:\/\/4glaza\.ru\/products\/[^\s&]+)/);
        if (productMatch) {
          productUrl = productMatch[1];
          logInfo(`Извлечен URL из поискового запроса: ${productUrl}`);
        }
      }
      // Прямой URL товара или строка с URL
      else if (urlOrArticle.includes('/products/')) {
        const urlMatch = urlOrArticle.match(/(https:\/\/4glaza\.ru\/products\/[^\s]+)/);
        if (urlMatch) {
          productUrl = urlMatch[1];
          logInfo(`Найден URL товара: ${productUrl}`);
        }
      }
    }
    
    // Проверка что у нас есть правильный URL
    if (!productUrl.includes('4glaza.ru/products/')) {
      logWarning(`Неподдерживаемый формат для 4glaza: ${urlOrArticle}`);
      return [];
    }
    
    logInfo(`Парсим страницу 4glaza: ${productUrl}`);
    
    // Загружаем страницу товара
    const response = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.getResponseCode() !== 200) {
      logError(`Ошибка загрузки страницы: ${response.getResponseCode()}`);
      return [];
    }
    
    const html = response.getContentText();
    const imageUrls = new Set();
    
    // Ищем все изображения с названием товара в имени файла
    const allImageMatches = html.match(/(?:href|src|data-src)="([^"]*binocular-levenhuk-labzz[^"]*\.jpg)"/gi);
    
    if (allImageMatches) {
      allImageMatches.forEach(match => {
        const urlMatch = match.match(/(?:href|src|data-src)="([^"]+\.jpg)"/);
        if (urlMatch) {
          let imageUrl = urlMatch[1];
          
          if (imageUrl.startsWith('/')) {
            imageUrl = 'https://4glaza.ru' + imageUrl;
          }
          
          // Исключаем только миниатюры
          if (!imageUrl.includes('_300_') && !imageUrl.includes('_190_')) {
            imageUrls.add(imageUrl);
          }
        }
      });
    }
    
    const uniqueImages = Array.from(imageUrls);
    logInfo(`Найдено изображений 4glaza: ${uniqueImages.length}`);
    
    return uniqueImages;
    
  } catch (error) {
    logError('Ошибка парсинга 4glaza', error);
    return [];
  }
}

function parseLevenhukOptImages(urlOrArticle) {
  try {
    logInfo(`Начинаем парсинг Levenhuk-opt для: ${urlOrArticle}`);
    
    const credentials = {
      login: 'пп058887',
      password: 'LHXLUURM'
    };
    
    // Авторизация
    const loginUrl = 'https://www.levenhuk-opt.ru/login/';
    const loginPayload = {
      'USER_LOGIN': credentials.login,
      'USER_PASSWORD': credentials.password,
      'AUTH_FORM': 'Y',
      'TYPE': 'AUTH'
    };
    
    const loginResponse = UrlFetchApp.fetch(loginUrl, {
      method: 'POST',
      payload: loginPayload,
      followRedirects: false,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    const cookies = loginResponse.getHeaders()['Set-Cookie'] || '';
    const cookieHeader = Array.isArray(cookies) ? cookies.join('; ') : cookies;
    
    logInfo('Авторизация выполнена');
    
    // Поиск товара
    let productUrl = urlOrArticle;
    
    if (!urlOrArticle.includes('levenhuk-opt.ru/catalogue/')) {
      const searchUrl = `https://www.levenhuk-opt.ru/search/?q=${encodeURIComponent(urlOrArticle)}`;
      
      const searchResponse = UrlFetchApp.fetch(searchUrl, {
        muteHttpExceptions: true,
        headers: { 
          'User-Agent': 'Mozilla/5.0',
          'Cookie': cookieHeader
        }
      });
      
      if (searchResponse.getResponseCode() === 200) {
        const searchHtml = searchResponse.getContentText();
        const linkMatch = searchHtml.match(/<a[^>]+href="([^"]*\/catalogue\/[^"]*)"[^>]*>/i);
        
        if (linkMatch) {
          productUrl = linkMatch[1].startsWith('/') ? 
            'https://www.levenhuk-opt.ru' + linkMatch[1] : linkMatch[1];
          logInfo(`Найдена страница товара: ${productUrl}`);
        } else {
          logWarning('Товар не найден в поиске');
          return [];
        }
      }
    }
    
    // Парсинг страницы товара
    const response = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 
        'User-Agent': 'Mozilla/5.0',
        'Cookie': cookieHeader
      }
    });
    
    if (response.getResponseCode() !== 200) {
      return [];
    }
    
    const html = response.getContentText();
    const imageUrls = new Set();
    
    // Ищем изображения в ссылках href (внешние ссылки на 4glaza.ru)
    const externalImageMatches = html.match(/href="(https:\/\/4glaza\.ru\/external\/[^"]+\.jpg)"/gi);
    
    if (externalImageMatches) {
      externalImageMatches.forEach(match => {
        const urlMatch = match.match(/href="([^"]+\.jpg)"/);
        if (urlMatch) {
          imageUrls.add(urlMatch[1]);
        }
      });
    }
    
    // Также ищем локальные изображения
    const localImageMatches = html.match(/(?:src|data-src)="([^"]*\/upload\/[^"]+\.jpg)"/gi);
    
    if (localImageMatches) {
      localImageMatches.forEach(match => {
        const urlMatch = match.match(/(?:src|data-src)="([^"]+\.jpg)"/);
        if (urlMatch) {
          let imageUrl = urlMatch[1];
          if (imageUrl.startsWith('/')) {
            imageUrl = 'https://levenhuk-opt.ru' + imageUrl;
          }
          imageUrls.add(imageUrl);
        }
      });
    }
    
    const uniqueImages = Array.from(imageUrls);
    logInfo(`Найдено изображений Levenhuk-opt: ${uniqueImages.length}`);
    
    return uniqueImages;
    
  } catch (error) {
    logError('Ошибка парсинга Levenhuk-opt', error);
    return [];
  }
}

function parseQuartaImages(url) { return []; }

/**
 * ИСПРАВЛЕННЫЙ parseSturmanImages с точной фильтрацией
 * Замените функцию в 05_supplier_parsing.gs
 */
function parseSturmanImages(articleOrUrl) {
  try {
    // Если передан простой артикул - формируем URL поиска
    let searchUrl;
    
    if (articleOrUrl.includes('http')) {
      searchUrl = articleOrUrl;
    } else {
      searchUrl = `https://sturman.ru/opt/search/?query=${encodeURIComponent(articleOrUrl)}`;
    }
    
    logInfo(`Поиск Sturman: ${searchUrl}`);
    
    // Выполняем поиск
    const searchResponse = UrlFetchApp.fetch(searchUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (searchResponse.getResponseCode() !== 200) {
      logWarning(`Ошибка поиска: ${searchResponse.getResponseCode()}`);
      return [];
    }
    
    const searchHtml = searchResponse.getContentText();
    
    // Ищем первую ссылку на товар
    const productLinkMatch = searchHtml.match(/<a[^>]+href="([^"]*\/product\/[^"]*)"[^>]*>/i);
    
    if (!productLinkMatch) {
      logWarning('Товар не найден в результатах поиска');
      return [];
    }
    
    const productUrl = productLinkMatch[1].startsWith('/') ?
      'https://sturman.ru' + productLinkMatch[1] : productLinkMatch[1];
    
    logInfo(`Найден товар: ${productUrl}`);
    
    // Загружаем страницу товара
    const productResponse = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (productResponse.getResponseCode() !== 200) {
      return [];
    }
    
    const html = productResponse.getContentText();
    const imageUrls = new Set();
    
    // МЕТОД 1: Ищем изображения в href (полноразмерные в галерее)
    // Паттерн: href="/wa-data/public/shop/products/.../images/.../filename.750x0.jpg"
    const hrefMatches = html.match(/href="([^"]*\/wa-data\/public\/shop\/products\/[^"]+\.750x0\.jpg)"/gi);
    
    if (hrefMatches) {
      hrefMatches.forEach(match => {
        const urlMatch = match.match(/href="([^"]+)"/);
        if (urlMatch) {
          let imageUrl = urlMatch[1];
          if (imageUrl.startsWith('/')) {
            imageUrl = 'https://sturman.ru' + imageUrl;
          }
          imageUrls.add(imageUrl);
        }
      });
    }
    
    // МЕТОД 2: Ищем data-fancybox с полноразмерными изображениями
    const fancyboxMatches = html.match(/data-fancybox="gallery"[^>]+href="([^"]*\/wa-data\/public\/shop\/products\/[^"]+\.750x0\.jpg)"/gi);
    
    if (fancyboxMatches) {
      fancyboxMatches.forEach(match => {
        const urlMatch = match.match(/href="([^"]+\.750x0\.jpg)"/);
        if (urlMatch) {
          let imageUrl = urlMatch[1];
          if (imageUrl.startsWith('/')) {
            imageUrl = 'https://sturman.ru' + imageUrl;
          }
          imageUrls.add(imageUrl);
        }
      });
    }
    
    // МЕТОД 3: Ищем через data-image (для слайдеров)
    const dataImageMatches = html.match(/data-image="([^"]*\/wa-data\/public\/shop\/products\/[^"]+\.750x0\.jpg)"/gi);
    
    if (dataImageMatches) {
      dataImageMatches.forEach(match => {
        const urlMatch = match.match(/data-image="([^"]+)"/);
        if (urlMatch) {
          let imageUrl = urlMatch[1];
          if (imageUrl.startsWith('/')) {
            imageUrl = 'https://sturman.ru' + imageUrl;
          }
          imageUrls.add(imageUrl);
        }
      });
    }
    
    // ДОПОЛНИТЕЛЬНАЯ ФИЛЬТРАЦИЯ: оставляем только товарные изображения
    const filteredImages = Array.from(imageUrls).filter(url => {
      // Включаем только изображения с правильным разрешением
      if (!url.includes('.750x0.jpg')) {
        return false;
      }
      
      // Исключаем системные изображения
      const excludePatterns = [
        'logo', 'banner', 'icon', 'button', 'badge',
        'watermark', 'social', 'payment', 'delivery',
        'brand', 'manufacturer', 'certificate'
      ];
      
      const lowerUrl = url.toLowerCase();
      return !excludePatterns.some(pattern => lowerUrl.includes(pattern));
    });
    
    // Сортируем по номеру изображения (если есть)
    const sortedImages = filteredImages.sort((a, b) => {
      const aNum = parseInt((a.match(/\/(\d+)\.750x0\.jpg/) || ['', '0'])[1]);
      const bNum = parseInt((b.match(/\/(\d+)\.750x0\.jpg/) || ['', '0'])[1]);
      return aNum - bNum;
    });
    
    logInfo(`Найдено ${sortedImages.length} изображений для ${articleOrUrl}`);
    
    // Показываем первые несколько для проверки
    if (sortedImages.length > 0) {
      logInfo('Первые изображения:');
      sortedImages.slice(0, 3).forEach((url, i) => {
        logInfo(`${i + 1}. ${url}`);
      });
    }
    
    return sortedImages;
    
  } catch (error) {
    logError('Ошибка парсинга Sturman', error);
    return [];
  }
}

/**
 * Быстрый тест отфильтрованного парсера
 */
function testFilteredSturman() {
  try {
    logInfo('🧪 Тестируем отфильтрованный парсер Sturman');
    
    const images = parseSturmanImages('3252');
    
    logInfo(`✅ Результат: ${images.length} отфильтрованных изображений`);
    
    images.forEach((url, i) => {
      logInfo(`${i + 1}. ${url}`);
    });
    
    return images;
    
  } catch (error) {
    logError('Ошибка теста', error);
    return [];
  }
}

/**
 * Обновленная конфигурация для включения Sturman
 * Замените соответствующую секцию в SUPPLIERS_CONFIG
 */
const STURMAN_CONFIG = {
  name: 'Sturman',
  baseUrl: 'https://sturman.ru',
  searchPath: '/opt/search/?query=',
  enabled: true, // ВКЛЮЧАЕМ парсер
  parser: 'parseSturmanImages'
};

/**
 * Тестовая функция для проверки парсера Sturman
 */
function testSturmanParser() {
  try {
    logInfo('🧪 Тестируем парсер Sturman');
    
    // Тест 1: Прямой URL товара
    const testUrl1 = 'https://sturman.ru/product/tsifrovoy-binokl-sturman-6-36x50-b-pro/';
    logInfo(`Тест 1 - прямой URL: ${testUrl1}`);
    
    const images1 = parseSturmanImages(testUrl1);
    logInfo(`Результат теста 1: ${images1.length} изображений`);
    
    // Тест 2: Поиск по артикулу
    const testArticle = '8888';
    logInfo(`Тест 2 - поиск по артикулу: ${testArticle}`);
    
    const images2 = parseSturmanImages(testArticle);
    logInfo(`Результат теста 2: ${images2.length} изображений`);
    
    // Тест 3: Другой товар
    const testUrl3 = 'https://sturman.ru/product/shtativ-slik-gx-m-compact/';
    logInfo(`Тест 3 - другой товар: ${testUrl3}`);
    
    const images3 = parseSturmanImages(testUrl3);
    logInfo(`Результат теста 3: ${images3.length} изображений`);
    
    const totalImages = images1.length + images2.length + images3.length;
    logInfo(`✅ Тестирование завершено. Всего найдено: ${totalImages} изображений`);
    
    return {
      test1: { url: testUrl1, count: images1.length, images: images1 },
      test2: { query: testArticle, count: images2.length, images: images2 },
      test3: { url: testUrl3, count: images3.length, images: images3 },
      total: totalImages
    };
    
  } catch (error) {
    logError('Ошибка тестирования парсера Sturman', error);
    return { error: error.message };
  }
}

function parseZoomaImages(url) { return []; }
function parseArtelvImages(url) { return []; }
function parseMirzreniyaImages(url) { return []; }
function parseSfhImages(url) { return []; }
function parseSuntcImages(url) { return []; }
function parseGearyImages(url) { return []; }
function parseOptic4uImages(url) { return []; }

function runSupplierParsingWrapper(selectedSuppliers, articlesMap) {
  try {
    showNotification('Запуск парсинга...', 'info');
    
    const results = runSupplierParsing(selectedSuppliers, articlesMap);
    
    if (results && results.length > 0) {
      showImagePreviewDialog(results);
    } else {
      showNotification('Изображения не найдены', 'warning');
    }
    
  } catch (error) {
    logError('Ошибка wrapper парсинга', error);
    throw error;
  }
}

function showImagePreviewDialog(results) {
  const html = HtmlService.createTemplateFromFile('ImagePreviewDialog');
  html.results = results;
  
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(800).setHeight(600),
    'Предпросмотр найденных изображений'
  );
}

function saveSelectedImages(selections) {
  try {
    logInfo('Сохраняем выбранные изображения', selections);
    
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    
    let savedCount = 0;
    
    for (const [article, images] of Object.entries(selections)) {
      logInfo(`Обрабатываем ${article}: ${images.length} изображений`);
      
      for (let i = 1; i < data.length; i++) {
        const rowArticle = String(data[i][IMAGES_COLUMNS.ARTICLE - 1]).trim();
        
        if (rowArticle === String(article).trim()) {
          const existing = data[i][IMAGES_COLUMNS.SUPPLIER_IMAGES - 1] || '';
          const newImages = images.join('\n');
          
          // ТОЛЬКО УБИРАЕМ ПРЕФИКС - остальная логика как была
          const combined = existing ? `${existing}\n${newImages}` : newImages;
          
          sheet.getRange(i + 1, IMAGES_COLUMNS.SUPPLIER_IMAGES).setValue(combined);
          
          logInfo(`Сохранено в строку ${i + 1}: ${images.length} изображений`);
          savedCount++;
          break;
        }
      }
    }
    
    // Подсчитываем общее количество изображений
    let totalImages = 0;
    for (const images of Object.values(selections)) {
      totalImages += images.length;
    }

    showNotification(
      `Сохранено: ${savedCount} товаров, ${totalImages} изображений`,
      'success'
    );
    
  } catch (error) {
    logError('Ошибка сохранения', error);
    throw error;
  }
}
