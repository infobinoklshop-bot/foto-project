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

// =============================================================================
// 🆕 ПАРСИНГ ПОЛНЫХ КАРТОЧЕК ТОВАРОВ
// =============================================================================

/**
 * ПАРСИНГ ПОЛНОЙ КАРТОЧКИ VEBER.RU
 */
function parseVeberFullProduct(articleOrUrl) {
  try {
    logInfo(`🔍 Парсим полную карточку Veber: ${articleOrUrl}`);

    let productUrl = articleOrUrl;

    // Если передан артикул, ищем товар
    if (!articleOrUrl.includes('http')) {
      const searchUrl = `https://veber.ru/search?q=${encodeURIComponent(articleOrUrl)}`;
      const searchHtml = UrlFetchApp.fetch(searchUrl, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      }).getContentText();

      // Извлекаем ссылку на товар
      const linkMatch = searchHtml.match(/href="(\/product\/[^"]+)"/i);
      if (!linkMatch) {
        logWarning(`⚠️ Товар ${articleOrUrl} не найден на Veber`);
        return null;
      }
      productUrl = 'https://veber.ru' + linkMatch[1];
    }

    logInfo(`📄 Загружаем страницу: ${productUrl}`);
    const html = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    }).getContentText();

    // Извлекаем данные
    const result = {
      url: productUrl,
      images: extractVeberImages(html),
      title: extractVeberTitle(html),
      description: extractVeberDescription(html),
      specifications: extractVeberSpecifications(html),
      price: extractVeberPrice(html),
      stock: extractVeberStock(html),
      brand: extractVeberBrand(html),
      warranty: extractVeberWarranty(html)  // 🆕 Гарантия (отдельное поле на Veber.ru)
    };

    logInfo(`✅ Спарсено: ${result.title}`);
    return result;

  } catch (error) {
    logError('Ошибка парсинга Veber', error);
    return null;
  }
}

function extractVeberImages(html) {
  const images = new Set();
  const matches = html.matchAll(/href="(\/upload\/iblock\/[^"]+\.jpg)"/gi);
  for (const match of matches) {
    const url = 'https://veber.ru' + match[1];
    if (!url.match(/logo|banner|icon|button/i)) {
      images.add(url);
    }
  }
  return Array.from(images).join('\n');
}

function extractVeberTitle(html) {
  const match = html.match(/<h1[^>]*>(.*?)<\/h1>/is);
  return match ? cleanHtml(match[1]).trim() : '';
}

function extractVeberDescription(html) {
  // Ищем блок описания
  let desc = html.match(/<div[^>]*class="[^"]*description[^"]*"[^>]*>(.*?)<\/div>/is);
  if (!desc) desc = html.match(/<div[^>]*itemprop="description"[^>]*>(.*?)<\/div>/is);
  return desc ? cleanHtml(desc[1]).trim() : '';
}

function extractVeberSpecifications(html) {
  const specs = {};

  // Парсим таблицу характеристик из вкладки v-specifications
  // Сначала пробуем найти таблицу внутри div с id="v-specifications"
  let tableMatch = html.match(/<div[^>]*id="v-specifications"[^>]*>.*?<table[^>]*>(.*?)<\/table>/is);

  // Если не нашли, пробуем просто найти таблицу с классом charact (старый формат)
  if (!tableMatch) {
    tableMatch = html.match(/<table[^>]*class="[^"]*charact[^"]*"[^>]*>(.*?)<\/table>/is);
  }

  // Если не нашли, ищем любую таблицу внутри tab-pane
  if (!tableMatch) {
    tableMatch = html.match(/<div[^>]*class="[^"]*tab-pane[^"]*"[^>]*>.*?<table[^>]*>(.*?)<\/table>/is);
  }

  if (tableMatch) {
    const rows = tableMatch[1].matchAll(/<tr[^>]*>.*?<td[^>]*>(.*?)<\/td>.*?<td[^>]*>(.*?)<\/td>.*?<\/tr>/gis);
    for (const row of rows) {
      const key = cleanHtml(row[1]).trim();
      const value = cleanHtml(row[2]).trim();
      if (key && value) specs[key] = value;
    }
  }

  return JSON.stringify(specs);
}

function extractVeberPrice(html) {
  // ✅ ИСПРАВЛЕНИЕ: Избегаем ложного срабатывания на priceRange в Schema.org
  // Сначала вырезаем Schema.org JSON-LD блок, чтобы он не мешал

  let cleanHtml = html;

  // Удаляем Schema.org script блоки
  cleanHtml = cleanHtml.replace(/<script[^>]*type="application\/ld\+json"[^>]*>.*?<\/script>/gis, '');

  // 1. Ищем первое число перед "руб" (это актуальная цена с учётом скидки)
  const priceWithSpaces = cleanHtml.match(/(\d+(?:\s+\d+)*)\s*руб/i);
  if (priceWithSpaces) {
    const price = priceWithSpaces[1].replace(/\s+/g, '');
    // Защита от priceRange: если цена > 100000, скорее всего это не цена товара
    if (parseInt(price) < 100000) {
      return price;
    }
  }

  // 2. Fallback: data-price атрибут
  const dataPrice = html.match(/data-price="(\d+)"/i);
  if (dataPrice) return dataPrice[1];

  // 3. Fallback: meta itemprop (может быть старая цена)
  const metaPrice = html.match(/<meta[^>]*itemprop="price"[^>]*content="(\d+)"/i);
  if (metaPrice) return metaPrice[1];

  // 4. Fallback: class="price"
  const classPrice = html.match(/<span[^>]*class="[^"]*price[^"]*"[^>]*>(\d+)/i);
  if (classPrice) return classPrice[1];

  return '';
}

function extractVeberStock(html) {
  // ✅ ИСПРАВЛЕНИЕ: Определяем наличие по кнопкам действий
  // Если есть кнопка "Купить" → товар в наличии → 5 шт
  // Если есть "Сообщить о поступлении" → товара нет → 0 шт

  // 1. Проверяем наличие кнопки "Купить" (товар в наличии)
  if (html.match(/купить|добавить в корзину|buy now/i)) {
    return 'В наличии';
  }

  // 2. Проверяем "Сообщить о поступлении" (товара нет)
  if (html.match(/сообщить о поступлении|notify.*available/i)) {
    return 'Нет в наличии';
  }

  // 3. Проверяем "под заказ"
  if (html.match(/под заказ/i)) {
    return 'Под заказ';
  }

  // 4. По умолчанию
  return 'Уточняйте';
}

function extractVeberBrand(html) {
  const match = html.match(/<span[^>]*itemprop="brand"[^>]*>(.*?)<\/span>/is) ||
                html.match(/Бренд:[^<]*<[^>]*>([^<]+)</is);
  return match ? cleanHtml(match[1]).trim() : '';
}

/**
 * 🆕 НОВАЯ ФУНКЦИЯ: Извлечение гарантии с Veber.ru
 *
 * Veber.ru отображает гарантию не в таблице характеристик,
 * а как отдельный элемент на странице с классом "unlimited-warranty"
 * или ссылкой на /unlimited-warranty/
 *
 * Примеры:
 * - "бессрочная ГАРАНТИЯ" → "Бессрочная"
 * - "Гарантия: 2 года" (в таблице) → "2 года"
 *
 * @param {string} html - HTML страницы товара
 * @returns {string} Значение гарантии
 */
function extractVeberWarranty(html) {
  // 1. Проверяем наличие бессрочной гарантии (отдельный элемент на странице)
  // Ищем ссылку с href на unlimited-warranty или элемент с классом unlimited-warranty
  if (html.match(/class="[^"]*unlimited-warranty[^"]*"|href="[^"]*unlimited-warranty[^"]*"/i)) {
    // Проверяем текст "бессрочная гарантия"
    if (html.match(/бессрочн[^\s<]*\s*(?:<[^>]*>)*\s*гарант/i)) {
      return 'Бессрочная';
    }
  }

  // 2. Ищем гарантию в таблице характеристик
  const tableMatch = html.match(/<tr[^>]*>.*?<td[^>]*>.*?[Гг]арантия.*?<\/td>.*?<td[^>]*>(.*?)<\/td>.*?<\/tr>/is);
  if (tableMatch) {
    return cleanHtml(tableMatch[1]).trim();
  }

  // 3. Ищем паттерн "Гарантия: X лет/года"
  const textMatch = html.match(/[Гг]арантия[:\s]+(\d+\s*(?:год|года|лет|мес|месяц)[а-я]*)/i);
  if (textMatch) {
    return textMatch[1].trim();
  }

  return '';
}

/**
 * ПАРСИНГ ПОЛНОЙ КАРТОЧКИ STURMAN.RU
 */
function parseSturmanFullProduct(articleOrUrl) {
  try {
    logInfo(`🔍 Парсим полную карточку Sturman: ${articleOrUrl}`);

    let productUrl = articleOrUrl;

    // Если передан артикул, ищем товар
    if (!articleOrUrl.includes('http')) {
      const searchUrl = `https://sturman.ru/opt/search/?query=${encodeURIComponent(articleOrUrl)}`;
      const searchHtml = UrlFetchApp.fetch(searchUrl, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      }).getContentText();

      // ИСПРАВЛЕНО: Ищем как /opt/product/, так и /product/ (обычный сайт)
      const linkMatch = searchHtml.match(/href="(\/(?:opt\/)?product\/[^"]+)"/i);
      if (!linkMatch) {
        logWarning(`⚠️ Товар ${articleOrUrl} не найден на Sturman`);
        return null;
      }
      productUrl = 'https://sturman.ru' + linkMatch[1];
    }

    logInfo(`📄 Загружаем страницу: ${productUrl}`);
    const html = UrlFetchApp.fetch(productUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    }).getContentText();

    const result = {
      url: productUrl,
      images: extractSturmanImages(html),
      title: extractSturmanTitle(html),
      description: extractSturmanDescription(html),
      specifications: extractSturmanSpecifications(html),
      price: extractSturmanPrice(html),
      stock: extractSturmanStock(html),
      brand: extractSturmanBrand(html)
    };

    logInfo(`✅ Спарсено: ${result.title}`);
    return result;

  } catch (error) {
    logError('Ошибка парсинга Sturman', error);
    return null;
  }
}

function extractSturmanImages(html) {
  const images = new Set();
  const patterns = [
    /href="([^"]*\/wa-data\/public\/shop\/products\/[^"]+\.750x0\.jpg)"/gi,
    /data-fancybox[^>]*href="([^"]+\.jpg)"/gi,
    /data-image="([^"]+\.jpg)"/gi
  ];

  for (const pattern of patterns) {
    const matches = html.matchAll(pattern);
    for (const match of matches) {
      let url = match[1];
      if (url.startsWith('/')) url = 'https://sturman.ru' + url;
      if (!url.match(/logo|banner|watermark/i)) {
        images.add(url);
      }
    }
  }

  return Array.from(images).join('\n');
}

function extractSturmanTitle(html) {
  const match = html.match(/<h1[^>]*itemprop="name"[^>]*>(.*?)<\/h1>/is) ||
                html.match(/<h1[^>]*>(.*?)<\/h1>/is);
  return match ? cleanHtml(match[1]).trim() : '';
}

function extractSturmanDescription(html) {
  const match = html.match(/<div[^>]*itemprop="description"[^>]*>(.*?)<\/div>/is) ||
                html.match(/<div[^>]*class="[^"]*description[^"]*"[^>]*>(.*?)<\/div>/is);
  return match ? cleanHtml(match[1]).trim() : '';
}

function extractSturmanSpecifications(html) {
  const specs = {};

  // МЕТОД 1: Webasyst структура характеристик (оптовый сайт /opt/)
  const featuresMatch = html.match(/<div[^>]*class="[^"]*wa-product-features[^"]*"[^>]*>(.*?)<\/div>/is);
  if (featuresMatch) {
    const items = featuresMatch[1].matchAll(/<div[^>]*class="feature[^"]*"[^>]*>.*?<span[^>]*class="name"[^>]*>(.*?)<\/span>.*?<span[^>]*class="value"[^>]*>(.*?)<\/span>/gis);
    for (const item of items) {
      const key = cleanHtml(item[1]).trim();
      const value = cleanHtml(item[2]).trim();
      if (key && value) specs[key] = value;
    }
  }

  // МЕТОД 2: Обычный сайт (структура p-f)
  // ИСПРАВЛЕНО: Используем правильный класс p-f вместо p-features-i
  if (Object.keys(specs).length === 0) {
    const rows = html.matchAll(/<div[^>]*class="p-f"[^>]*>\s*<span[^>]*class="p-f-name"[^>]*>(.*?)<\/span>\s*<span[^>]*class="p-f-value"[^>]*>(.*?)<\/span>/gis);

    for (const row of rows) {
      const key = cleanHtml(row[1]).trim();
      const value = cleanHtml(row[2]).trim();

      if (key && value && key.length > 2 && value.length > 0) {
        specs[key] = value;
      }
    }
  }

  return JSON.stringify(specs);
}

function extractSturmanPrice(html) {
  // ИСПРАВЛЕНО: Ищем <meta itemprop="price"> вместо <span>
  const match = html.match(/<meta[^>]*itemprop="price"[^>]*content="(\d+)"/i) ||
                html.match(/<span[^>]*itemprop="price"[^>]*content="(\d+)"/i) ||
                html.match(/<(?:span|div)[^>]*class="[^"]*price[^"]*"[^>]*>(\d+)/i);
  return match ? match[1] : '';
}

function extractSturmanStock(html) {
  if (html.match(/в наличии/i) || html.match(/itemprop="availability"[^>]*>InStock/i)) return 'В наличии';
  return 'Уточняйте';
}

function extractSturmanBrand(html) {
  const match = html.match(/<span[^>]*itemprop="brand"[^>]*>(.*?)<\/span>/is);
  return match ? cleanHtml(match[1]).trim() : '';
}

/**
 * ВСПОМОГАТЕЛЬНАЯ: Очистка HTML
 */
function cleanHtml(text) {
  return text
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();
}
