/**
 * ========================================
 * МОДУЛЬ 06: НОРМАЛИЗАЦИЯ ХАРАКТЕРИСТИК
 * ========================================
 *
 * Преобразует характеристики от поставщиков в унифицированный формат
 * на основе JSON-справочника
 */

// =============================================================================
// УНИВЕРСАЛЬНЫЕ ФУНКЦИИ НОРМАЛИЗАЦИИ
// =============================================================================

/**
 * Глобальное хранилище для ненормализованных значений
 * ✅ ИСПРАВЛЕНИЕ: Теперь использует ScriptProperties для постоянного хранения
 * ✅ НОВОЕ: Хранит артикулы товаров для последующей ре-нормализации
 * Значения не теряются между вызовами функций
 */
const UNNORMALIZED_VALUES = {
  STORAGE_KEY: 'unnormalized_enum_values',

  /**
   * Добавляет ненормализованное значение в список
   * @param {string} parameterName - Название параметра
   * @param {string} rawValue - Ненормализованное значение
   * @param {string} allowedValues - Допустимые значения (enum)
   * @param {string} supplierFieldName - Название поля у поставщика
   * @param {string} article - Артикул товара (опционально)
   */
  add: function (parameterName, rawValue, allowedValues, supplierFieldName, article) {
    const current = this.getAll();

    // ✅ ДИАГНОСТИКА
    logInfo(`🔍 UNNORMALIZED_VALUES.add() вызван:`);
    logInfo(`   - Параметр: "${parameterName}"`);
    logInfo(`   - Значение: "${rawValue}"`);
    logInfo(`   - Допустимые: "${allowedValues}"`);
    logInfo(`   - Поле поставщика: "${supplierFieldName || 'не указано'}"`);
    logInfo(`   - Артикул: "${article || 'не указано'}"`);
    logInfo(`   - Текущий размер списка: ${current.length}`);

    // Проверяем, нет ли уже такой комбинации параметр+значение
    const existingIndex = current.findIndex(item =>
      item.parameter === parameterName &&
      item.value === rawValue
    );

    if (existingIndex === -1) {
      // Добавляем новое значение
      current.push({
        parameter: parameterName,
        value: rawValue,
        allowedValues: allowedValues,
        supplierFieldName: supplierFieldName || '',
        articles: article ? [article] : [],  // ✅ НОВОЕ: массив артикулов
        timestamp: new Date().toISOString()
      });

      logInfo(`   ✅ Добавлено! Новый размер списка: ${current.length}`);
    } else {
      // Значение уже существует - добавляем артикул если его еще нет
      if (article && !current[existingIndex].articles) {
        current[existingIndex].articles = [];
      }

      if (article && !current[existingIndex].articles.includes(article)) {
        current[existingIndex].articles.push(article);
        logInfo(`   ➕ Добавлен артикул "${article}" к существующему значению`);
      } else {
        logInfo(`   ⏭️ Уже существует, пропускаем`);
      }
    }

    // Сохраняем в постоянное хранилище
    PropertiesService.getScriptProperties().setProperty(
      this.STORAGE_KEY,
      JSON.stringify(current)
    );
  },

  /**
   * Возвращает все ненормализованные значения
   */
  getAll: function () {
    try {
      const stored = PropertiesService.getScriptProperties().getProperty(this.STORAGE_KEY);
      return stored ? JSON.parse(stored) : [];
    } catch (e) {
      return [];
    }
  },

  /**
   * Очищает список
   */
  clear: function () {
    PropertiesService.getScriptProperties().deleteProperty(this.STORAGE_KEY);
  },

  /**
   * Проверяет, есть ли ненормализованные значения
   */
  hasValues: function () {
    return this.getAll().length > 0;
  }
};

/**
 * Извлекает первое число из строки
 * @param {string|number} value - Значение для извлечения числа
 * @returns {string} Извлеченное число или исходное значение
 */
function extractNumber(value) {
  const str = String(value);

  // ✅ ИСПРАВЛЕНИЕ: Для диапазонов температур сохраняем полный диапазон
  // Проверяем паттерны: "-43°C ...+55°C", "-43 ...+55", "-43°С …+55°С", "-10 - 40"
  // ✅ ИСПРАВЛЕНИЕ 2: Добавлены пробелы вокруг разделителей (\s* для гибкости)
  if (str.match(/[-−]?\d+(?:[.,]\d+)?.*?(?:\s*[\.…]+\s*|\s+-\s+).*?[+]?\d+(?:[.,]\d+)?/)) {
    // Это диапазон - извлекаем оба числа с знаками
    const matches = str.match(/([-−]?\d+(?:[.,]\d+)?)/g);
    if (matches && matches.length >= 2) {
      // Возвращаем диапазон в формате "-43 ...+55"
      const first = matches[0].replace(',', '.');
      const second = matches[1].replace(',', '.');
      // Добавляем + перед вторым числом если оно положительное и не ноль
      const secondWithSign = (second.startsWith('-') || second.startsWith('−') || second == '0') ? second : `+${second}`;
      return `${first} …${secondWithSign}`;
    }
  }

  // Обычный случай - просто извлекаем первое число (с поддержкой отрицательных)
  const match = str.match(/([-−]?\d+(?:[.,]\d+)?)/);
  return match ? match[1].replace(',', '.') : value;
}

// =============================================================================
// HELPER ФУНКЦИИ ДЛЯ РАБОТЫ С МЕТАДАННЫМИ ПАРАМЕТРОВ
// =============================================================================

/**
 * Парсит ключ параметра с метаданными
 *
 * Формат: "ParameterName@@SupplierField@@SupplierValue"
 * Пример: "Объективы@@Диаметр объектива: 50 мм@@50 мм"
 *
 * Где:
 * - ParameterName - название параметра из справочника
 * - SupplierField - исходное название поля у поставщика
 * - SupplierValue - исходное значение от поставщика
 *
 * Если метаданных нет, возвращает только имя параметра.
 * Это обеспечивает обратную совместимость со старым форматом.
 *
 * @param {string} key - Ключ параметра (с метаданными или без)
 * @returns {{parameterName: string, supplierField: string|null, supplierValue: string|null}}
 */
function parseParameterKey(key) {
  if (!key) {
    return { parameterName: '', supplierField: null, supplierValue: null };
  }

  const match = key.match(/^(.+?)@@(.+?)@@(.+)$/);
  if (match) {
    return {
      parameterName: match[1],
      supplierField: match[2],
      supplierValue: match[3]
    };
  }

  // Обратная совместимость: если нет метаданных, возвращаем только имя
  return {
    parameterName: key,
    supplierField: null,
    supplierValue: null
  };
}

/**
 * Создаёт ключ параметра с метаданными
 *
 * Если указаны supplierField и supplierValue, создаёт ключ с метаданными:
 * "ParameterName@@SupplierField@@SupplierValue"
 *
 * Если нет - возвращает только имя параметра (обратная совместимость).
 *
 * @param {string} parameterName - Название параметра из справочника
 * @param {string} supplierField - Исходное название поля у поставщика (опционально)
 * @param {string} supplierValue - Исходное значение от поставщика (опционально)
 * @returns {string} Ключ параметра
 */
function createParameterKey(parameterName, supplierField, supplierValue) {
  if (!parameterName) {
    return '';
  }

  if (supplierField && supplierValue) {
    return `${parameterName}@@${supplierField}@@${supplierValue}`;
  }

  return parameterName;
}

/**
 * Обновляет значение параметра с сохранением порядка параметров
 *
 * Проблема: Обычное обновление объекта JavaScript помещает изменённый ключ в конец.
 * Решение: Пересоздаём объект, сохраняя порядок всех ключей.
 *
 * @param {Object} specs - Исходный объект с характеристиками
 * @param {string} oldKey - Старый ключ параметра (может быть без метаданных)
 * @param {string} newKey - Новый ключ параметра (обычно с метаданными)
 * @param {*} newValue - Новое значение
 * @returns {Object} Обновлённый объект с сохранённым порядком
 */
function preserveOrderUpdate(specs, oldKey, newKey, newValue) {
  const result = {};

  // Парсим оба ключа для сравнения только по имени параметра
  const oldParsed = parseParameterKey(oldKey);
  const newParsed = parseParameterKey(newKey);

  let found = false;

  // Проходим по всем существующим ключам
  for (const [key, value] of Object.entries(specs)) {
    const parsed = parseParameterKey(key);

    // Если это тот же параметр (по имени), заменяем его
    if (parsed.parameterName === oldParsed.parameterName ||
      parsed.parameterName === newParsed.parameterName) {
      result[newKey] = newValue; // Заменяем на месте
      found = true;
    } else {
      result[key] = value; // Сохраняем как есть
    }
  }

  // Если не нашли - это новый параметр, добавляем в конец
  if (!found) {
    result[newKey] = newValue;
  }

  return result;
}

/**
 * Нормализует строку для сравнения (убирает дефисы, пробелы, приводит к lowercase)
 * @param {string} str - Строка для нормализации
 * @returns {string} Нормализованная строка
 */
function normalizeForComparison(str) {
  return String(str)
    .toLowerCase()
    .replace(/[-\s]/g, '');  // Убираем дефисы и пробелы
}

/**
 * Нормализует enum значение путем поиска ближайшего совпадения
 * из списка допустимых значений
 *
 * @param {string} rawValue - Сырое значение от поставщика
 * @param {string} allowedValuesString - Строка с допустимыми значениями через ";"
 * @param {string} parameterName - Название параметра (для логирования ненормализованных значений)
 * @param {string} supplierFieldName - Название поля у поставщика (опционально)
 * @param {string} article - Артикул товара (опционально)
 * @returns {string} Нормализованное значение из списка допустимых или исходное
 */
function normalizeEnumValue(rawValue, allowedValuesString, parameterName, supplierFieldName, article) {
  if (!rawValue || !allowedValuesString) {
    return rawValue;
  }

  const raw = String(rawValue).trim();
  const allowedValues = allowedValuesString.split(';').map(v => v.trim()).filter(v => v);

  if (allowedValues.length === 0) {
    return rawValue;
  }

  // 1. Точное совпадение (case-insensitive)
  const exactMatch = allowedValues.find(allowed =>
    allowed.toLowerCase() === raw.toLowerCase()
  );
  if (exactMatch) {
    return exactMatch;
  }

  // 2. Совпадение без учета дефисов и пробелов (например, "BAK4" → "BaK-4")
  const rawNormalized = normalizeForComparison(raw);
  const dashesMatch = allowedValues.find(allowed =>
    normalizeForComparison(allowed) === rawNormalized
  );
  if (dashesMatch) {
    return dashesMatch;
  }

  // 2.5. ПРИОРИТЕТ: Поиск по аббревиатуре в скобках (например, "FMC" в "Полное многослойное (FMC)")
  // Извлекаем аббревиатуры из сырого значения (ищем заглавные буквы/цифры, 2+ символов)
  const abbreviationMatches = raw.match(/\b[A-ZА-Я0-9]{2,}\b/g);

  if (abbreviationMatches && abbreviationMatches.length > 0) {
    for (const abbr of abbreviationMatches) {
      // Ищем эту аббревиатуру в скобках в допустимых значениях
      const abbrInParentheses = allowedValues.find(allowed => {
        const parenthesesMatch = allowed.match(/\(([^)]+)\)/);
        if (parenthesesMatch) {
          const textInParentheses = parenthesesMatch[1].toUpperCase();
          return textInParentheses === abbr.toUpperCase();
        }
        return false;
      });

      if (abbrInParentheses) {
        logInfo(`   🎯 Найдено совпадение по аббревиатуре "${abbr}" → "${abbrInParentheses}"`);
        return abbrInParentheses;
      }
    }
  }

  // ✅ ИСПРАВЛЕНИЕ: Шаги 3-4-5 ОТКЛЮЧЕНЫ - они делают слишком агрессивный matching
  // Например, "алюминиево-магниевый сплав" находит "Алюминиевый сплав" через partial match
  // Это неправильно - пользователь должен сам решить, добавлять ли новое значение

  // 3. ОТКЛЮЧЕНО: Поиск вхождения (слишком агрессивно)
  // const containsMatch = allowedValues.find(allowed =>
  //   allowed.toLowerCase().includes(raw.toLowerCase())
  // );
  // if (containsMatch) {
  //   return containsMatch;
  // }

  // 4. ОТКЛЮЧЕНО: Reverse contains (слишком агрессивно)
  // const reverseContainsMatch = allowedValues.find(allowed =>
  //   raw.toLowerCase().includes(allowed.toLowerCase())
  // );
  // if (reverseContainsMatch) {
  //   return reverseContainsMatch;
  // }

  // 5. ОТКЛЮЧЕНО: Fuzzy matching (слишком агрессивно)
  // let bestMatch = null;
  // let bestScore = 0;
  //
  // for (const allowed of allowedValues) {
  //   const score = calculateSimilarity(raw.toLowerCase(), allowed.toLowerCase());
  //   if (score > bestScore) {
  //     bestScore = score;
  //     bestMatch = allowed;
  //   }
  // }
  //
  // // Если схожесть выше 60%, считаем это совпадением
  // if (bestScore > 0.6) {
  //   return bestMatch;
  // }

  // Не найдено совпадение - логируем и возвращаем исходное значение
  logWarning(`⚠️ Не найдено совпадение для enum "${parameterName}": "${rawValue}" (допустимые: ${allowedValuesString})`);

  // ✅ ВАЖНО: Сохраняем для последующего предложения пользователю добавить в справочник
  if (parameterName) {
    UNNORMALIZED_VALUES.add(parameterName, rawValue, allowedValuesString, supplierFieldName, article);
  }

  return rawValue;
}

/**
 * Вычисляет схожесть двух строк (0 до 1)
 * Использует комбинацию метрик для лучшего результата
 *
 * @param {string} str1 - Первая строка
 * @param {string} str2 - Вторая строка
 * @returns {number} Оценка схожести от 0 до 1
 */
function calculateSimilarity(str1, str2) {
  // Используем расстояние Левенштейна
  const distance = levenshteinDistance(str1, str2);
  const maxLength = Math.max(str1.length, str2.length);
  const levenshteinScore = maxLength === 0 ? 1 : 1 - (distance / maxLength);

  // Дополнительная проверка на общие слова
  const words1 = str1.split(/\s+/);
  const words2 = str2.split(/\s+/);
  const commonWords = words1.filter(w => words2.includes(w)).length;
  const wordScore = commonWords / Math.max(words1.length, words2.length);

  // Комбинируем оценки (70% Левенштейн, 30% общие слова)
  return levenshteinScore * 0.7 + wordScore * 0.3;
}

/**
 * Вычисляет расстояние Левенштейна между двумя строками
 *
 * @param {string} str1 - Первая строка
 * @param {string} str2 - Вторая строка
 * @returns {number} Расстояние Левенштейна
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];

  // Инициализация первой строки и столбца
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  // Заполнение матрицы
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // замена
          matrix[i][j - 1] + 1,     // вставка
          matrix[i - 1][j] + 1      // удаление
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

// =============================================================================
// БИБЛИОТЕКА ФУНКЦИЙ НОРМАЛИЗАЦИИ (LEGACY - для обратной совместимости)
// =============================================================================

const NORMALIZER_FUNCTIONS = {
  // Извлечение числа из строки
  // Извлечение числа из строки
  extractNumber: (value) => {
    return extractNumber(value);
  },

  // Нормализация веса (кг -> г)
  normalizeWeight: (value, key) => {
    const str = String(value).toLowerCase();
    const keyStr = String(key || '').toLowerCase();

    // ✅ ИСПРАВЛЕНО: Расширенная проверка на кг в ключе и значении
    // Проверяем: "(кг)", ", кг", "кг)", "кг " и т.д.
    const isKgInKey = /кг|kg/i.test(keyStr);
    const isKgInValue = /(?:\s|^)кг(?:\s|$|[.,)])|(?:\s|^)kg(?:\s|$|[.,)])/i.test(str);
    const isKg = isKgInKey || isKgInValue;

    // Извлекаем число (с запятой или точкой)
    const numMatch = str.match(/(\d+(?:[.,]\d+)?)/);
    if (!numMatch) return str;

    const numStr = numMatch[1];
    const floatVal = parseFloat(numStr.replace(',', '.'));

    if (isKg && !isNaN(floatVal)) {
      // Если это кг, переводим в граммы
      return Math.round(floatVal * 1000);
    }

    // Если в граммах, возвращаем как есть (целое число)
    if (!isNaN(floatVal)) {
      return Math.round(floatVal);
    }

    return numStr;
  },

  // Приведение к верхнему регистру
  toUpperCase: (value) => {
    return String(value).toUpperCase();
  },

  // Первая буква заглавная
  capitalizeFirst: (value) => {
    const str = String(value);
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  },

  // Нормализация типа призмы (только извлекает тип, игнорирует марку стекла)
  normalizePrismType: (value) => {
    const str = String(value);
    // Извлекаем только тип призмы, игнорируя марку стекла
    // Примеры: "Porro BK7" → "PORRO", "Roof BaK-4" → "ROOF"
    if (/porro/i.test(str)) return 'PORRO';
    if (/roof/i.test(str)) return 'ROOF';
    // Если не найден тип призмы, возвращаем пустую строку
    // (значит это только марка стекла без типа призмы)
    return '';
  },

  // Нормализация марки стекла (извлекает марку из комбинированного значения)
  normalizeGlassType: (value) => {
    const str = String(value);
    // Ищем марку стекла в строке (может быть одна или в комбинации с типом призмы)
    if (/bak-?4/i.test(str)) return 'BaK-4';
    if (/bk-?7/i.test(str)) return 'BK-7';
    if (/\bED\b/i.test(str)) return 'ED';
    if (/k-?9/i.test(str)) return 'K-9';
    // Если не найдена марка - возвращаем пустую строку
    return '';
  },

  /**
   * 🆕 НОВАЯ ФУНКЦИЯ: Разделяет комбинированное значение "Porro BK7" на два параметра
   * 
   * Эта функция используется ТОЛЬКО для поля поставщика "Призменная схема",
   * когда оно содержит и тип призмы, и марку стекла в одном значении.
   * 
   * Возвращает объект с двумя параметрами:
   * - "Призменная схема": PORRO/ROOF
   * - "Марка стекла": BaK-4/BK-7/ED/K-9
   * 
   * Примеры:
   * "Porro BK7" → {"Призменная схема": "PORRO", "Марка стекла": "BK-7"}
   * "Roof BaK-4" → {"Призменная схема": "ROOF", "Марка стекла": "BaK-4"}
   * "PORRO" → {"Призменная схема": "PORRO"} (марка не указана)
   * "BK-7" → {"Марка стекла": "BK-7"} (тип призмы не указан)
   * 
   * @param {string} value - Исходное значение от поставщика
   * @returns {Object} Объект с разделенными параметрами
   */
  splitPrismAndGlass: (value) => {
    const result = {};
    const str = String(value);

    // Извлекаем тип призмы
    const prismType = NORMALIZER_FUNCTIONS.normalizePrismType(str);
    if (prismType) {
      result['Призменная схема'] = prismType;
    }

    // Извлекаем марку стекла
    const glassType = NORMALIZER_FUNCTIONS.normalizeGlassType(str);
    if (glassType) {
      result['Марка стекла'] = glassType;
    }

    // 🔍 ДИАГНОСТИКА
    logInfo(`   🔬 splitPrismAndGlass("${value}"):`);
    logInfo(`      → Призменная схема: ${prismType || '(нет)'}`);
    logInfo(`      → Марка стекла: ${glassType || '(нет)'}`);

    return result;
  },

  // Нормализация оптического покрытия
  normalizeCoating: (value) => {
    const str = String(value);
    if (/fmc|полн.*многослойн/i.test(str)) return 'Полное многослойное (FMC)';
    if (/многослойн/i.test(str)) return 'Многослойное';
    if (/однослойн/i.test(str)) return 'Однослойное';
    return str;
  },

  // Нормализация влагозащиты
  normalizeWaterproofing: (value) => {
    const str = String(value).trim().toLowerCase();

    // ✅ ИСПРАВЛЕНИЕ: Обработка значения "да" (поставщик просто указывает наличие защиты)
    if (str === 'да' || str === 'yes' || str === '+') {
      return 'Влагозащищенный';  // По умолчанию считаем базовой защитой
    }

    // ✅ НОВОЕ: Обработка "нет"/"no"/"без защиты" → минимальный уровень защиты
    // Современная оптика всегда имеет хотя бы базовую защиту от брызг
    if (str === 'нет' || str === 'no' || str === 'без защиты' || str === 'none' || str === '-') {
      return 'Влагозащищенный';  // Минимальный уровень защиты (IPX4-IPX5)
    }

    if (/водонепроницаем|ipx7|ipx8|полн.*погруж/i.test(str)) return 'Водонепроницаемый';
    if (/водозащит|ipx6/i.test(str)) return 'Водозащищенный';
    if (/влагозащит|ipx4|ipx5/i.test(str)) return 'Влагозащищенный';
    return value;  // Возвращаем исходное значение (не приведенное к нижнему регистру)
  },

  /**
   * 🆕 НОВАЯ ФУНКЦИЯ: Нормализация наглазников окуляров
   * 
   * Преобразует нестандартные значения от поставщиков ("да", "есть", "+")
   * в стандартное значение из справочника.
   * 
   * Допустимые значения enum (из справочника):
   * - Поворотно-выдвижные
   * - Складные
   * - Резиновые съемные
   * - 4-позиционные
   * - Поворотные
   * 
   * Логика нормализации:
   * - "да" / "yes" / "+" / "есть" → "Поворотно-выдвижные" (наиболее распространенный тип)
   * - "нет" / "no" / "-" → значение не записывается (параметр опционален)
   * - Если значение уже из списка допустимых → возвращается как есть
   * 
   * Примеры:
   * normalizeEyecups("да") → "Поворотно-выдвижные"
   * normalizeEyecups("Складные") → "Складные" (уже корректное)
   * normalizeEyecups("нет") → "" (пустая строка)
   * 
   * @param {string} value - Исходное значение от поставщика
   * @returns {string} Нормализованное значение или пустая строка
   */
  normalizeEyecups: (value) => {
    const str = String(value).trim().toLowerCase();

    // ✅ Обработка "да"/"есть"/"+" → дефолтное значение
    // Поворотно-выдвижные - наиболее распространённый тип в современной оптике
    if (str === 'да' || str === 'yes' || str === '+' || str === 'есть') {
      return 'Поворотно-выдвижные';
    }

    // ✅ Обработка "нет" → параметр не заполняется
    if (str === 'нет' || str === 'no' || str === '-' || str === 'без наглазников') {
      return '';  // Пустое значение - параметр опционален
    }

    // ✅ Проверка на известные типы наглазников
    // Если значение уже корректное - возвращаем исходное (с учетом регистра из value, не str)
    const allowedTypes = [
      'поворотно-выдвижные',
      'складные',
      'резиновые съемные',
      '4-позиционные',
      'поворотные'
    ];

    if (allowedTypes.includes(str)) {
      // Возвращаем исходное значение с сохраненным регистром
      return value.trim();
    }

    // ✅ Частичное совпадение (для вариантов "поворотный", "twist-up" и т.д.)
    if (/поворот.*выдвиж|twist.*up/i.test(str)) return 'Поворотно-выдвижные';
    if (/складн|fold/i.test(str)) return 'Складные';
    if (/резинов.*съе|rubber.*removable/i.test(str)) return 'Резиновые съемные';
    if (/4.*позиц|4.*click/i.test(str)) return '4-позиционные';
    if (/поворотн|twist|rotating/i.test(str)) return 'Поворотные';

    // Не найдено совпадение - возвращаем исходное значение для последующей обработки через enum matching
    return value;
  },

  // Нормализация цвета
  normalizeColor: (value) => {
    const colorMap = {
      'черный': 'Чёрный', 'black': 'Чёрный',
      'зеленый': 'Зелёный', 'green': 'Зелёный',
      'коричневый': 'Коричневый', 'brown': 'Коричневый',
      'камуфляж': 'Камуфляжный', 'camo': 'Камуфляжный',
      'серый': 'Серый', 'gray': 'Серый', 'grey': 'Серый'
    };
    const lower = String(value).toLowerCase();
    return colorMap[lower] || value;
  },

  // Нормализация газового заполнения
  normalizeGas: (value) => {
    const str = String(value);
    if (/азот/i.test(str) || /nitrogen/i.test(str)) return 'Азот';
    return str;
  },

  /**
   * 🆕 НОВАЯ ФУНКЦИЯ: Нормализация гарантии
   *
   * Преобразует значения гарантии в стандартный формат с единицами измерения.
   *
   * Допустимые значения enum (из справочника):
   * - 1 год; 2 года; 3 года; 5 лет; 10 лет; 25 лет; 30 лет
   * - 6 мес.; 12 мес.; 24 месяца
   * - Пожизненная; Бессрочная
   *
   * Логика нормализации:
   * - "1" → "1 год"
   * - "2", "3", "4" → "X года"
   * - "5" и выше → "X лет"
   * - "бессрочная", "unlimited", "lifetime" → "Бессрочная"
   * - "пожизненная" → "Пожизненная"
   *
   * @param {string} value - Исходное значение от поставщика
   * @returns {string} Нормализованное значение с единицами измерения
   */
  normalizeWarranty: (value) => {
    const str = String(value).trim().toLowerCase();

    // ✅ Бессрочная/пожизненная гарантия
    if (/бессрочн|unlimited|lifetime|вечная|навсегда/i.test(str)) {
      return 'Бессрочная';
    }
    if (/пожизнен/i.test(str)) {
      return 'Пожизненная';
    }

    // ✅ Проверяем, есть ли уже единицы измерения
    if (/\d+\s*(год|года|лет|мес|месяц)/i.test(str)) {
      // Уже есть единицы - нормализуем формат
      const match = str.match(/(\d+)\s*(год|года|лет|мес|месяц)/i);
      if (match) {
        const num = parseInt(match[1]);
        const unit = match[2].toLowerCase();

        if (/мес/.test(unit)) {
          return `${num} мес.`;
        }

        // Для лет - правильное склонение
        if (num === 1) return '1 год';
        if (num >= 2 && num <= 4) return `${num} года`;
        return `${num} лет`;
      }
    }

    // ✅ Только число - добавляем единицы (предполагаем годы)
    const numMatch = str.match(/^(\d+)$/);
    if (numMatch) {
      const num = parseInt(numMatch[1]);

      // Числа <= 30 скорее всего годы
      if (num === 1) return '1 год';
      if (num >= 2 && num <= 4) return `${num} года`;
      return `${num} лет`;
    }

    // Возвращаем исходное значение для обработки через enum matching
    return value;
  },

  // Форматирование модели (сохраняет полный формат, заменяет * на x)
  normalizeModel: (value) => {
    const str = String(value).trim();
    // Заменяем * на x для обозначения кратности (7*40 → 7x40)
    return str.replace(/\*/g, 'x');
  },

  // Форматирование диапазона расстояний (сохраняет диапазон, убирает единицы измерения)
  normalizeRange: (value) => {
    const str = String(value).trim();
    // Извлекаем диапазон чисел (например, "8х-24х" → "8-24", "4.15°–3.08°" → "4.15-3.08", "72–54 мм" → "72-54")
    // ✅ ИСПРАВЛЕНО: Теперь поддерживает отрицательные числа (например, "-10 - 40" → "-10…+40")
    // ✅ ИСПРАВЛЕНО 2: Добавлена поддержка многоточия как разделителя ("-43°C ...+55°C")
    // Ищем паттерн: [знак]число [любые символы] разделитель [любые символы] [знак]число
    // Разделители: тире (-, –, —), многоточие (..., …), пробелы с тире
    const match = str.match(/(-?\d+(?:[.,]\d+)?)\s*[^\d\-]*?(?:[-–—]|\.{2,}|…)\s*[^\d\-]*?(\+?-?\d+(?:[.,]\d+)?)/);
    if (match) {
      let first = match[1].replace(',', '.');
      let second = match[2].replace(',', '.');
      // Нормализуем формат: добавляем + перед положительным вторым числом
      const secondWithSign = (second.startsWith('-') || second.startsWith('+')) ? second : `+${second}`;
      return `${first}…${secondWithSign}`;
    }
    // Проверяем формат ±X или -X/+X (для диоптрий)
    const plusMinusMatch = str.match(/[±](\d+(?:[.,]\d+)?)/);
    if (plusMinusMatch) {
      return `±${plusMinusMatch[1]}`;
    }
    const slashMatch = str.match(/(-?\d+(?:[.,]\d+)?)\s*\/\s*(\+?\d+(?:[.,]\d+)?)/);
    if (slashMatch) {
      return `${slashMatch[1]}…${slashMatch[2]}`;
    }
    // Если не диапазон, убираем единицы измерения
    const singleMatch = str.match(/(-?\d+(?:[.,]\d+)?)/);
    return singleMatch ? singleMatch[1] : str;
  },

  // Форматирование габаритов (сохраняет все три размера, убирает единицы измерения)
  normalizeDimensions: (value) => {
    const str = String(value).trim();
    // Извлекаем три числа в формате ДхШхВ (например, "180х110х240 мм" → "180x110x240")
    const match = str.match(/(\d+(?:[.,]\d+)?)\s*[xхXХ×]\s*(\d+(?:[.,]\d+)?)\s*[xхXХ×]\s*(\d+(?:[.,]\d+)?)/);
    if (match) {
      // Убираем единицы измерения из исходного значения - возвращаем только числа
      return `${match[1]}x${match[2]}x${match[3]}`;
    }
    // Если не три размера, возвращаем как есть (возможно 2 размера)
    const twoMatch = str.match(/(\d+(?:[.,]\d+)?)\s*[xхXХ×]\s*(\d+(?:[.,]\d+)?)/);
    if (twoMatch) {
      return `${twoMatch[1]}x${twoMatch[2]}`;
    }
    // Если совсем не найдено чисел, убираем хотя бы единицы измерения
    const singleMatch = str.match(/(\d+(?:[.,]\d+)?)/);
    return singleMatch ? singleMatch[1] : str;
  },

  // Форматирование размера упаковки (сортирует размеры в правильный порядок Д×Ш×В)
  normalizePackageDimensions: (value) => {
    const str = String(value).trim();
    // Извлекаем три числа
    const match = str.match(/(\d+(?:[.,]\d+)?)\s*[xхXХ×]\s*(\d+(?:[.,]\d+)?)\s*[xхXХ×]\s*(\d+(?:[.,]\d+)?)/);
    if (match) {
      // Парсим числа
      const dims = [
        parseFloat(match[1].replace(',', '.')),
        parseFloat(match[2].replace(',', '.')),
        parseFloat(match[3].replace(',', '.'))
      ];

      // Сортируем по убыванию: Длина (наибольшая) × Ширина × Высота (наименьшая)
      dims.sort((a, b) => b - a);

      // Возвращаем в формате Д×Ш×В без единиц измерения
      return `${dims[0]}x${dims[1]}x${dims[2]}`;
    }

    // Если не три размера, используем обычную нормализацию
    return NORMALIZER_FUNCTIONS.normalizeDimensions(value);
  }
};

// =============================================================================
// МАППИНГ ПАРАМЕТРОВ ПОСТАВЩИКОВ (LEGACY - для обратной совместимости)
// =============================================================================

const SPEC_MAPPING = {
  'Параметр: Кратность увеличения, крат': {
    veber: ['Увеличение', 'Кратность', 'Zoom', 'Magnification'],
    sturman: ['Увеличение', 'Кратность', 'Magnification'],
    normalizer: NORMALIZER_FUNCTIONS.extractNumber
  },

  'Параметр: Диаметр объектива, мм': {
    veber: ['Диаметр объектива', 'Апертура', 'Objective', 'Диам. объектива'],
    sturman: ['Объектив', 'Диаметр линзы', 'Диаметр объектива'],
    normalizer: NORMALIZER_FUNCTIONS.extractNumber
  },

  'Параметр: Призменная схема': {
    veber: ['Тип призмы', 'Призмы', 'Prism type', 'Призменная система'],
    sturman: ['Призменная система', 'Призмы', 'Тип призмы'],
    normalizer: NORMALIZER_FUNCTIONS.normalizePrismType
  },

  'Параметр: Марка стекла': {
    veber: ['Стекло призм', 'Материал призм', 'Тип стекла', 'Марка стекла'],
    sturman: ['Материал призм', 'Стекло', 'Марка стекла'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeGlassType
  },

  'Параметр: Оптическое покрытие': {
    veber: ['Покрытие линз', 'Оптическое покрытие', 'Просветление'],
    sturman: ['Покрытие', 'Просветление', 'Оптическое покрытие'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeCoating
  },

  'Параметр: Бренд': {
    veber: ['Бренд', 'Производитель', 'Brand'],
    sturman: ['Бренд', 'Производитель', 'Brand'],
    normalizer: NORMALIZER_FUNCTIONS.capitalizeFirst
  },

  'Параметр: Вес, г': {
    veber: ['Вес', 'Weight', 'Масса'],
    sturman: ['Вес', 'Weight'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeWeight
  },

  'Параметр: Защита от влаги и пыли': {
    veber: ['Влагозащита', 'Водонепроницаемость', 'Защита от влаги'],
    sturman: ['Влагозащита', 'Защита', 'Водонепроницаемость'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeWaterproofing
  },

  'Параметр: Газовое заполнение': {
    veber: ['Заполнение', 'Газ', 'Nitrogen'],
    sturman: ['Заполнение', 'Газовое заполнение'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeGas
  },

  'Параметр: Цвет': {
    veber: ['Цвет', 'Color', 'Окраска'],
    sturman: ['Цвет', 'Color'],
    normalizer: NORMALIZER_FUNCTIONS.normalizeColor
  }
};

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * НОРМАЛИЗАЦИЯ ХАРАКТЕРИСТИК
 *
 * @param {Object} rawSpecs - Сырые характеристики от поставщика
 * @param {string} supplier - Имя поставщика (veber, sturman)
 * @param {string} article - Артикул товара (для связи с ненормализованными значениями)
 * @returns {Object} Нормализованные характеристики
 */
function normalizeSpecifications(rawSpecs, supplier, article) {
  try {
    logInfo(`🔄 Нормализуем характеристики от ${supplier}`);

    // ✅ ИСПРАВЛЕНИЕ: НЕ очищаем список здесь!
    // Очистка должна происходить в начале импорта (в 99_menu.js), а не при каждом вызове нормализации
    // UNNORMALIZED_VALUES.clear();

    if (typeof rawSpecs === 'string') {
      rawSpecs = JSON.parse(rawSpecs);
    }

    // Создаем копию для отслеживания необработанных параметров
    const rawSpecsCopy = Object.assign({}, rawSpecs);

    const normalized = {};
    const unmapped = {};

    // Загружаем справочник параметров из листа
    const reference = loadSpecificationReference();

    if (!reference || reference.length === 0) {
      logWarning('⚠️ Справочник пуст, используем SPEC_MAPPING');
      return normalizeSpecificationsLegacy(rawSpecs, supplier);
    }

    // Проходим по каждому параметру справочника
    for (const refParam of reference) {
      const targetParam = refParam['Параметр (эталонное название)'];
      const synonymsField = supplier === 'veber' ? 'Синонимы Veber' : 'Синонимы Sturman';
      const synonymsStr = refParam[synonymsField] || '';
      let normalizerName = String(refParam['Функция нормализации'] || '').trim();

      // Разбиваем синонимы через точку с запятой
      const synonyms = synonymsStr.split(';').map(s => s.trim()).filter(s => s);

      if (synonyms.length === 0) continue;

      // Ищем значение в сырых данных по синонимам
      let found = false;
      for (const synonym of synonyms) {
        for (const [key, value] of Object.entries(rawSpecsCopy)) {
          if (key.toLowerCase().includes(synonym.toLowerCase())) {
            // ✅ ИСПРАВЛЕНИЕ: Исключаем "Покрытие корпуса" из "Оптическое покрытие"
            if (targetParam === 'Параметр: Оптическое покрытие' && key.toLowerCase().includes('корпус')) {
              continue;
            }

            let normalizedValue = value;

            // ✅ ИСПРАВЛЕНИЕ: Применяем normalizer-функцию ПЕРЕД enum нормализацией
            const fieldType = refParam['Тип поля'] || '';
            const allowedValues = refParam['Допустимые значения (enum)'] || '';

            // 🔍 ДИАГНОСТИКА: Логируем для параметра "Защита от влаги и пыли" и "Габариты"
            if (targetParam.includes('Защита от влаги') || targetParam.includes('абарит')) {
              logInfo(`🔍 ДИАГНОСТИКА параметра "${targetParam}":`);
              logInfo(`   - normalizerName: "${normalizerName}" (тип: ${typeof normalizerName}, длина: ${normalizerName ? normalizerName.length : 0})`);
              logInfo(`   - NORMALIZER_FUNCTIONS[normalizerName] существует: ${normalizerName && NORMALIZER_FUNCTIONS[normalizerName] ? 'ДА' : 'НЕТ'}`);
              logInfo(`   - fieldType: "${fieldType}"`);
              logInfo(`   - allowedValues: "${allowedValues}"`);
              logInfo(`   - Исходное значение: "${value}"`);
            }

            // 1️⃣ ШАГ 1: Сначала применяем custom normalizer (если указан)
            let customNormalizerApplied = false;
            let isMultiParameterNormalizer = false; // 🆕 НОВОЕ: флаг для функций, возвращающих объект

            // ✅ ИСПРАВЛЕНИЕ: Принудительно назначаем нормализаторы для критических параметров,
            // если они не заданы в справочнике или заданы неверно
            if (targetParam.includes('Вес, г')) {
              normalizerName = 'normalizeWeight';
            } else if (targetParam.includes('Оптическое покрытие')) {
              normalizerName = 'normalizeCoating';
            }

            if (normalizerName && NORMALIZER_FUNCTIONS[normalizerName]) {
              try {
                normalizedValue = NORMALIZER_FUNCTIONS[normalizerName](value, key);
                customNormalizerApplied = true;

                // 🆕 НОВОЕ: Проверяем, вернула ли функция объект с несколькими параметрами
                if (typeof normalizedValue === 'object' && !Array.isArray(normalizedValue) && normalizedValue !== null) {
                  isMultiParameterNormalizer = true;
                  logInfo(`   🔬 ${normalizerName} вернул объект с ${Object.keys(normalizedValue).length} параметрами`);

                  // Добавляем все параметры из объекта
                  for (const [paramName, paramValue] of Object.entries(normalizedValue)) {
                    // Находим параметр в справочнике для enum нормализации
                    const paramRef = reference.find(r => r['Параметр (эталонное название)'] === paramName);

                    if (paramRef) {
                      const paramFieldType = paramRef['Тип поля'] || '';
                      const paramAllowedValues = paramRef['Допустимые значения (enum)'] || '';

                      // Применяем enum нормализацию к каждому параметру
                      if (paramFieldType === 'enum' && paramAllowedValues) {
                        normalized[paramName] = normalizeEnumValue(paramValue, paramAllowedValues, paramName, key, article);
                        logInfo(`      → ${paramName}: "${paramValue}" → "${normalized[paramName]}"`);
                      } else {
                        normalized[paramName] = paramValue;
                        logInfo(`      → ${paramName}: "${paramValue}"`);
                      }
                    } else {
                      // Параметр не найден в справочнике - добавляем как есть
                      normalized[paramName] = paramValue;
                      logWarning(`      ⚠️ Параметр "${paramName}" не найден в справочнике`);
                    }
                  }
                } else {
                  logInfo(`   ✨ Предобработка ${normalizerName}: "${value}" → "${normalizedValue}"`);
                }
              } catch (e) {
                logWarning(`⚠️ Ошибка функции ${normalizerName}: ${e.message}`);
              }
            }

            // 🆕 НОВОЕ: Если функция вернула объект - пропускаем дальнейшую обработку
            if (!isMultiParameterNormalizer) {
              // 2️⃣ ШАГ 2: Затем применяем типовую нормализацию на основе типа поля
              if (fieldType === 'enum' && allowedValues) {
                // Для enum - ищем ближайшее совпадение (используем уже предобработанное значение)
                normalizedValue = normalizeEnumValue(normalizedValue, allowedValues, targetParam, key, article);
                logInfo(`   🔄 Enum нормализация: проверка "${normalizedValue}" в (${allowedValues})`);
              } else if ((fieldType === 'число' || fieldType === 'формат') && !customNormalizerApplied) {
                // ✅ ИСПРАВЛЕНИЕ: Пропускаем extractNumber для текстовых полей (Комплектация), даже если в справочнике указано "число"
                if (targetParam.includes('Комплектац') || targetParam.includes('Комплект поставки')) {
                  logInfo(`   🚫 Пропуск extractNumber для "${targetParam}" (текстовое поле)`);
                } else {
                  // ⚠️ ВАЖНО: НЕ применяем extractNumber если уже была применена пользовательская функция (например, normalizeRange)
                  normalizedValue = extractNumber(normalizedValue);
                  logInfo(`   🔢 Число извлечено: "${value}" → "${normalizedValue}"`);
                }
              }

              normalized[targetParam] = normalizedValue;
            }

            delete rawSpecsCopy[key];
            found = true;
            break;
          }
        }
        if (found) break;
      }
    }

    // Остальные параметры сохраняем как unmapped
    for (const [key, value] of Object.entries(rawSpecsCopy)) {
      if (value) unmapped[key] = value;
    }

    logInfo(`✅ Нормализовано: ${Object.keys(normalized).length} параметров`);
    if (Object.keys(unmapped).length > 0) {
      logWarning(`⚠️ Не сопоставлено: ${Object.keys(unmapped).length} параметров`);
      logInfo(`   Несопоставленные параметры: ${Object.keys(unmapped).join(', ')}`);
    }

    // Логируем количество ненормализованных enum значений
    if (UNNORMALIZED_VALUES.hasValues()) {
      const unnormalizedCount = UNNORMALIZED_VALUES.getAll().length;
      logWarning(`⚠️ Найдено ${unnormalizedCount} ненормализованных enum значений`);
      logInfo(`   💡 Рекомендуется добавить их в справочник через меню "📖 Управление справочником параметров → ➕ Добавить ненормализованные значения"`);
    }

    return {
      normalized: normalized,
      unmapped: unmapped
    };

  } catch (error) {
    handleError(error, 'Нормализация характеристик');
    return { normalized: {}, unmapped: rawSpecs };
  }
}

/**
 * Возвращает список ненормализованных enum значений из последней нормализации
 *
 * @returns {Array} Массив объектов с ненормализованными значениями
 */
function getUnnormalizedValues() {
  return UNNORMALIZED_VALUES.getAll();
}

/**
 * Очищает список ненормализованных значений
 */
function clearUnnormalizedValues() {
  UNNORMALIZED_VALUES.clear();
}

/**
 * Форматирует СЫРЫЕ спецификации в порядке нормализованных (для синхронизации колонок P и Q)
 *
 * @param {Object} rawSpecs - Сырые спецификации от поставщика (исходные ключи)
 * @param {Object} normalized - Нормализованные спецификации с mapped параметрами
 * @param {string} supplier - Поставщик (veber/sturman) для поиска синонимов
 * @returns {string} Отформатированная строка в порядке normalized
 */
function formatRawSpecsInNormalizedOrder(rawSpecs, normalized, supplier) {
  try {
    // ✅ ИСПРАВЛЕНИЕ: Возвращаем упорядоченный JSON объект вместо текста
    const orderedSpecs = {};
    const usedRawKeys = new Set();  // Отслеживаем использованные сырые ключи
    const reference = loadSpecificationReference();

    // Для каждого нормализованного параметра ищем соответствующий сырой ключ
    for (const [normalizedKey, normalizedValue] of Object.entries(normalized)) {
      // Ищем параметр в справочнике
      const refParam = reference.find(r => r['Параметр (эталонное название)'] === normalizedKey);

      if (!refParam) {
        // Если параметра нет в справочнике - пропускаем
        continue;
      }

      // Получаем синонимы для этого поставщика
      const synonymsField = supplier === 'veber'
        ? refParam['Синонимы Veber']
        : refParam['Синонимы Sturman'];

      if (!synonymsField) continue;

      const synonyms = synonymsField.split(';').map(s => s.trim()).filter(s => s);

      // Ищем совпадение в rawSpecs по синонимам
      for (const synonym of synonyms) {
        for (const [rawKey, rawValue] of Object.entries(rawSpecs)) {
          if (rawKey.toLowerCase().includes(synonym.toLowerCase())) {
            // Нашли соответствие! Добавляем в упорядоченный объект
            orderedSpecs[rawKey] = rawValue;
            usedRawKeys.add(rawKey);
            break;
          }
        }
        if (usedRawKeys.has(Object.keys(rawSpecs).find(k =>
          k.toLowerCase().includes(synonym.toLowerCase())
        ))) break;
      }
    }

    // Добавляем несопоставленные параметры в конец
    for (const [rawKey, rawValue] of Object.entries(rawSpecs)) {
      if (!usedRawKeys.has(rawKey)) {
        orderedSpecs[rawKey] = rawValue;
      }
    }

    // ✅ Возвращаем JSON строку с упорядоченным объектом
    return JSON.stringify(orderedSpecs);

  } catch (error) {
    logError('Ошибка formatRawSpecsInNormalizedOrder', error);
    // Fallback: просто возвращаем JSON исходных данных
    return JSON.stringify(rawSpecs);
  }
}

/**
 * Форматирует нормализованные характеристики для отображения
 * Преобразует JSON в читаемый многострочный формат
 *
 * @param {Object|string} specifications - Объект спецификаций или JSON строка
 * @param {Object|string} referenceOrder - Эталонный объект для определения порядка параметров (опционально)
 * @returns {string} Форматированная строка, каждый параметр на новой строке
 *
 * @example
 * // Вход: {"Кратность увеличения, крат": "7", "Диаметр объектива, мм": "40"}
 * // Выход:
 * // Кратность увеличения, крат: 7
 * // Диаметр объектива, мм: 40
 */
function formatSpecificationsForDisplay(specifications, referenceOrder) {
  try {
    // Если это строка - парсим JSON
    let specs = specifications;
    if (typeof specifications === 'string') {
      try {
        specs = JSON.parse(specifications);
      } catch (e) {
        // Если не JSON - возвращаем как есть
        return specifications;
      }
    }

    // Если пустой объект - возвращаем пустую строку
    if (!specs || typeof specs !== 'object' || Object.keys(specs).length === 0) {
      return '';
    }

    // Парсим эталонный порядок если передан как строка
    let referenceKeys = [];
    if (referenceOrder) {
      if (typeof referenceOrder === 'string') {
        try {
          const refObj = JSON.parse(referenceOrder);
          referenceKeys = Object.keys(refObj);
        } catch (e) {
          // Игнорируем ошибку парсинга
        }
      } else if (typeof referenceOrder === 'object') {
        referenceKeys = Object.keys(referenceOrder);
      }
    }

    // Формат: "Параметр: значение" на каждой строке
    const lines = [];

    // Если есть эталонный порядок - используем его
    if (referenceKeys.length > 0) {
      // Сначала выводим параметры в порядке эталона
      for (const key of referenceKeys) {
        if (key in specs) {
          const value = specs[key];
          if (value !== null && value !== undefined && value !== '') {
            lines.push(`${key}: ${value}`);
          }
          // ✅ ИСПРАВЛЕНО: Убрали добавление "—" для пустых значений
          // Если параметр в эталоне, но пустой в specs - просто пропускаем
        }
      }

      // Затем добавляем параметры, которых нет в эталоне (если есть)
      for (const [key, value] of Object.entries(specs)) {
        if (!referenceKeys.includes(key)) {
          if (value !== null && value !== undefined && value !== '') {
            lines.push(`${key}: ${value}`);
          }
        }
      }
    } else {
      // Без эталонного порядка - просто перечисляем
      for (const [key, value] of Object.entries(specs)) {
        if (value !== null && value !== undefined && value !== '') {
          lines.push(`${key}: ${value}`);
        }
      }
    }

    return lines.join('\n');

  } catch (error) {
    logError('Ошибка форматирования спецификаций', error);
    return String(specifications);
  }
}

/**
 * НОРМАЛИЗАЦИЯ ХАРАКТЕРИСТИК (LEGACY)
 *
 * Использует старый SPEC_MAPPING для обратной совместимости
 */
function normalizeSpecificationsLegacy(rawSpecs, supplier) {
  try {
    const normalized = {};
    const unmapped = Object.assign({}, rawSpecs);

    // Проходим по маппингу
    for (const [targetParam, config] of Object.entries(SPEC_MAPPING)) {
      const synonyms = config[supplier] || [];

      // Ищем значение в сырых данных
      for (const synonym of synonyms) {
        for (const [key, value] of Object.entries(unmapped)) {
          if (key.toLowerCase().includes(synonym.toLowerCase())) {
            let normalizedValue = value;

            // Применяем нормализатор
            if (config.normalizer) {
              try {
                normalizedValue = config.normalizer(value);
              } catch (e) {
                logWarning(`⚠️ Ошибка нормализации ${targetParam}: ${e.message}`);
              }
            }

            normalized[targetParam] = normalizedValue;
            delete unmapped[key];
            break;
          }
        }
        if (normalized[targetParam]) break;
      }
    }

    logInfo(`✅ Нормализовано (legacy): ${Object.keys(normalized).length} параметров`);
    return { normalized: normalized, unmapped: unmapped };

  } catch (error) {
    handleError(error, 'Legacy нормализация');
    return { normalized: {}, unmapped: rawSpecs };
  }
}

/**
 * ВАЛИДАЦИЯ ПО СПРАВОЧНИКУ
 *
 * @param {Object} specs - Нормализованные характеристики
 * @returns {Object} Результат валидации
 */
function validateAgainstReference(specs) {
  try {
    logInfo('✅ Валидируем характеристики по справочнику');

    // Загружаем справочник
    const reference = loadSpecificationReference();

    const errors = [];
    const warnings = [];

    for (const [param, value] of Object.entries(specs)) {
      const refEntry = reference.find(e =>
        e['Параметр (эталонное название)'] === param
      );

      if (!refEntry) {
        warnings.push(`Параметр "${param}" не найден в справочнике`);
        continue;
      }

      // Проверка обязательности
      if (refEntry['Обязательное поле'] === 'да' && !value) {
        errors.push(`Обязательное поле "${param}" пустое`);
      }

      // Проверка типа
      if (refEntry['Тип поля'] === 'число' && isNaN(value)) {
        errors.push(`"${param}" должно быть числом, получено: ${value}`);
      }

      // Проверка enum
      if (refEntry['Тип поля'] === 'enum') {
        const allowedValues = refEntry['Допустимые значения (enum)']
          .split(';')
          .map(s => s.trim())
          .filter(s => s);

        if (allowedValues.length > 0 && !allowedValues.includes(value)) {
          warnings.push(`"${param}": значение "${value}" не в списке: ${allowedValues.join(', ')}`);
        }
      }
    }

    return {
      isValid: errors.length === 0,
      errors: errors,
      warnings: warnings
    };

  } catch (error) {
    handleError(error, 'Валидация характеристик');
    return { isValid: false, errors: [error.message], warnings: [] };
  }
}

/**
 * ЗАГРУЗКА СПРАВОЧНИКА ИЗ ЛИСТА
 *
 * Читает справочник параметров из служебного листа "Справочник параметров"
 */
function loadSpecificationReference() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    // Если лист не существует, создаем его с базовыми данными
    if (!sheet) {
      logWarning('⚠️ Лист справочника не найден, создаем новый');
      sheet = initializeSpecificationReferenceSheet();
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      logWarning('⚠️ Справочник пуст, возвращаем базовую структуру');
      return BASIC_REFERENCE;
    }

    // Преобразуем данные листа в массив объектов
    const reference = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Пропускаем строки с пустым названием параметра
      if (!row[SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1]) continue;

      // Пропускаем неактивные параметры (чекбокс не установлен)
      const isActive = row[SPEC_REFERENCE_COLUMNS.CHECKBOX - 1];
      if (!isActive) continue;

      // ✅ Убираем префикс "Параметр:" из названия при загрузке
      const parameterName = String(row[SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1] || '').replace(/^Параметр:\s*/i, '');

      reference.push({
        'Параметр (эталонное название)': parameterName,
        'Тип поля': row[SPEC_REFERENCE_COLUMNS.FIELD_TYPE - 1] || '',
        'Допустимые значения (enum)': row[SPEC_REFERENCE_COLUMNS.ALLOWED_VALUES - 1] || '',
        'Обязательное поле': row[SPEC_REFERENCE_COLUMNS.REQUIRED - 1] || '',
        'Синонимы Veber': row[SPEC_REFERENCE_COLUMNS.VEBER_SYNONYMS - 1] || '',
        'Синонимы Sturman': row[SPEC_REFERENCE_COLUMNS.STURMAN_SYNONYMS - 1] || '',
        'Функция нормализации': row[SPEC_REFERENCE_COLUMNS.NORMALIZER_FUNCTION - 1] || '',
        'Описание': row[SPEC_REFERENCE_COLUMNS.DESCRIPTION - 1] || ''
      });
    }

    logInfo(`✅ Загружено ${reference.length} активных параметров из справочника`);
    return reference;

  } catch (error) {
    logError('Ошибка загрузки справочника', error);
    return BASIC_REFERENCE;
  }
}

// Базовый справочник (сокращенная версия для работы без файла)
const BASIC_REFERENCE = [
  {
    'Параметр (эталонное название)': 'Параметр: Кратность увеличения, крат',
    'Тип поля': 'формат',
    'Допустимые значения (enum)': 'Целое число',
    'Обязательное поле': ''
  },
  {
    'Параметр (эталонное название)': 'Параметр: Диаметр объектива, мм',
    'Тип поля': 'формат',
    'Допустимые значения (enum)': 'Целое число',
    'Обязательное поле': ''
  },
  {
    'Параметр (эталонное название)': 'Параметр: Призменная схема',
    'Тип поля': 'enum',
    'Допустимые значения (enum)': 'PORRO; ROOF',
    'Обязательное поле': ''
  },
  {
    'Параметр (эталонное название)': 'Параметр: Марка стекла',
    'Тип поля': 'enum',
    'Допустимые значения (enum)': 'BaK-4; BK-7; ED; K-9',
    'Обязательное поле': ''
  },
  {
    'Параметр (эталонное название)': 'Параметр: Оптическое покрытие',
    'Тип поля': 'enum',
    'Допустимые значения (enum)': 'Однослойное; Многослойное; Полное многослойное (FMC)',
    'Обязательное поле': ''
  }
];

// =============================================================================
// УПРАВЛЕНИЕ СПРАВОЧНИКОМ ПАРАМЕТРОВ
// =============================================================================

/**
 * ПАРСИНГ CSV СПРАВОЧНИКА
 *
 * Парсит CSV-данные справочника параметров
 * @param {string} csvText - Содержимое CSV файла
 * @returns {Array} Массив объектов с параметрами
 */
function parseSpecificationCSV(csvText) {
  try {
    logInfo('📄 Парсим CSV справочник');

    // Удаляем UTF-8 BOM если есть
    csvText = csvText.replace(/^\uFEFF/, '');

    // Разбиваем на строки
    const lines = csvText.split('\n').filter(line => line.trim());

    if (lines.length < 2) {
      throw new Error('CSV файл пуст или содержит только заголовки');
    }

    // Парсим заголовки
    const headers = parseCSVLine(lines[0]);
    logInfo(`  Найдено колонок: ${headers.length}`);

    // Парсим данные
    const data = [];
    for (let i = 1; i < lines.length; i++) {
      const values = parseCSVLine(lines[i]);

      if (values.length === 0 || !values[0]) continue; // Пропускаем пустые строки

      const row = {};
      for (let j = 0; j < headers.length && j < values.length; j++) {
        row[headers[j]] = values[j];
      }

      data.push(row);
    }

    logInfo(`✅ Распарсено ${data.length} параметров из CSV`);
    return data;

  } catch (error) {
    handleError(error, 'Парсинг CSV');
    return [];
  }
}

/**
 * ПАРСИНГ ОДНОЙ СТРОКИ CSV
 *
 * Корректно обрабатывает кавычки и разделители
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      // Двойная кавычка внутри кавычек = экранированная кавычка
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  result.push(current.trim());
  return result;
}

/**
 * ИНИЦИАЛИЗАЦИЯ ЛИСТА СПРАВОЧНИКА
 *
 * Создает служебный лист с параметрами из CSV
 * @param {string} csvText - Опциональный CSV текст (если не передан, создает базовый справочник)
 * @returns {Sheet} Созданный лист
 */
function initializeSpecificationReferenceSheet(csvText = null) {
  try {
    logInfo('🔧 Создаем лист справочника параметров');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Удаляем старый лист если существует
    let sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);
    if (sheet) {
      logWarning('⚠️ Удаляем существующий лист справочника');
      ss.deleteSheet(sheet);
    }

    // Создаем новый лист
    sheet = ss.insertSheet(SHEET_NAMES.SPEC_REFERENCE);

    // Устанавливаем заголовки
    const headers = [
      '✓',                           // A - Чекбокс активности
      'Параметр',                    // B - Эталонное название
      'Тип поля',                    // C - enum/число/текст/формат
      'Допустимые значения',         // D - Через точку с запятой
      'Обязательное',                // E - да/нет
      'Синонимы Veber',              // F - Через точку с запятой
      'Синонимы Sturman',            // G - Через точку с запятой
      'Функция нормализации',        // H - extractNumber/toUpperCase и т.д.
      'Описание'                     // I - Комментарий
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Форматируем заголовки
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setHorizontalAlignment('center');

    // Замораживаем первую строку
    sheet.setFrozenRows(1);

    // Настраиваем ширину колонок
    sheet.setColumnWidth(1, 40);   // Чекбокс
    sheet.setColumnWidth(2, 300);  // Параметр
    sheet.setColumnWidth(3, 100);  // Тип поля
    sheet.setColumnWidth(4, 250);  // Допустимые значения
    sheet.setColumnWidth(5, 100);  // Обязательное
    sheet.setColumnWidth(6, 250);  // Синонимы Veber
    sheet.setColumnWidth(7, 250);  // Синонимы Sturman
    sheet.setColumnWidth(8, 150);  // Функция нормализации
    sheet.setColumnWidth(9, 400);  // Описание

    // Если передан CSV, парсим и заполняем данными
    if (csvText) {
      const csvData = parseSpecificationCSV(csvText);
      populateReferenceSheetFromCSV(sheet, csvData);
    } else {
      // Иначе заполняем базовыми данными из SPEC_MAPPING
      populateReferenceSheetFromMapping(sheet);
    }

    logInfo(`✅ Лист справочника создан: ${sheet.getName()}`);
    return sheet;

  } catch (error) {
    handleError(error, 'Инициализация справочника');
    return null;
  }
}

/**
 * ЗАПОЛНЕНИЕ ЛИСТА ДАННЫМИ ИЗ CSV
 */
function populateReferenceSheetFromCSV(sheet, csvData) {
  try {
    logInfo(`📝 Заполняем справочник из CSV (${csvData.length} параметров)`);

    const rows = [];

    for (const item of csvData) {
      // Извлекаем синонимы Veber и Sturman из комментария или общих синонимов
      const veberSynonyms = extractVeberSynonyms(item);
      const sturmanSynonyms = extractSturmanSynonyms(item);

      // Определяем функцию нормализации на основе типа поля
      const normalizerFunc = determineNormalizerFunction(item);

      rows.push([
        true,                                          // Чекбокс (все активны по умолчанию)
        item['Параметр (эталонное название)'] || '',
        item['Тип поля'] || '',
        item['Допустимые значения (enum)'] || '',
        item['Обязательное поле'] || '',
        veberSynonyms,
        sturmanSynonyms,
        normalizerFunc,
        item['Комментарий'] || ''
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 9).setValues(rows);

      // Устанавливаем чекбоксы в колонке A
      sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();

      logInfo(`✅ Добавлено ${rows.length} параметров`);
    }

  } catch (error) {
    handleError(error, 'Заполнение справочника из CSV');
  }
}

/**
 * ЗАПОЛНЕНИЕ ЛИСТА ИЗ SPEC_MAPPING
 */
function populateReferenceSheetFromMapping(sheet) {
  try {
    logInfo('📝 Заполняем справочник из SPEC_MAPPING');

    const rows = [];

    for (const [paramName, config] of Object.entries(SPEC_MAPPING)) {
      rows.push([
        true,                                           // Чекбокс
        paramName,                                      // Параметр
        determineFieldType(paramName),                  // Тип поля
        '',                                             // Допустимые значения
        '',                                             // Обязательное
        config.veber ? config.veber.join('; ') : '',    // Синонимы Veber
        config.sturman ? config.sturman.join('; ') : '',// Синонимы Sturman
        config.normalizer ? 'custom' : '',              // Функция
        ''                                              // Описание
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 9).setValues(rows);
      sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
      logInfo(`✅ Добавлено ${rows.length} параметров из SPEC_MAPPING`);
    }

  } catch (error) {
    handleError(error, 'Заполнение справочника из маппинга');
  }
}

// =============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ИЗВЛЕЧЕНИЯ ДАННЫХ
// =============================================================================

/**
 * ИЗВЛЕЧЕНИЕ СИНОНИМОВ VEBER ИЗ CSV
 */
function extractVeberSynonyms(csvItem) {
  // Ищем синонимы в комментарии или общих синонимах
  const comment = csvItem['Комментарий'] || '';
  const synonyms = csvItem['Синонимы'] || '';

  // Пытаемся извлечь из комментария паттерн "синонимы: xxx, yyy"
  const match = comment.match(/синонимы?:\s*([^.;]+)/i);
  if (match) {
    return match[1].trim().replace(/,/g, ';');
  }

  // Иначе используем общие синонимы
  if (synonyms) {
    return synonyms.replace(/,/g, ';');
  }

  return '';
}

/**
 * ИЗВЛЕЧЕНИЕ СИНОНИМОВ STURMAN ИЗ CSV
 */
function extractSturmanSynonyms(csvItem) {
  // Для Sturman используем те же синонимы что и для Veber
  // В дальнейшем можно доработать для специфичных синонимов
  return extractVeberSynonyms(csvItem);
}

/**
 * ОПРЕДЕЛЕНИЕ ФУНКЦИИ НОРМАЛИЗАЦИИ
 */
function determineNormalizerFunction(csvItem) {
  const fieldType = csvItem['Тип поля'] || '';
  const paramName = csvItem['Параметр (эталонное название)'] || '';

  // Для чисел - extractNumber
  if (fieldType === 'число' || fieldType === 'формат') {
    if (paramName.includes('мм') || paramName.includes('крат') || paramName.includes('г')) {
      return 'extractNumber';
    }
  }

  // Для enum - может быть toUpperCase или специфичная функция
  if (fieldType === 'enum') {
    if (paramName.includes('Призменная схема')) return 'normalizePrismType';
    if (paramName.includes('Марка стекла')) return 'normalizeGlassType';
    if (paramName.includes('Оптическое покрытие')) return 'normalizeCoating';
    if (paramName.includes('Защита')) return 'normalizeWaterproofing';
    if (paramName.includes('Цвет')) return 'normalizeColor';
    return 'toUpperCase';
  }

  // Для текстовых полей
  if (paramName.includes('Бренд')) return 'capitalizeFirst';

  return '';
}

/**
 * ОПРЕДЕЛЕНИЕ ТИПА ПОЛЯ ПО НАЗВАНИЮ ПАРАМЕТРА
 */
function determineFieldType(paramName) {
  if (paramName.includes('крат') || paramName.includes('мм') || paramName.includes('г')) {
    return 'число';
  }
  if (paramName.includes('схема') || paramName.includes('стекла') || paramName.includes('покрытие')) {
    return 'enum';
  }
  return 'текст';
}

/**
 * ТЕСТ НОРМАЛИЗАЦИИ
 */
function testNormalization() {
  logInfo('🧪 Тестируем нормализацию');

  const testSpecs = {
    'Увеличение': '10x',
    'Диаметр объектива': '42 мм',
    'Призмы': 'Roof',
    'Стекло призм': 'BaK4',
    'Покрытие линз': 'FMC',
    'Бренд': 'VEBER',
    'Вес': '580 г'
  };

  const result = normalizeSpecifications(testSpecs, 'veber');
  logInfo('Результат:', JSON.stringify(result, null, 2));

  const validation = validateAgainstReference(result.normalized);
  logInfo('Валидация:', JSON.stringify(validation, null, 2));
}

/**
 * ТЕСТ СПРАВОЧНИКА ПАРАМЕТРОВ
 *
 * Проверяет работу системы справочника:
 * 1. Создает лист справочника (если не существует)
 * 2. Загружает данные из справочника
 * 3. Нормализует тестовые характеристики
 */
function testSpecificationReference() {
  try {
    logInfo('🧪 ТЕСТИРОВАНИЕ СПРАВОЧНИКА ПАРАМЕТРОВ');
    logInfo('='.repeat(60));

    // 1. Проверка существования справочника
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      logInfo('📝 Справочник не найден, создаем новый...');
      sheet = initializeSpecificationReferenceSheet();

      if (!sheet) {
        logError('❌ Не удалось создать справочник');
        return;
      }
    }

    logInfo(`✅ Справочник найден: "${sheet.getName()}"`);
    logInfo(`   Параметров в справочнике: ${sheet.getLastRow() - 1}`);

    // 2. Загрузка справочника
    const reference = loadSpecificationReference();
    logInfo(`✅ Загружено активных параметров: ${reference.length}`);

    if (reference.length > 0) {
      logInfo('\n📋 Первые 5 параметров справочника:');
      for (let i = 0; i < Math.min(5, reference.length); i++) {
        const param = reference[i];
        logInfo(`   ${i + 1}. ${param['Параметр (эталонное название)']}`);
        logInfo(`      Тип: ${param['Тип поля']}`);
        logInfo(`      Синонимы Veber: ${param['Синонимы Veber']}`);
        logInfo(`      Функция: ${param['Функция нормализации']}`);
      }
    }

    // 3. Тестовая нормализация
    logInfo('\n🔄 Тестируем нормализацию с тестовыми данными...');

    const testSpecs = {
      'Увеличение': '10x',
      'Диаметр объектива': '42 мм',
      'Призмы': 'Roof',
      'Стекло призм': 'BaK4',
      'Покрытие линз': 'FMC',
      'Бренд': 'VEBER',
      'Вес': '580 г',
      'Цвет': 'черный',
      'Влагозащита': 'Водонепроницаемый',
      'Заполнение': 'Азот'
    };

    logInfo('📥 Входные характеристики:');
    logInfo(JSON.stringify(testSpecs, null, 2));

    const result = normalizeSpecifications(testSpecs, 'veber');

    logInfo('\n📤 Результат нормализации:');
    logInfo(`   Нормализовано: ${Object.keys(result.normalized).length} параметров`);
    logInfo(`   Не сопоставлено: ${Object.keys(result.unmapped).length} параметров`);

    if (Object.keys(result.normalized).length > 0) {
      logInfo('\n✅ Нормализованные параметры:');
      for (const [key, value] of Object.entries(result.normalized)) {
        logInfo(`   • ${key}: "${value}"`);
      }
    }

    if (Object.keys(result.unmapped).length > 0) {
      logInfo('\n⚠️ Не сопоставленные параметры:');
      for (const [key, value] of Object.entries(result.unmapped)) {
        logInfo(`   • ${key}: "${value}"`);
      }
    }

    // 4. Валидация
    logInfo('\n🔍 Валидация нормализованных параметров...');
    const validation = validateAgainstReference(result.normalized);

    logInfo(`   Результат: ${validation.isValid ? '✅ Валидно' : '❌ Есть ошибки'}`);

    if (validation.errors.length > 0) {
      logInfo('\n❌ Ошибки валидации:');
      validation.errors.forEach(err => logInfo(`   • ${err}`));
    }

    if (validation.warnings.length > 0) {
      logInfo('\n⚠️ Предупреждения:');
      validation.warnings.forEach(warn => logInfo(`   • ${warn}`));
    }

    logInfo('\n' + '='.repeat(60));
    logInfo('✅ ТЕСТ ЗАВЕРШЕН');

    // Показываем результат пользователю
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Тест справочника выполнен',
      `Справочник работает!\n\n` +
      `Параметров в справочнике: ${reference.length}\n` +
      `Нормализовано из тестовых данных: ${Object.keys(result.normalized).length}\n` +
      `Не сопоставлено: ${Object.keys(result.unmapped).length}\n\n` +
      `Проверьте логи выполнения для подробностей.`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    handleError(error, 'Тест справочника параметров');
    const ui = SpreadsheetApp.getUi();
    ui.alert('❌ Ошибка теста', `Произошла ошибка:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

// =============================================================================
// УПРАВЛЕНИЕ НЕНОРМАЛИЗОВАННЫМИ ЗНАЧЕНИЯМИ
// =============================================================================

/**
 * ОБЪЕДИНЕННЫЙ ДИАЛОГ УПРАВЛЕНИЯ ПАРАМЕТРАМИ
 *
 * Показывает единый диалог с двумя вкладками:
 * - Вкладка 1: Сопоставление параметров конкретного товара
 * - Вкладка 2: Пакетная обработка ненормализованных значений
 */
function showUnifiedParameterDialog() {
  try {
    logInfo('🔄 Открываем объединенный диалог управления параметрами');

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
        'Пожалуйста, выберите товар с помощью чекбокса в первой колонке таблицы.\n\n' +
        'Это необходимо для вкладки "Товар".\n' +
        'Вкладка "Все ненормализованные значения" будет доступна в любом случае.',
        ui.ButtonSet.OK
      );
      return;
    }

    logInfo(`   Выбран товар: ${selectedArticle}`);

    // Создаем HTML диалог из файла
    const htmlTemplate = HtmlService.createTemplateFromFile('UnifiedParameterDialog');

    // ✅ ИСПРАВЛЕНИЕ: Передаем артикул через скриптлет в шаблон
    htmlTemplate.selectedArticle = selectedArticle;

    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(1200)
      .setHeight(700);

    // Показываем модальный диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔄 Управление параметрами');

    logInfo('✅ Объединенный диалог открыт');

  } catch (error) {
    handleError(error, 'Открытие объединенного диалога параметров');
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка',
      `Не удалось открыть диалог: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Показывает диалог для добавления ненормализованных значений в справочник
 * @deprecated Используйте showUnifiedParameterDialog() вместо этой функции
 */
function showUnnormalizedValuesDialog() {
  try {
    const unnormalizedValues = getUnnormalizedValues();

    // ✅ ДИАГНОСТИКА: Детальное логирование для отладки
    logInfo(`🔍 ДИАГНОСТИКА getUnnormalizedValues():`);
    logInfo(`   - Тип: ${typeof unnormalizedValues}`);
    logInfo(`   - Длина: ${unnormalizedValues ? unnormalizedValues.length : 'null/undefined'}`);
    logInfo(`   - Содержимое: ${JSON.stringify(unnormalizedValues, null, 2)}`);

    if (!unnormalizedValues || unnormalizedValues.length === 0) {
      SpreadsheetApp.getUi().alert(
        '✅ Нет ненормализованных значений',
        'Все enum значения из последнего импорта соответствуют справочнику.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    logInfo(`📋 Открываем диалог с ${unnormalizedValues.length} ненормализованными значениями`);

    // Создаем HTML диалог из файла
    // Примечание: данные будут загружены через google.script.run.getUnnormalizedValues() на клиенте
    const htmlTemplate = HtmlService.createTemplateFromFile('UnnormalizedValuesDialog');

    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(700)
      .setHeight(600);

    // Показываем модальный диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '➕ Добавить ненормализованные значения');

    logInfo('✅ Диалог открыт');

  } catch (error) {
    handleError(error, 'Открытие диалога ненормализованных значений');
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка',
      `Не удалось открыть диалог: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Добавляет значение в список допустимых значений параметра в справочнике
 *
 * @param {string} parameterName - Название параметра
 * @param {string} newValue - Новое значение для добавления
 * @returns {Object} Результат операции
 */
function addValueToSpecificationReference(parameterName, newValue) {
  const context = 'addValueToSpecificationReference';

  try {
    logInfo(`➕ Добавляем значение "${newValue}" к параметру "${parameterName}"`, null, context);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      throw new Error('Справочник параметров не найден');
    }

    // Получаем все данные справочника
    const data = sheet.getDataRange().getValues();

    // ✅ ИСПРАВЛЕНИЕ: Используем константы вместо поиска по заголовкам
    const parameterCol = SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1;  // -1 для индекса массива (0-based)
    const allowedValuesCol = SPEC_REFERENCE_COLUMNS.ALLOWED_VALUES - 1;

    // Ищем строку с нужным параметром
    // ✅ ИСПРАВЛЕНИЕ: Пробуем искать как с префиксом "Параметр:", так и без него
    let targetRow = -1;
    const parameterVariants = [
      parameterName,                    // Точное совпадение
      `Параметр: ${parameterName}`,    // С префиксом
      parameterName.replace(/^Параметр:\s*/i, '')  // Без префикса
    ];

    for (let i = 1; i < data.length; i++) {
      const cellValue = String(data[i][parameterCol] || '').trim();

      // Проверяем все варианты
      if (parameterVariants.some(variant => cellValue === variant)) {
        targetRow = i + 1; // +1 потому что getRange начинается с 1
        logInfo(`   🔍 Найден параметр в строке ${targetRow}: "${cellValue}"`, null, context);
        break;
      }
    }

    if (targetRow === -1) {
      // Логируем все параметры для отладки
      logWarning(`⚠️ Параметр "${parameterName}" не найден. Доступные параметры:`, null, context);
      for (let i = 1; i < Math.min(data.length, 10); i++) {
        logInfo(`   - "${data[i][parameterCol]}"`, null, context);
      }
      throw new Error(`Параметр "${parameterName}" не найден в справочнике`);
    }

    // Получаем текущие допустимые значения
    const currentValues = data[targetRow - 1][allowedValuesCol] || '';
    const valuesArray = currentValues
      .split(';')
      .map(v => v.trim())
      .filter(v => v);

    // Проверяем, нет ли уже такого значения
    const trimmedNewValue = newValue.trim();
    const existingValue = valuesArray.find(v => v.toLowerCase() === trimmedNewValue.toLowerCase());

    if (existingValue) {
      logWarning(`⚠️ Значение "${trimmedNewValue}" уже существует как "${existingValue}"`, null, context);
      return {
        success: false,
        error: `Значение "${existingValue}" уже есть в списке допустимых значений`
      };
    }

    // Добавляем новое значение
    valuesArray.push(trimmedNewValue);
    const newValuesString = valuesArray.join('; ');

    // Записываем обновленные значения
    sheet.getRange(targetRow, allowedValuesCol + 1).setValue(newValuesString);

    logInfo(`✅ Значение "${trimmedNewValue}" добавлено к параметру "${parameterName}"`, null, context);
    logInfo(`   Новый список: ${newValuesString}`, null, context);

    // Показываем уведомление пользователю
    ss.toast(
      `Значение "${trimmedNewValue}" добавлено к параметру "${parameterName}"`,
      '✅ Успех',
      3
    );

    return {
      success: true,
      parameterName: parameterName,
      addedValue: trimmedNewValue,
      allValues: newValuesString
    };

  } catch (error) {
    handleError(error, `Добавление значения "${newValue}" к параметру "${parameterName}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Добавляет новое допустимое значение к параметру в справочнике
 * 
 * @param {string} parameterName - Название параметра в справочнике
 * @param {string} newValue - Новое значение для добавления в список допустимых
 * @returns {Object} Результат операции {success, error}
 */
function addAllowedValueToParameter(parameterName, newValue) {
  const context = 'addAllowedValueToParameter';

  try {
    logInfo(`💾 Добавление значения "${newValue}" к параметру "${parameterName}"`, null, context);

    if (!parameterName || !newValue) {
      throw new Error('Название параметра и значение обязательны');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      throw new Error('Справочник параметров не найден. Создайте его через меню.');
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      throw new Error('Справочник пуст');
    }

    // Ищем строку с параметром
    // ✅ ИСПРАВЛЕНИЕ: Удаляем префикс "Параметр:" при сравнении (как в loadSpecificationReference)
    let parameterRow = -1;
    const normalizedSearchParam = String(parameterName).trim();

    for (let i = 1; i < data.length; i++) {
      const cellParamName = data[i][SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1];
      const normalizedCellParam = String(cellParamName || '').replace(/^Параметр:\s*/i, '').trim();

      if (normalizedCellParam === normalizedSearchParam) {
        parameterRow = i + 1; // +1 для индекса листа (начинается с 1)
        break;
      }
    }

    if (parameterRow === -1) {
      throw new Error(`Параметр "${parameterName}" не найден в справочнике`);
    }

    // Получаем текущие допустимые значения
    const allowedValuesCell = sheet.getRange(parameterRow, SPEC_REFERENCE_COLUMNS.ALLOWED_VALUES);
    const currentAllowedValues = allowedValuesCell.getValue() || '';

    // Разбираем существующие значения
    const valuesArray = currentAllowedValues
      .split(';')
      .map(v => v.trim())
      .filter(v => v.length > 0);

    // Проверяем, не существует ли уже такое значение
    const normalizedNewValue = newValue.trim();
    if (valuesArray.some(v => v === normalizedNewValue)) {
      logWarning(`⚠️ Значение "${normalizedNewValue}" уже существует в списке допустимых`, null, context);
      return {
        success: true,
        message: 'Значение уже существует в списке',
        alreadyExists: true
      };
    }

    // Добавляем новое значение
    valuesArray.push(normalizedNewValue);

    // Сортируем для удобства (опционально)
    valuesArray.sort();

    // Обновляем ячейку
    const updatedAllowedValues = valuesArray.join('; ');
    allowedValuesCell.setValue(updatedAllowedValues);

    logInfo(`✅ Значение "${normalizedNewValue}" добавлено к параметру "${parameterName}"`, {
      row: parameterRow,
      totalValues: valuesArray.length
    }, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Значение "${normalizedNewValue}" добавлено в список допустимых значений параметра "${parameterName}"`,
      '✅ Справочник обновлен',
      5
    );

    return {
      success: true,
      message: 'Значение успешно добавлено',
      parameterName: parameterName,
      newValue: normalizedNewValue,
      totalValues: valuesArray.length
    };

  } catch (error) {
    handleError(error, `Добавление значения "${newValue}" к параметру "${parameterName}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ✅ НОВОЕ: Возвращает список всех параметров из справочника для автоподстановки
 *
 * @returns {Array<string>} Массив названий параметров
 */
function getAllParameterNames() {
  const context = 'getAllParameterNames';

  try {
    logInfo('📋 Загружаем список параметров из справочника', null, context);

    const reference = loadSpecificationReference();

    if (!reference || reference.length === 0) {
      logWarning('⚠️ Справочник пуст', null, context);
      return [];
    }

    // Извлекаем только названия параметров
    const parameterNames = reference
      .map(row => row['Параметр (эталонное название)'])
      .filter(name => name && name.trim() !== '');

    logInfo(`✅ Загружено ${parameterNames.length} параметров`, null, context);

    return parameterNames;

  } catch (error) {
    handleError(error, 'Загрузка списка параметров');
    return [];
  }
}

/**
 * ✅ НОВОЕ: Ре-нормализует характеристики конкретного товара
 * Вызывается после изменения значений в справочнике
 *
 * @param {string} article - Артикул товара
 * @returns {Object} Результат операции
 */
function renormalizeProduct(article) {
  const context = 'renormalizeProduct';

  try {
    logInfo(`🔄 Ре-нормализуем товар "${article}"`, null, context);

    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    // ✅ ДИАГНОСТИКА: Проверяем тип искомого артикула
    logInfo(`🔍 Ищем артикул: "${article}" (тип: ${typeof article})`, null, context);

    // Ищем строку товара
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const cellArticle = data[i][IMAGES_COLUMNS.ARTICLE - 1];

      // ✅ ИСПРАВЛЕНИЕ: Приводим оба значения к строке для сравнения
      if (String(cellArticle) === String(article)) {
        logInfo(`   ✅ Найдено совпадение в строке ${i + 1}: "${cellArticle}" (тип: ${typeof cellArticle})`, null, context);
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      // ✅ ДИАГНОСТИКА: Показываем первые несколько артикулов из таблицы
      logInfo(`   ❌ Не найдено! Первые 5 артикулов в таблице:`, null, context);
      for (let i = 1; i < Math.min(6, data.length); i++) {
        const cellArticle = data[i][IMAGES_COLUMNS.ARTICLE - 1];
        logInfo(`      Строка ${i + 1}: "${cellArticle}" (тип: ${typeof cellArticle})`, null, context);
      }
      throw new Error(`Товар с артикулом "${article}" не найден`);
    }

    // Получаем сырые характеристики из колонки P
    const rawSpecsStr = data[targetRow - 1][IMAGES_COLUMNS.SPECIFICATIONS_RAW - 1];
    if (!rawSpecsStr) {
      throw new Error('Сырые характеристики отсутствуют');
    }

    // ✅ ИСПРАВЛЕНИЕ: Проверяем, что это валидный JSON
    let rawSpecs;
    try {
      rawSpecs = JSON.parse(rawSpecsStr);
    } catch (parseError) {
      logWarning(`⚠️ В колонке P хранится НЕ JSON для товара "${article}"`, null, context);
      logInfo(`   Содержимое (первые 200 символов): ${String(rawSpecsStr).substring(0, 200)}`, null, context);

      throw new Error(
        `Невозможно ре-нормализовать товар "${article}": ` +
        `в колонке P (Сырые характеристики) хранится не JSON. ` +
        `Возможно, товар был импортирован некорректно. ` +
        `Попробуйте удалить строку и импортировать товар заново.`
      );
    }

    // Определяем поставщика (из названия или другого поля)
    // TODO: добавить колонку SUPPLIER если нужно
    const supplier = 'veber'; // Временно используем veber

    // Нормализуем заново
    const result = normalizeSpecifications(rawSpecs, supplier, article);

    // Записываем обновленные нормализованные характеристики в колонку Q
    const normalizedJson = JSON.stringify(result.normalized);
    sheet.getRange(targetRow, IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED).setValue(normalizedJson);

    logInfo(`✅ Товар "${article}" ре-нормализован`, null, context);

    return {
      success: true,
      article: article,
      normalizedCount: Object.keys(result.normalized).length
    };

  } catch (error) {
    handleError(error, `Ре-нормализация товара "${article}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ✅ НОВОЕ: Применяет изменения после редактирования значений в диалоге
 * Ре-нормализует все товары, которые содержали измененные значения
 *
 * @param {Array<Object>} changes - Массив изменений [{parameter, oldValue, newValue, action}]
 * @returns {Object} Результат операции с количеством обновленных товаров
 */
function applyUnnormalizedValuesChanges(changes) {
  const context = 'applyUnnormalizedValuesChanges';

  try {
    logInfo(`🔄 Применяем изменения из диалога (${changes.length} изменений)`, null, context);

    const unnormalizedValues = UNNORMALIZED_VALUES.getAll();
    const articlesToRenormalize = new Set();

    // 🔍 ДИАГНОСТИКА: показываем все изменения и доступные ненормализованные значения
    logInfo(`📋 Изменения из диалога:`, null, context);
    changes.forEach(ch => {
      logInfo(`   - Параметр: "${ch.parameter}", Старое: "${ch.oldValue}", Новое: "${ch.newValue}", Действие: ${ch.action}`, null, context);
    });

    logInfo(`📋 Доступные ненормализованные значения (${unnormalizedValues.length}):`, null, context);
    unnormalizedValues.forEach(uv => {
      logInfo(`   - Параметр: "${uv.parameter}", Значение: "${uv.value}", Артикулов: ${uv.articles ? uv.articles.length : 0}`, null, context);
    });

    // Собираем артикулы товаров, которые нужно обновить
    for (const change of changes) {
      logInfo(`🔍 Ищем ненормализованное значение для: параметр="${change.parameter}", значение="${change.oldValue}"`, null, context);

      // Ищем соответствующее ненормализованное значение
      const unnormalized = unnormalizedValues.find(item =>
        item.parameter === change.parameter &&
        item.value === change.oldValue
      );

      if (unnormalized) {
        logInfo(`   ✅ Найдено! Артикулов: ${unnormalized.articles ? unnormalized.articles.length : 0}`, null, context);
        if (unnormalized.articles) {
          unnormalized.articles.forEach(article => {
            logInfo(`      - ${article}`, null, context);
            articlesToRenormalize.add(article);
          });
        }
      } else {
        logInfo(`   ❌ НЕ НАЙДЕНО совпадение!`, null, context);
      }
    }

    logInfo(`📋 Найдено товаров для ре-нормализации: ${articlesToRenormalize.size}`, null, context);

    // Ре-нормализуем каждый товар
    const results = {
      success: 0,
      failed: 0,
      articles: []
    };

    for (const article of articlesToRenormalize) {
      const result = renormalizeProduct(article);
      if (result.success) {
        results.success++;
        results.articles.push(article);
      } else {
        results.failed++;
      }
    }

    logInfo(`✅ Обновлено товаров: ${results.success}, ошибок: ${results.failed}`, null, context);

    // Показываем уведомление
    if (results.success > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Обновлено характеристик для ${results.success} товаров`,
        '✅ Изменения применены',
        5
      );
    }

    return {
      success: true,
      updatedCount: results.success,
      failedCount: results.failed,
      articles: results.articles
    };

  } catch (error) {
    handleError(error, 'Применение изменений');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🔧 ВРЕМЕННАЯ ФУНКЦИЯ ДИАГНОСТИКИ
 * Проверяет содержимое колонки P для товара 32894
 */
function DEBUG_checkRawSpecsFor32894() {
  const context = 'DEBUG_checkRawSpecsFor32894';

  try {
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    logInfo('🔍 Ищем товар 32894...', null, context);

    // Ищем товар 32894
    for (let i = 1; i < data.length; i++) {
      const article = String(data[i][IMAGES_COLUMNS.ARTICLE - 1]);

      if (article === '32894') {
        const rawSpecs = data[i][IMAGES_COLUMNS.SPECIFICATIONS_RAW - 1];

        logInfo(`✅ Найден товар 32894 в строке ${i + 1}`, null, context);
        logInfo(`📊 Тип данных: ${typeof rawSpecs}`, null, context);
        logInfo(`📏 Длина: ${rawSpecs ? String(rawSpecs).length : 0} символов`, null, context);
        logInfo(`📝 Первые 500 символов:`, null, context);
        logInfo(String(rawSpecs).substring(0, 500), null, context);

        // Пробуем распарсить
        try {
          const parsed = JSON.parse(rawSpecs);
          logInfo(`✅ JSON валиден! Полей: ${Object.keys(parsed).length}`, null, context);
          logInfo(`📋 Первые 3 поля:`, null, context);
          Object.keys(parsed).slice(0, 3).forEach(key => {
            logInfo(`   - "${key}": "${parsed[key]}"`, null, context);
          });
        } catch (e) {
          logCritical(`❌ JSON НЕВАЛИДЕН! Ошибка: ${e.message}`, null, context);
          logInfo(`🔍 Содержимое (полностью):`, null, context);
          logInfo(String(rawSpecs), null, context);
        }

        return {
          success: true,
          row: i + 1,
          dataType: typeof rawSpecs,
          length: rawSpecs ? String(rawSpecs).length : 0,
          preview: String(rawSpecs).substring(0, 500)
        };
      }
    }

    logCritical('❌ Товар 32894 не найден!', null, context);
    return {
      success: false,
      error: 'Товар не найден'
    };

  } catch (error) {
    handleError(error, 'Диагностика товара 32894');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ✅ НОВОЕ: Автоматически исправляет пустое поле нормализатора
 * для параметра "Защита от влаги и пыли"
 */
function fixWaterproofingNormalizerInReference() {
  const context = 'fixWaterproofingNormalizerInReference';

  try {
    logInfo('🔧 Исправляем пустое поле нормализатора для "Защита от влаги и пыли"', null, context);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      throw new Error(`Лист "${SHEET_NAMES.SPEC_REFERENCE}" не найден`);
    }

    const data = sheet.getDataRange().getValues();

    // Ищем строку с параметром "Защита от влаги и пыли"
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const parameterName = data[i][SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1];

      if (String(parameterName).includes('Защита от влаги')) {
        targetRow = i + 1; // +1 потому что индексация с 1

        const currentNormalizer = data[i][SPEC_REFERENCE_COLUMNS.NORMALIZER_FUNCTION - 1];
        logInfo(`   ✅ Найдена строка ${targetRow}: "${parameterName}"`, null, context);
        logInfo(`   - Текущее значение нормализатора: "${currentNormalizer}"`, null, context);

        // Проверяем, пусто ли поле
        if (!currentNormalizer || String(currentNormalizer).trim() === '') {
          // Записываем правильное название функции
          sheet.getRange(targetRow, SPEC_REFERENCE_COLUMNS.NORMALIZER_FUNCTION)
            .setValue('normalizeWaterproofing');

          logInfo(`   ✅ Исправлено! Установлено: "normalizeWaterproofing"`, null, context);

          SpreadsheetApp.getActiveSpreadsheet().toast(
            'Поле нормализатора исправлено для параметра "Защита от влаги и пыли"',
            '✅ Исправление применено',
            5
          );

          return {
            success: true,
            message: 'Поле нормализатора успешно исправлено'
          };
        } else {
          logInfo(`   ⚠️ Поле уже заполнено: "${currentNormalizer}"`, null, context);
          return {
            success: true,
            message: 'Поле нормализатора уже заполнено',
            currentValue: currentNormalizer
          };
        }
      }
    }

    throw new Error('Параметр "Защита от влаги и пыли" не найден в справочнике');

  } catch (error) {
    handleError(error, 'Исправление нормализатора');
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// ДИАЛОГ СОПОСТАВЛЕНИЯ ПАРАМЕТРОВ
// =============================================================================

/**
 * Получает данные для диалога сопоставления параметров
 * Читает сырые спецификации товара и возвращает matched/unmapped параметры
 *
 * @param {string} article - Артикул товара
 * @returns {Object} Данные для диалога
 */
function getParameterMappingData(article) {
  const context = 'getParameterMappingData';

  try {
    // 🔍 ДИАГНОСТИКА: Что пришло на сервер
    logInfo(`🔍 DEBUG getParameterMappingData:`);
    logInfo(`   - article = "${article}"`);
    logInfo(`   - typeof article = ${typeof article}`);
    logInfo(`   - article.length = ${article ? article.length : 'undefined'}`);

    logInfo(`📋 Загружаем данные сопоставления для артикула "${article}"`, null, context);

    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    // Ищем строку товара
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const cellArticle = data[i][IMAGES_COLUMNS.ARTICLE - 1];

      if (String(cellArticle) === String(article)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(`Товар с артикулом "${article}" не найден`);
    }

    // Получаем данные товара
    const productName = data[targetRow - 1][IMAGES_COLUMNS.PRODUCT_NAME - 1];
    const rawSpecsStr = data[targetRow - 1][IMAGES_COLUMNS.SPECIFICATIONS_RAW - 1];
    const normalizedSpecsStr = data[targetRow - 1][IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED - 1];

    if (!rawSpecsStr) {
      throw new Error('Сырые характеристики отсутствуют');
    }

    // Парсим сырые характеристики
    let rawSpecs;
    try {
      rawSpecs = JSON.parse(rawSpecsStr);
    } catch (parseError) {
      throw new Error(`Невозможно распарсить характеристики: ${parseError.message}`);
    }

    // ✅ ИСПРАВЛЕНИЕ: Загружаем актуальные нормализованные данные из столбца Q
    let normalizedSpecs = {};
    if (normalizedSpecsStr) {
      try {
        normalizedSpecs = JSON.parse(normalizedSpecsStr);
        logInfo(`✅ Загружены актуальные нормализованные данные из столбца Q (${Object.keys(normalizedSpecs).length} параметров)`, null, context);
      } catch (parseError) {
        logWarning(`⚠️ Не удалось распарсить SPECIFICATIONS_NORMALIZED, будет проведена свежая нормализация: ${parseError.message}`, null, context);
      }
    }

    // Определяем поставщика (TODO: добавить колонку SUPPLIER)
    const supplier = 'veber'; // Временно

    // ✅ ИСПРАВЛЕНИЕ: Если нет нормализованных данных, проводим свежую нормализацию
    let unmappedSpecs = {};
    if (Object.keys(normalizedSpecs).length === 0) {
      logInfo('⚠️ Нет нормализованных данных в столбце Q, проводим свежую нормализацию', null, context);
      const result = normalizeSpecifications(rawSpecs, supplier, article);
      normalizedSpecs = result.normalized;
      unmappedSpecs = result.unmapped;
    } else {
      // Находим unmapped поля (есть в raw, но нет в normalized)
      for (const [rawKey, rawValue] of Object.entries(rawSpecs)) {
        // Проверяем, есть ли это поле в нормализованных данных
        let isMapped = false;
        const reference = loadSpecificationReference();

        for (const refParam of reference) {
          const synonymsField = supplier === 'veber' ? 'Синонимы Veber' : 'Синонимы Sturman';
          const synonyms = (refParam[synonymsField] || '').split(';').map(s => s.trim()).filter(s => s);
          const paramName = refParam['Параметр (эталонное название)'];

          // Проверяем, совпадает ли поле с каким-то синонимом
          for (const synonym of synonyms) {
            if (rawKey.toLowerCase().includes(synonym.toLowerCase())) {
              // Проверяем, есть ли этот параметр в normalizedSpecs
              if (normalizedSpecs.hasOwnProperty(paramName)) {
                isMapped = true;
                break;
              }
            }
          }
          if (isMapped) break;
        }

        if (!isMapped) {
          unmappedSpecs[rawKey] = rawValue;
        }
      }
    }

    // Загружаем все параметры справочника для autocomplete
    const reference = loadSpecificationReference();
    const dictionaryParams = reference.map(param => ({
      name: param['Параметр (эталонное название)'],
      fieldType: param['Тип поля'],
      allowedValues: param['Допустимые значения (enum)'],
      normalizerFunction: param['Функция нормализации'] || ''
    }));

    // ✅ ИСПРАВЛЕНИЕ: Формируем matched параметры из normalizedSpecs (актуальные данные из столбца Q)
    const matched = [];
    for (const [normalizedKey, normalizedValue] of Object.entries(normalizedSpecs)) {
      // ✅ НОВОЕ: Парсим ключ для извлечения метаданных
      const parsed = parseParameterKey(normalizedKey);
      const parameterName = parsed.parameterName;

      // Находим исходное поле поставщика для этого параметра
      const refParam = reference.find(r => r['Параметр (эталонное название)'] === parameterName);

      let supplierFieldName = '';
      let supplierValue = '';

      // ✅ НОВОЕ: Если в ключе есть метаданные - используем их (приоритет!)
      if (parsed.supplierField && parsed.supplierValue) {
        supplierFieldName = parsed.supplierField;
        supplierValue = parsed.supplierValue;
        logInfo(`   📦 Метаданные найдены в ключе для "${parameterName}": ${supplierFieldName} = ${supplierValue}`, null, context);
      } else if (refParam) {
        // Метаданных нет - ищем через синонимы (старый способ)
        const synonymsField = supplier === 'veber' ? 'Синонимы Veber' : 'Синонимы Sturman';
        const synonyms = (refParam[synonymsField] || '').split(';').map(s => s.trim()).filter(s => s);

        // ✅ УЛУЧШЕНИЕ: Ищем совпадение в rawSpecs с двусторонней проверкой
        for (const synonym of synonyms) {
          const synonymLower = synonym.toLowerCase();

          for (const [rawKey, rawValue] of Object.entries(rawSpecs)) {
            const rawKeyLower = rawKey.toLowerCase();

            // Проверка 1: Точное совпадение
            if (rawKeyLower === synonymLower) {
              supplierFieldName = rawKey;
              supplierValue = rawValue;
              break;
            }

            // Проверка 2: Поле поставщика содержит синоним
            if (rawKeyLower.includes(synonymLower)) {
              supplierFieldName = rawKey;
              supplierValue = rawValue;
              break;
            }

            // Проверка 3: Синоним содержит поле поставщика
            if (synonymLower.includes(rawKeyLower)) {
              supplierFieldName = rawKey;
              supplierValue = rawValue;
              break;
            }
          }
          if (supplierFieldName) break;
        }
      }

      matched.push({
        supplierFieldName: supplierFieldName,
        supplierValue: supplierValue,
        dictionaryParam: parameterName, // ✅ ИЗМЕНЕНО: используем очищенное имя параметра
        normalizedValue: normalizedValue,
        fieldType: refParam ? refParam['Тип поля'] : 'текст',
        allowedValues: refParam ? refParam['Допустимые значения (enum)'] : ''
      });
    }

    // ✅ ИСПРАВЛЕНИЕ: Формируем unmapped параметры из unmappedSpecs
    const unmapped = [];
    for (const [supplierFieldName, supplierValue] of Object.entries(unmappedSpecs)) {
      unmapped.push({
        supplierFieldName: supplierFieldName,
        supplierValue: supplierValue,
        dictionaryParam: '', // Пусто - пользователь должен выбрать
        normalizedValue: '',
        fieldType: '',
        allowedValues: ''
      });
    }

    logInfo(`✅ Данные загружены: matched=${matched.length}, unmapped=${unmapped.length}`, null, context);

    return {
      success: true,
      article: article,
      productName: productName,
      supplier: supplier,
      matched: matched,
      unmapped: unmapped,
      dictionaryParams: dictionaryParams
    };

  } catch (error) {
    handleError(error, `Загрузка данных сопоставления для "${article}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Загружает все параметры из справочника для dropdown-списков
 * 
 * @returns {Object} { success, parameters: [{name, fieldType, allowedValues, required}] }
 */
function getParametersListData() {
  const context = 'getParametersListData';

  try {
    logInfo('📋 Загружаем список параметров для dropdown', null, context);

    const reference = loadSpecificationReference();

    if (!reference || reference.length === 0) {
      return {
        success: true,
        parameters: []
      };
    }

    // Формируем список параметров с их значениями
    const parameters = reference.map(param => ({
      name: param['Параметр (эталонное название)'],
      fieldType: param['Тип поля'] || 'текст',
      allowedValues: param['Допустимые значения (enum)'] || '',
      required: param['Обязательное поле'] === 'да',
      normalizerFunction: param['Функция нормализации'] || '',
      description: param['Описание'] || ''
    }));

    // Сортируем по названию
    parameters.sort((a, b) => a.name.localeCompare(b.name, 'ru'));

    logInfo(`✅ Загружено ${parameters.length} параметров`, null, context);

    return {
      success: true,
      parameters: parameters
    };

  } catch (error) {
    handleError(error, 'Загрузка списка параметров');
    return {
      success: false,
      error: error.message,
      parameters: []
    };
  }
}

/**
 * Создает новый параметр в справочнике
 *
 * @param {Object} paramData - Данные параметра
 * @param {string} paramData.name - Название параметра
 * @param {string} paramData.fieldType - Тип поля (enum/число/текст/формат)
 * @param {string} paramData.allowedValues - Допустимые значения через ";" (для enum)
 * @param {boolean} paramData.required - Обязательное поле
 * @param {string} paramData.normalizerFunction - Функция нормализации
 * @param {string} paramData.description - Описание параметра
 * @returns {Object} Результат операции
 */
function createNewParameter(paramData) {
  const context = 'createNewParameter';

  try {
    logInfo(`➕ Создаем новый параметр "${paramData.name}"`, null, context);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      throw new Error(`Лист "${SHEET_NAMES.SPEC_REFERENCE}" не найден`);
    }

    // Проверяем, не существует ли уже такой параметр
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const existingName = String(data[i][SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1] || '');
      // Убираем префикс "Параметр:" для сравнения
      const existingNameNormalized = existingName.replace(/^Параметр:\s*/i, '').toLowerCase();
      const paramNameNormalized = paramData.name.replace(/^Параметр:\s*/i, '').toLowerCase();

      if (existingNameNormalized === paramNameNormalized) {
        throw new Error(`Параметр "${paramData.name}" уже существует в справочнике`);
      }
    }

    // Добавляем новую строку в конец
    const lastRow = sheet.getLastRow() + 1;

    const rowData = [
      true, // Чекбокс (активен)
      paramData.name,
      paramData.fieldType || 'текст',
      paramData.allowedValues || '',
      paramData.required ? 'да' : 'нет',
      '', // Синонимы Veber (пока пусто)
      '', // Синонимы Sturman (пока пусто)
      paramData.normalizerFunction || '',
      paramData.description || ''
    ];

    sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);

    // Устанавливаем чекбокс
    sheet.getRange(lastRow, 1).insertCheckboxes();

    logInfo(`✅ Параметр "${paramData.name}" создан в строке ${lastRow}`, null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Параметр "${paramData.name}" добавлен в справочник`,
      '✅ Параметр создан',
      3
    );

    return {
      success: true,
      parameter: {
        name: paramData.name,
        fieldType: paramData.fieldType,
        allowedValues: paramData.allowedValues,
        row: lastRow
      }
    };

  } catch (error) {
    handleError(error, `Создание параметра "${paramData.name}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Добавляет синоним к параметру в справочнике
 *
 * @param {string} parameterName - Название параметра
 * @param {string} supplierName - Поставщик (veber/sturman)
 * @param {string} synonym - Синоним для добавления
 * @returns {Object} Результат операции
 */
function addSynonymToParameter(parameterName, supplierName, synonym) {
  const context = 'addSynonymToParameter';

  try {
    logInfo(`➕ Добавляем синоним "${synonym}" к параметру "${parameterName}" (${supplierName})`, null, context);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SPEC_REFERENCE);

    if (!sheet) {
      throw new Error(`Лист "${SHEET_NAMES.SPEC_REFERENCE}" не найден`);
    }

    const data = sheet.getDataRange().getValues();

    // Ищем параметр
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const cellParam = String(data[i][SPEC_REFERENCE_COLUMNS.PARAMETER_NAME - 1] || '').trim();
      // Убираем префикс "Параметр:" для сравнения (как в loadSpecificationReference)
      const cellParamNormalized = cellParam.replace(/^Параметр:\s*/i, '');

      if (cellParamNormalized === parameterName || cellParam === parameterName) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(`Параметр "${parameterName}" не найден в справочнике`);
    }

    // Определяем колонку синонимов
    const synonymCol = supplierName === 'veber'
      ? SPEC_REFERENCE_COLUMNS.VEBER_SYNONYMS
      : SPEC_REFERENCE_COLUMNS.STURMAN_SYNONYMS;

    // Получаем текущие синонимы
    const currentSynonyms = data[targetRow - 1][synonymCol - 1] || '';
    const synonymsArray = currentSynonyms
      .split(';')
      .map(s => s.trim())
      .filter(s => s);

    // Проверяем, не существует ли уже такой синоним
    const trimmedSynonym = synonym.trim();
    if (synonymsArray.some(s => s.toLowerCase() === trimmedSynonym.toLowerCase())) {
      logWarning(`⚠️ Синоним "${trimmedSynonym}" уже существует`, null, context);
      return {
        success: false,
        error: `Синоним "${trimmedSynonym}" уже существует в списке`
      };
    }

    // Добавляем новый синоним
    synonymsArray.push(trimmedSynonym);
    const newSynonymsString = synonymsArray.join('; ');

    // Записываем обновленные синонимы
    sheet.getRange(targetRow, synonymCol).setValue(newSynonymsString);

    logInfo(`✅ Синоним "${trimmedSynonym}" добавлен к параметру "${parameterName}"`, null, context);

    return {
      success: true,
      parameterName: parameterName,
      addedSynonym: trimmedSynonym,
      allSynonyms: newSynonymsString
    };

  } catch (error) {
    handleError(error, `Добавление синонима "${synonym}" к "${parameterName}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Добавляет enum значение к параметру в справочнике
 * (Обертка над существующей функцией addValueToSpecificationReference)
 *
 * @param {string} parameterName - Название параметра
 * @param {string} newValue - Новое значение
 * @returns {Object} Результат операции
 */
function addEnumValueToParameter(parameterName, newValue) {
  return addValueToSpecificationReference(parameterName, newValue);
}

/**
 * Применяет сопоставления только к текущему товару (локальная операция)
 *
 * @param {string} article - Артикул товара
 * @param {Array<Object>} mappings - Массив сопоставлений
 * @returns {Object} Результат операции
 */
function applyLocalMappings(article, mappings) {
  const context = 'applyLocalMappings';

  try {
    logInfo(`🔄 Применяем локальные сопоставления к товару "${article}" (${mappings.length} изменений)`, null, context);

    // 1. Обрабатываем изменения matched параметров - обновляем SPECIFICATIONS_NORMALIZED напрямую
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();
    let productRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][IMAGES_COLUMNS.ARTICLE - 1]) === String(article)) {
        productRow = i + 1;
        break;
      }
    }

    if (productRow === -1) {
      throw new Error(`Товар с артикулом "${article}" не найден`);
    }

    // Загружаем текущие нормализованные характеристики
    const normalizedJson = data[productRow - 1][IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED - 1];
    let normalized = {};

    if (normalizedJson) {
      try {
        normalized = JSON.parse(normalizedJson);
      } catch (e) {
        logWarning(`⚠️ Не удалось распарсить SPECIFICATIONS_NORMALIZED: ${e.message}`, null, context);
      }
    }

    let hasMatchedChanges = false;
    for (const mapping of mappings) {
      if (mapping.action === 'modifyMatched') {
        logInfo(`🔄 Изменен matched параметр: "${mapping.oldParam}" → "${mapping.newParam}"`, null, context);
        logInfo(`   Старое значение: "${mapping.oldValue}"`, null, context);
        logInfo(`   Новое значение: "${mapping.newValue}"`, null, context);

        hasMatchedChanges = true;

        // ✅ НОВОЕ: Создаём ключ с метаданными (сохраняем связь с полем поставщика)
        const newKey = createParameterKey(
          mapping.newParam,
          mapping.supplierFieldName || '',
          mapping.supplierValue || ''
        );

        // ✅ НОВОЕ: Используем preserveOrderUpdate для сохранения порядка параметров
        // Находим старый ключ (может быть с метаданными или без)
        let oldKey = null;
        for (const key of Object.keys(normalized)) {
          const parsed = parseParameterKey(key);
          if (parsed.parameterName === mapping.oldParam) {
            oldKey = key;
            break;
          }
        }

        // Если не нашли старый ключ, используем просто имя параметра
        if (!oldKey) {
          oldKey = mapping.oldParam;
        }

        // Обновляем с сохранением порядка
        normalized = preserveOrderUpdate(normalized, oldKey, newKey, mapping.newValue);

        logInfo(`✅ Обновлено в normalized: "${newKey}" = "${mapping.newValue}"`, null, context);
        logInfo(`   📦 Метаданные: supplier="${mapping.supplierFieldName || 'нет'}", value="${mapping.supplierValue || 'нет'}"`, null, context);
      }

      if (mapping.action === 'addSynonym' && mapping.supplierFieldName && mapping.dictionaryParam) {
        const result = addSynonymToParameter(
          mapping.dictionaryParam,
          mapping.supplier || 'veber',
          mapping.supplierFieldName
        );

        if (!result.success) {
          logWarning(`⚠️ Не удалось добавить синоним: ${result.error}`, null, context);
        }
      }

      if (mapping.action === 'addEnumValue' && mapping.dictionaryParam && mapping.newEnumValue) {
        const result = addEnumValueToParameter(mapping.dictionaryParam, mapping.newEnumValue);

        if (!result.success) {
          logWarning(`⚠️ Не удалось добавить enum значение: ${result.error}`, null, context);
        }
      }
    }

    // 2. Сохраняем обновленные характеристики, если были изменения matched параметров
    if (hasMatchedChanges) {
      const normalizedJsonUpdated = JSON.stringify(normalized);
      sheet.getRange(productRow, IMAGES_COLUMNS.SPECIFICATIONS_NORMALIZED).setValue(normalizedJsonUpdated);
      logInfo(`✅ Обновлены нормализованные характеристики в ячейке Q${productRow}`, null, context);
      SpreadsheetApp.flush();
    }

    // 3. Пере-нормализуем только текущий товар (для unmapped параметров)
    const hasUnmappedChanges = mappings.some(m => m.action === 'addSynonym' || m.action === 'addEnumValue');

    let result = { success: true, normalizedCount: 0 };

    if (hasUnmappedChanges) {
      result = renormalizeProduct(article);
    }

    if (!result.success) {
      throw new Error(`Не удалось пере-нормализовать товар: ${result.error}`);
    }

    logInfo(`✅ Локальные сопоставления применены к товару "${article}"`, null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Характеристики товара "${article}" обновлены`,
      '✅ Изменения применены',
      3
    );

    return {
      success: true,
      article: article,
      normalizedCount: result.normalizedCount
    };

  } catch (error) {
    handleError(error, `Применение локальных сопоставлений к "${article}"`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Применяет сопоставления глобально (обновляет Справочник + пере-нормализует все товары)
 *
 * @param {Array<Object>} mappings - Массив сопоставлений
 * @returns {Object} Результат операции с количеством обновленных товаров
 */
function applyGlobalMappings(mappings) {
  const context = 'applyGlobalMappings';

  try {
    logInfo(`🌍 Применяем глобальные сопоставления (${mappings.length} изменений)`, null, context);

    // 1. Сохраняем все новые синонимы и значения в Справочник
    for (const mapping of mappings) {
      if (mapping.action === 'addSynonym' && mapping.supplierFieldName && mapping.dictionaryParam) {
        const result = addSynonymToParameter(
          mapping.dictionaryParam,
          mapping.supplier || 'veber',
          mapping.supplierFieldName
        );

        if (!result.success) {
          logWarning(`⚠️ Не удалось добавить синоним: ${result.error}`, null, context);
        }
      }

      if (mapping.action === 'addEnumValue' && mapping.dictionaryParam && mapping.newEnumValue) {
        const result = addEnumValueToParameter(mapping.dictionaryParam, mapping.newEnumValue);

        if (!result.success) {
          logWarning(`⚠️ Не удалось добавить enum значение: ${result.error}`, null, context);
        }
      }
    }

    // 2. Пере-нормализуем ВСЕ товары в таблице
    const sheet = getImagesSheet();
    const data = sheet.getDataRange().getValues();

    let updatedCount = 0;
    let failedCount = 0;
    const updatedArticles = [];

    for (let i = 1; i < data.length; i++) {
      const article = String(data[i][IMAGES_COLUMNS.ARTICLE - 1] || '');
      const rawSpecs = data[i][IMAGES_COLUMNS.SPECIFICATIONS_RAW - 1];

      // Пропускаем товары без сырых характеристик
      if (!rawSpecs) continue;

      // Проверяем, что это валидный JSON
      try {
        JSON.parse(rawSpecs);
      } catch (e) {
        logWarning(`⚠️ Пропускаем товар "${article}" - невалидный JSON в column P`, null, context);
        failedCount++;
        continue;
      }

      // Пере-нормализуем
      const result = renormalizeProduct(article);

      if (result.success) {
        updatedCount++;
        updatedArticles.push(article);
      } else {
        failedCount++;
      }

      // Flush каждые 10 товаров для сохранения изменений
      if (updatedCount % 10 === 0) {
        SpreadsheetApp.flush();
      }
    }

    logInfo(`✅ Глобальные сопоставления применены: обновлено ${updatedCount}, ошибок ${failedCount}`, null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Обновлено товаров: ${updatedCount}${failedCount > 0 ? `, ошибок: ${failedCount}` : ''}`,
      '✅ Справочник обновлен',
      5
    );

    return {
      success: true,
      updatedCount: updatedCount,
      failedCount: failedCount,
      articles: updatedArticles
    };

  } catch (error) {
    handleError(error, 'Применение глобальных сопоставлений');
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// ТЕСТОВЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Тестовая функция для ре-нормализации товара 32734
 */
function TEST_renormalize32734() {
  return renormalizeProduct('32734');
}
