# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script** project for automated product image processing and e-commerce integration. The system manages product images from InSales, parses images from competitor suppliers, and uses AI for SEO optimization (alt-tags and filenames).

**Key Technologies:**
- Google Apps Script (JavaScript runtime for Google Sheets)
- InSales API integration
- OpenAI GPT-4 Vision API
- Web scraping (UrlFetchApp)
- Google Script Properties for secure credential storage

## Development Environment

### Deployment Method
This project uses [clasp](https://github.com/google/clasp) for deployment:
```bash
# Push changes to Google Apps Script
clasp push

# Pull changes from remote
clasp pull

# Open project in browser
clasp open
```

### Testing
Since this is Google Apps Script, there is no traditional test suite. Testing is done via:
1. Running functions directly in Google Apps Script IDE (`Script Editor` → `Run`)
2. Using the Google Sheets menu: `🖼️ Фото` → `⚙️ Проверить настройки API`
3. Test functions exist in each module (e.g., `testConfigModule()`)

### Viewing Logs
```javascript
// In Google Apps Script IDE
View → Execution log

// Or use Stackdriver Logging
View → Stackdriver Logging
```

## Architecture

### Module Structure

The codebase follows a **hierarchical dependency model** with numbered prefixes indicating load order:

```
00_config.gs.js          → Configuration, constants, API settings
01_shared_utilities.js   → Logging, error handling, Sheets helpers
02_data_manager.js       → Google Sheets CRUD operations
03_insales_api.js        → InSales platform integration
04_image_processing.js   → OpenAI GPT-4 Vision integration
05_supplier_parsing.gs.js→ Web scraping from suppliers
99_menu.js               → UI menu and workflow orchestration
*.html                   → Dialog UIs (HtmlService)
```

**Dependency Rule:** Higher-numbered modules can depend on lower-numbered modules, never the reverse.

### Key Data Model

The system operates on a single Google Sheet with this structure:

| Column | Name | Purpose |
|--------|------|---------|
| A | CHECKBOX | User selection for batch operations |
| B | ARTICLE | Product SKU (primary key) |
| C | INSALES_ID | Product ID in InSales |
| D | PRODUCT_NAME | Full product title |
| E | ORIGINAL_IMAGES | URLs from InSales (newline-separated) |
| F | SUPPLIER_IMAGES | Parsed from suppliers (newline-separated) |
| G | ADDITIONAL_IMAGES | User-added images |
| H | PROCESSED_IMAGES | Enhanced/uploaded images |
| I | ALT_TAGS | SEO alt-text (newline-separated) |
| J | SEO_FILENAMES | SEO filenames (newline-separated) |
| K | PROCESSING_STATUS | Current processing state |
| L | INSALES_STATUS | Upload status |

Access columns via constants: `IMAGES_COLUMNS.ARTICLE`, `IMAGES_COLUMNS.ALT_TAGS`, etc.

### Main Workflows

#### 1. Loading Products from InSales
```javascript
// Entry: 99_menu.js → loadProductsFromInSalesMenu()
// Flow:
03_insales_api → loadProductsFromInSales()
  → loadCatalogStructure()        // Get categories
  → loadProductsFromCategory()    // Paginated fetch (250/page)
  → loadProductVariants()         // Extract SKUs from variants
  → syncProductData()             // Write to Sheets
    → 02_data_manager → writeProductData()
```

**Important:** Always get article numbers from `variant.sku`, not `product.sku`.

#### 2. Parsing Supplier Images
```javascript
// Entry: 99_menu.js → showSupplierParsingDialog()
// Flow:
05_supplier_parsing → runSupplierParsing()
  → executeSupplierParser()       // For each enabled supplier
    → parse<Supplier>Images()     // Scrape product pages
  → saveSelectedImages()          // Store URLs in sheet
```

#### 3. AI Image Analysis
```javascript
// Entry: 99_menu.js → showImageSelectionForProcessing()
// Flow:
04_image_processing → processSelectedImages()
  → analyzeImageSimple()
    → callOpenAIVision()          // GPT-4 Vision API
  → generateAltTag()              // Max 125 chars, no stop words
  → generateSeoFilename()         // Latin chars only, max 80 chars
  → 02_data_manager → Update sheet
```

## Configuration

### API Keys Setup

All credentials are stored in **Script Properties** (never in code):

```javascript
// View/set via Apps Script IDE:
Project Settings → Script Properties

// Or programmatically:
const properties = PropertiesService.getScriptProperties();
properties.setProperty(SCRIPT_PROPERTIES_KEYS.INSALES_API_KEY, 'your_key');
```

**Required Keys:**
- `insalesApiKey` - InSales API key
- `insalesPassword` - InSales API password
- `insalesShop` - InSales shop subdomain (e.g., "myshop")

**Optional Keys:**
- `openaiApiKey` - For image analysis
- `replicateToken` - For image upscaling (stub)
- `tinypngKey` - For compression (stub)
- `imgbbKey` - For image hosting (stub)
- `telegramToken`, `telegramChatId` - For notifications

### Validation

Always validate configuration before operations:
```javascript
const validation = validateConfig();
if (!validation.isValid) {
  console.log('Errors:', validation.errors);
}
```

## Common Patterns

### Error Handling
```javascript
try {
  // Your code
} catch (error) {
  handleError(error, 'Context description', {
    additionalInfo: 'for debugging'
  });
}
```

### Logging
```javascript
logInfo('Operation started', {data: 'context'});
logWarning('Non-critical issue detected');
logError('Error occurred', errorObject);
logCritical('System failure'); // Sends Telegram notification
```

### Reading from Sheets
```javascript
const sheet = getImagesSheet(); // Safe accessor with validation
const data = sheet.getDataRange().getValues();
const headers = data[0];
const rows = data.slice(1);

// Or use data manager:
const products = readSelectedProducts(); // Gets checked items
```

### Writing to Sheets
```javascript
// Single product:
writeProductData({
  article: 'SKU123',
  productName: 'Product Name',
  originalImages: 'url1\nurl2',
  processingStatus: STATUS_VALUES.PROCESSING.COMPLETED
});

// Update status only:
updateProductStatus('SKU123', STATUS_VALUES.INSALES.SENT);
```

### API Requests with Retry Logic
```javascript
const response = makeInsalesRequest(
  'GET',
  '/admin/products.json',
  null,
  {per_page: 250, page: 1}
);
// Automatically handles: auth, pagination, rate limits (429), retries
```

## Important Implementation Details

### SKU Extraction
**Always** extract article numbers from product variants, not the product itself:
```javascript
// CORRECT:
const article = variant.sku;

// WRONG:
const article = product.sku; // Often empty or incorrect
```

### Image URL Validation
```javascript
// Always validate before storing:
if (url && url.startsWith('http')) {
  // Safe to store
}
```

### Multi-line Data Storage
Images, alt-tags, and filenames are stored as newline-separated strings:
```javascript
const imageUrls = ['url1', 'url2', 'url3'];
const storedValue = imageUrls.join('\n');

// Reading:
const imageArray = cellValue.split('\n').filter(url => url.trim());
```

### Rate Limiting
```javascript
// For supplier parsing, respect rate limits:
Utilities.sleep(1000); // 1 second between requests

// For API calls, use exponential backoff on 429:
if (responseCode === 429) {
  Utilities.sleep(Math.pow(2, attempt) * 1000);
}
```

### Transliteration for SEO
```javascript
// Russian to Latin for filenames:
const latinFilename = transliterate(russianText);
// Removes special chars, max 80 chars, lowercase
```

### Batch Processing Limits
- Process max **50 products per batch** to avoid Google Apps Script 6-minute execution timeout
- Use `SpreadsheetApp.flush()` periodically to save changes

## Adding New Features

### Adding a New Supplier Parser
1. Add to `SUPPLIERS_CONFIG` in [05_supplier_parsing.gs.js](05_supplier_parsing.gs.js):
```javascript
NEW_SUPPLIER: {
  name: 'Supplier Name',
  enabled: true,
  searchUrl: (article) => `https://example.com/search?q=${article}`,
  selectors: {
    images: 'img.product-image',
    excludePatterns: ['logo', 'banner']
  }
}
```

2. Implement parser function:
```javascript
function parseNewSupplierImages(url) {
  try {
    const html = UrlFetchApp.fetch(url).getContentText();
    // Extract images using regex or XML parsing
    return imageUrls;
  } catch (error) {
    handleError(error, 'Parse New Supplier');
    return [];
  }
}
```

3. Add case to `executeSupplierParser()` switch statement

### Adding New Columns
1. Update `IMAGES_COLUMNS` in [00_config.gs.js](00_config.gs.js:62)
2. Modify `setupHeaders()` in [02_data_manager.js](02_data_manager.js)
3. Update read/write logic in data manager functions
4. Adjust column widths in `setupColumnWidths()`

## Debugging Tips

### Common Issues

**"Лист не найден" (Sheet not found):**
- Ensure sheet name matches `SHEET_NAMES.IMAGES` exactly
- Run `createImagesSheet()` to initialize structure

**"Authorization required":**
- Check Script Properties contain all required API keys
- Validate with `validateConfig()`

**API Rate Limits (429):**
- Already handled by `makeInsalesRequest()` with exponential backoff
- For custom requests, add `Utilities.sleep()` between calls

**Execution Timeout:**
- Reduce batch size (max 50 items)
- Split workflow into smaller functions
- Use Google Apps Script triggers for long operations

**Empty SKUs:**
- Check that `loadProductVariants()` is called
- Verify variants exist in InSales product data

## Google Apps Script Limitations

- **6-minute execution timeout** per function call
- **No native async/await** (use sequential processing)
- **Single-threaded** execution model
- **URLFetch quotas**: 20,000 calls/day
- **Cannot render JavaScript** during web scraping
- **No file system access** (must upload to external services)

## Code Style

- **Language:** Russian comments and variable names (mixed with English APIs)
- **Logging:** Always use emoji prefixes (✅ success, ❌ error, ⚠️ warning, 📝 info)
- **Error messages:** Russian language for end-user display
- **Constants:** UPPERCASE_SNAKE_CASE
- **Functions:** camelCase
- **Status values:** Use constants from `STATUS_VALUES` object

## Resources

- **InSales API Docs:** https://www.insales.ru/collection/api
- **Google Apps Script Reference:** https://developers.google.com/apps-script/reference
- **OpenAI Vision API:** https://platform.openai.com/docs/guides/vision
- **clasp Documentation:** https://github.com/google/clasp
