# BI Publisher Template Builder - Service Layer

Complete service/utility layer for Microsoft Office Word Add-in that replicates Oracle BI Publisher Template Builder functionality.

## Overview

This directory contains five production-quality service modules that handle:
- BI Publisher server connections and REST API communication
- Word document manipulation using Office.js
- XML data parsing and XPath operations
- Template tag generation and validation
- SQL query building for Oracle databases

## Modules

### 1. bipConnection.js
**BI Publisher Server Connection Manager**

Handles all communication with Oracle BI Publisher REST API v12c+

**Key Classes:**
- `BIPConnection` - Main connection manager

**Key Methods:**
- `connect(serverUrl, username, password)` - Authenticate with BI Publisher
- `disconnect()` - Close session
- `testConnection()` - Verify server connectivity
- `getReports(folder)` - List available reports
- `getReportData(reportPath, parameters, format)` - Execute report and retrieve data
- `getDataModel(reportPath)` - Get report data structure
- `uploadTemplate(reportPath, templateFile, templateName)` - Upload template
- `downloadTemplate(reportPath, templateName)` - Download template file
- `getParameters(reportPath)` - Get report parameters
- `getLOV(parameterName)` - Get list of valid values
- `runReport(reportPath, params, format)` - Execute report
- `getServerInfo()` - Get server version and capabilities

**Features:**
- Token-based authentication with automatic refresh
- Connection state management
- Result caching for performance
- Automatic retry logic with exponential backoff
- CORS proxy support
- Session management
- Offline mode support
- Comprehensive error handling

**Usage Example:**
```javascript
import BIPConnection from './bipConnection.js';

const bip = new BIPConnection();
await bip.connect('http://bipserver:9502/xmlpserver', 'admin', 'password');

const reports = await bip.getReports();
const data = await bip.getReportData('/Reports/Sales', {year: 2023});

await bip.disconnect();
```

---

### 2. wordApi.js
**Word Document Manipulation Layer**

High-level wrapper around Office.js Word API for template building operations.

**Key Classes:**
- `WordApiHelper` - Word document manipulation interface

**Key Methods:**
- `insertTextField(fieldName, xpath, defaultValue, options)` - Insert data field
- `insertTable(columns, groupBy, sortBy)` - Insert repeating table
- `insertRepeatingGroup(xpath, content)` - Insert for-each group
- `insertConditionalRegion(condition, trueContent, falseContent)` - Insert if/else
- `insertChart(chartType, dataMapping)` - Insert chart placeholder
- `insertCrossTab(rowField, colField, dataField)` - Insert pivot table
- `insertBarcode(data, barcodeType)` - Insert barcode
- `insertPageBreak()` - Insert page break
- `insertHeaderFooter(type, content)` - Insert header/footer
- `insertPageNumber(format)` - Insert page number field
- `insertTOC(options)` - Insert table of contents
- `insertRunningTotal(field, resetGroup)` - Insert running total
- `setConditionalFormatting(field, conditions)` - Apply conditional formatting
- `insertImage(imageUrl, options)` - Insert image
- `insertHyperlink(url, text)` - Insert hyperlink
- `insertWatermark(text)` - Insert watermark
- `getDocumentContent()` - Get full document text
- `getSelectedText()` - Get current selection
- `replaceBookmark(name, content)` - Update bookmark
- `getAllFormFields()` - Enumerate all template fields
- `updateFormField(fieldId, newProperties)` - Modify field
- `insertXSLTag(tag, attributes)` - Insert raw XSL tag
- `wrapSelectionWithTag(startTag, endTag)` - Wrap selected content

**Features:**
- Proper Word.run() context management
- BI Publisher RTF tag format generation
- Error logging and reporting
- Field tracking and reference management
- Support for both ContentControls and form fields
- Comprehensive error handling with user feedback

**Usage Example:**
```javascript
import WordApiHelper from './wordApi.js';

const wordApi = new WordApiHelper();

await wordApi.insertTextField('Product Name', 'PRODUCTS/PRODUCT_NAME', 'N/A');
await wordApi.insertTable(['ID', 'NAME', 'PRICE'], 'PRODUCTS/PRODUCT');
await wordApi.insertConditionalRegion(
  'PRICE > 1000',
  'Premium Product',
  'Standard Product'
);

const errors = wordApi.getErrorLog();
```

---

### 3. xmlParser.js
**XML Data Parser and XPath Utilities**

Parses XML documents, extracts field information, and provides XPath helpers.

**Key Classes:**
- `XMLDataParser` - XML parsing and analysis

**Key Methods:**
- `parseXML(xmlString)` - Parse XML string into DOM
- `parseXSD(xsdString)` - Parse XML Schema
- `getFieldPaths(xmlTree)` - Extract all XPath expressions
- `getFieldTree(xmlTree)` - Build hierarchical field tree
- `getFieldType(xpath)` - Determine data type
- `getSampleData(xmlTree, xpath, limit)` - Get sample values
- `validateXPath(xpath)` - Validate XPath expression
- `buildXPathFromTree(treeNode)` - Generate XPath from tree
- `transformToDataFields(xmlTree)` - Convert to flat field list
- `getRepeatingElements(xmlTree)` - Find repeating groups
- `getAttributes(xpath)` - Get element attributes

**Features:**
- Full DOM parsing with error detection
- XPath validation and generation
- Namespace support
- Data type inference
- Sample data extraction
- Efficient caching
- Support for large XML files
- Repeating element detection

**Usage Example:**
```javascript
import XMLDataParser from './xmlParser.js';

const parser = new XMLDataParser();
const xmlDoc = parser.parseXML(xmlString);

const fields = parser.getFieldPaths(xmlDoc);
const tree = parser.getFieldTree(xmlDoc);
const repeating = parser.getRepeatingElements(xmlDoc);

const flatFields = parser.transformToDataFields(xmlDoc);
console.log(flatFields); // [{name, xpath, type, label, sampleValue}, ...]
```

---

### 4. templateEngine.js
**BI Publisher Template Tag Engine**

Generates and validates Oracle BI Publisher RTF template tags.

**Key Classes:**
- `TemplateEngine` - Template tag generation and validation

**Key Methods:**
- `generateFieldTag(fieldPath, options)` - Generate field tag
- `generateForEachTag(groupPath, options)` - Generate for-each opening
- `generateEndForEach()` - Generate for-each closing
- `generateIfTag(condition)` - Generate if opening
- `generateEndIf()` - Generate if closing
- `generateChooseTag(conditions)` - Generate choose/when/otherwise
- `generateSortTag(field, order)` - Generate sort specification
- `generateGroupTag(field, options)` - Generate grouping tag
- `generateRunningTotal(field, resetGroup)` - Generate running total
- `generateSubtotal(field, func)` - Generate subtotal aggregation
- `generateCrossTabTags(config)` - Generate cross-tab structure
- `generateChartXML(chartConfig)` - Generate chart XML
- `generateBarcodeTag(field, type, options)` - Generate barcode tag
- `generateDynamicColumnTags(config)` - Generate dynamic columns
- `generatePageBreakTag(condition)` - Generate conditional page break
- `generateVariableTag(name, value)` - Generate variable definition
- `generateParameterTag(name, dataType, defaultValue)` - Generate parameter
- `parseExistingTags(documentContent)` - Extract tags from document
- `validateTemplate(documentContent)` - Validate template syntax
- `getTagDescription(tag)` - Get tag documentation

**Features:**
- Full RTF template syntax support
- Tag validation and parsing
- Nested tag structure support
- Syntax error detection
- Mismatched tag detection
- Comprehensive tag reference
- Support for aggregation functions
- Chart type validation

**Supported Functions:**
SUM, AVG, MIN, MAX, COUNT, COUNTALL, STDDEV, VARIANCE, RANK, DENSE_RANK, ROW_NUMBER

**Supported Chart Types:**
column, bar, line, pie, area, bubble, scatter, stock, gauge, funnel

**Usage Example:**
```javascript
import TemplateEngine from './templateEngine.js';

const engine = new TemplateEngine();

const fieldTag = engine.generateFieldTag('PRODUCTS/NAME', {defaultValue: 'N/A'});
const forEachTag = engine.generateForEachTag('PRODUCTS/PRODUCT');
const ifTag = engine.generateIfTag('STATUS = "ACTIVE"');
const tableTag = engine.generateSubtotal('AMOUNT', 'SUM');

const validation = engine.validateTemplate(documentText);
console.log(validation); // {isValid, errors, tagCount}
```

---

### 5. sqlBuilder.js
**SQL Query Builder**

Constructs and validates SQL queries for Oracle databases.

**Key Classes:**
- `SQLBuilder` - SQL query construction and validation

**Key Methods:**
- `buildSelect(table, columns, conditions)` - Build SELECT query
- `addJoin(type, table, condition, alias)` - Add JOIN clause
- `addWhere(condition)` - Add WHERE conditions
- `addGroupBy(columns)` - Add GROUP BY clause
- `addOrderBy(columns, directions)` - Add ORDER BY clause
- `addHaving(condition)` - Add HAVING clause
- `setLimit(limit, offset)` - Set LIMIT/OFFSET
- `buildQuery()` - Compile full SQL string
- `parseQuery(sql)` - Parse existing SQL
- `validateQuery(sql)` - Validate SQL syntax
- `formatQuery(sql)` - Format/beautify SQL
- `getQueryMetadata(sql)` - Extract metadata
- `getSupportedFunctions()` - List available SQL functions
- `reset()` - Clear builder state
- `getHistory()` - Get previous queries
- `clearHistory()` - Clear query history

**Features:**
- Fluent API for query building
- SQL injection prevention
- Query length validation
- Bracket/parenthesis matching
- Multiple statement detection
- Oracle SQL dialect support
- Comprehensive validation
- Query history tracking
- Metadata extraction
- SQL beautification

**Supported JOIN Types:**
INNER, LEFT, RIGHT, FULL, CROSS

**Usage Example:**
```javascript
import SQLBuilder from './sqlBuilder.js';

const builder = new SQLBuilder();

const sql = builder
  .buildSelect('PRODUCTS', ['ID', 'NAME', 'PRICE'], 'STATUS = "ACTIVE"')
  .addJoin('INNER', 'CATEGORIES', 'PRODUCTS.CAT_ID = CATEGORIES.ID')
  .addOrderBy('NAME', 'ASC')
  .setLimit(100, 0)
  .buildQuery();

// Validate before execution
const validation = builder.validateQuery(sql);
if (validation.isValid) {
  // Execute query
}
```

---

## Installation and Import

All modules use ES6 export syntax:

```javascript
// Default exports
import BIPConnection from './services/bipConnection.js';
import WordApiHelper from './services/wordApi.js';
import XMLDataParser from './services/xmlParser.js';
import TemplateEngine from './services/templateEngine.js';
import SQLBuilder from './services/sqlBuilder.js';

// Or named exports
import { BIPConnection } from './services/bipConnection.js';
import { WordApiHelper } from './services/wordApi.js';
import { XMLDataParser } from './services/xmlParser.js';
import { TemplateEngine } from './services/templateEngine.js';
import { SQLBuilder } from './services/sqlBuilder.js';
```

## Requirements

- **Office.js** - For Word API functionality (bipPublisher-mac uses this)
- **Modern browser** - ES6 support, DOMParser, XPathEvaluator
- **Oracle BI Publisher 11g+** - For server connection

## Error Handling

All modules include comprehensive error handling:
- Input validation
- Detailed error messages
- Error logging (WordApiHelper)
- Graceful fallbacks
- Exception throwing for critical failures

## Performance Considerations

- **Caching**: BIPConnection and XMLDataParser use caching
- **Lazy Loading**: Data loaded on-demand
- **Memory**: Efficient for large XML documents
- **Network**: Automatic retry logic, token refresh
- **Query Length**: Oracle maximum enforced

## Security Features

- **SQL Injection Prevention**: Input validation, parameter binding
- **XPath Validation**: Expression checking
- **XML Parsing**: Error detection, namespace handling
- **Authentication**: Secure token handling
- **Password Handling**: Not stored in memory longer than necessary

## Documentation

Each method includes JSDoc comments with:
- Full parameter descriptions
- Return type information
- Error conditions
- Usage examples
- Related methods

All classes are documented with detailed descriptions of functionality and usage patterns.

## License

Part of Oracle BI Publisher Template Builder add-in for Word
