/**
 * Word Document Manipulation Layer
 * Provides high-level API for inserting BI Publisher template elements into Word documents
 * Uses Office.js Word API with proper context management and error handling
 *
 * @module wordApi
 * @requires Office.js (Word API)
 */

/**
 * WordApiHelper class wraps Office.js Word API for template building
 * Handles insertion of BI Publisher-specific elements and formatting
 */
class WordApiHelper {
  /**
   * Creates a new Word API helper instance
   */
  constructor() {
    this.lastOperation = null;
    this.bookmarks = new Map();
    this.formFields = new Map();
    this.errorLog = [];
  }

  /**
   * Insert a form field with BI Publisher XPath tag
   * @async
   * @param {string} fieldName - Display name for the field
   * @param {string} xpath - XPath expression for the data source
   * @param {string} [defaultValue=''] - Default value if data not found
   * @param {Object} [options={}] - Additional field options
   * @param {string} [options.format=''] - Number/date format (optional)
   * @param {string} [options.placeholder=''] - Placeholder text
   * @returns {Promise<Object>} Field insertion result
   * @throws {Error} If context operation fails
   *
   * @example
   * await wordApi.insertTextField('Product Name', 'PRODUCTS/PRODUCT_NAME', '', {
   *   placeholder: 'Enter product name'
   * });
   */
  async insertTextField(fieldName, xpath, defaultValue = '', options = {}) {
    try {
      return await Word.run(async (context) => {
        const field = this._generateFieldTag(fieldName, xpath, defaultValue, options);

        // Insert field into selection or end of document
        const selection = context.document.body.getRange('End');
        const fieldParagraph = selection.insertParagraph(
          field,
          Word.InsertLocation.before
        );

        // Set formatting if specified
        if (options.format) {
          fieldParagraph.font.name = 'Courier New';
          fieldParagraph.font.size = 10;
        }

        // Store field reference
        this._storeFieldReference(fieldName, xpath, fieldParagraph);

        await context.sync();

        return {
          success: true,
          fieldName,
          xpath,
          message: `Field '${fieldName}' inserted successfully`
        };
      });
    } catch (error) {
      this._logError(`insertTextField for ${fieldName}`, error);
      throw new Error(`Failed to insert field: ${error.message}`);
    }
  }

  /**
   * Insert a data-driven table with repeating rows
   * @async
   * @param {Array<string>} columns - Column field names/XPaths
   * @param {string} [groupBy=''] - XPath for row grouping
   * @param {Array<Object>} [sortBy=[]] - Sort specifications
   * @returns {Promise<Object>} Table insertion result
   * @throws {Error} If table creation fails
   *
   * @example
   * await wordApi.insertTable(
   *   ['PRODUCT_ID', 'PRODUCT_NAME', 'PRICE'],
   *   'PRODUCTS/PRODUCT',
   *   [{field: 'PRODUCT_NAME', order: 'asc'}]
   * );
   */
  async insertTable(columns, groupBy = '', sortBy = []) {
    try {
      return await Word.run(async (context) => {
        const rows = columns.length;
        const cols = 1;

        // Insert table with header row
        const selection = context.document.body.getRange('End');
        const table = selection.insertTable(rows + 1, 1, Word.InsertLocation.before);

        // Add header row with column names
        columns.forEach((col, idx) => {
          if (idx > 0) {
            table.addColumns(Word.InsertLocation.end, 1, [['']]);
          }
          const cell = table.cell(0, idx);
          cell.body.clear();
          cell.body.insertParagraph(col, Word.InsertLocation.replace);
          cell.body.paragraphs[0].font.bold = true;
        });

        // Add data row with field tags
        const dataRow = table.insertRows(Word.InsertLocation.end, 1);
        columns.forEach((col, idx) => {
          const cellText = `<?${this._generateXPathTag(col)}?>`;
          const cell = table.cell(1, idx);
          cell.body.clear();
          cell.body.insertParagraph(cellText, Word.InsertLocation.replace);
        });

        // Apply repeating group tags if groupBy specified
        if (groupBy) {
          const groupTag = `<?for-each:${groupBy}?>`;
          const tableRange = table.getRange('Whole');
          tableRange.insertParagraph(groupTag, Word.InsertLocation.before);
          tableRange.insertParagraph('<?end for-each?>', Word.InsertLocation.after);
        }

        // Apply sorting if specified
        if (sortBy && sortBy.length > 0) {
          const sortTags = sortBy.map(sort =>
            `<?sort:${sort.field}:${sort.order || 'asc'}?>`
          ).join('\n');
          tableRange.insertParagraph(sortTags, Word.InsertLocation.before);
        }

        await context.sync();

        return {
          success: true,
          columns,
          groupBy,
          message: `Table with ${columns.length} columns inserted successfully`
        };
      });
    } catch (error) {
      this._logError('insertTable', error);
      throw new Error(`Failed to insert table: ${error.message}`);
    }
  }

  /**
   * Insert a repeating group (for-each loop) around content
   * @async
   * @param {string} xpath - XPath expression for repeating element
   * @param {string} [content='<content>'] - Content to repeat
   * @returns {Promise<Object>} Group insertion result
   * @throws {Error} If group insertion fails
   *
   * @example
   * await wordApi.insertRepeatingGroup('PRODUCTS/PRODUCT', 'Product Details');
   */
  async insertRepeatingGroup(xpath, content = '<content>') {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Insert opening tag
        const openTag = `<?for-each:${xpath}?>`;
        selection.insertParagraph(openTag, Word.InsertLocation.before);

        // Insert content
        selection.insertParagraph(content, Word.InsertLocation.after);

        // Insert closing tag
        selection.insertParagraph('<?end for-each?>', Word.InsertLocation.after);

        await context.sync();

        return {
          success: true,
          xpath,
          message: `Repeating group '${xpath}' inserted successfully`
        };
      });
    } catch (error) {
      this._logError('insertRepeatingGroup', error);
      throw new Error(`Failed to insert repeating group: ${error.message}`);
    }
  }

  /**
   * Insert conditional region (if/choose/when structure)
   * @async
   * @param {string} condition - XPath condition expression
   * @param {string} [trueContent='<true>'] - Content if condition is true
   * @param {string} [falseContent='<false>'] - Content if condition is false
   * @returns {Promise<Object>} Conditional insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertConditionalRegion(
   *   'PRICE > 1000',
   *   'Premium Product',
   *   'Standard Product'
   * );
   */
  async insertConditionalRegion(
    condition,
    trueContent = '<true>',
    falseContent = '<false>'
  ) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Insert if tag
        selection.insertParagraph(`<?if:${condition}?>`, Word.InsertLocation.before);

        // Insert true content
        selection.insertParagraph(trueContent, Word.InsertLocation.after);

        // Insert else/otherwise
        if (falseContent && falseContent !== '<false>') {
          selection.insertParagraph('<?else?>', Word.InsertLocation.after);
          selection.insertParagraph(falseContent, Word.InsertLocation.after);
        }

        // Insert closing tag
        selection.insertParagraph('<?end if?>', Word.InsertLocation.after);

        await context.sync();

        return {
          success: true,
          condition,
          message: 'Conditional region inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertConditionalRegion', error);
      throw new Error(`Failed to insert conditional region: ${error.message}`);
    }
  }

  /**
   * Insert chart placeholder with data mapping
   * @async
   * @param {string} chartType - Chart type (column, bar, line, pie, area)
   * @param {Object} dataMapping - Mapping of chart axes to XPath expressions
   * @param {string} [dataMapping.categories=''] - X-axis/category XPath
   * @param {string} [dataMapping.values=''] - Y-axis/value XPath
   * @param {string} [dataMapping.series=''] - Series/legend XPath
   * @returns {Promise<Object>} Chart insertion result
   * @throws {Error} If chart insertion fails
   *
   * @example
   * await wordApi.insertChart('column', {
   *   categories: 'MONTHS/MONTH_NAME',
   *   values: 'MONTHS/SALES_AMOUNT',
   *   series: 'REGIONS/REGION_NAME'
   * });
   */
  async insertChart(chartType, dataMapping) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Build chart XML tag
        const chartTag = this._generateChartTag(chartType, dataMapping);
        selection.insertParagraph(chartTag, Word.InsertLocation.before);

        // Add placeholder text
        const placeholderText = `[Chart: ${chartType} - ${dataMapping.values || 'Data'}]`;
        const para = selection.insertParagraph(placeholderText, Word.InsertLocation.after);
        para.font.italic = true;
        para.font.color = '#999999';

        await context.sync();

        return {
          success: true,
          chartType,
          dataMapping,
          message: 'Chart placeholder inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertChart', error);
      throw new Error(`Failed to insert chart: ${error.message}`);
    }
  }

  /**
   * Insert cross-tabulation (pivot table) structure
   * @async
   * @param {string} rowField - XPath for row dimension
   * @param {string} colField - XPath for column dimension
   * @param {string} dataField - XPath for data values
   * @param {string} [aggregation='sum'] - Aggregation function
   * @returns {Promise<Object>} CrossTab insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertCrossTab('REGIONS/REGION', 'MONTHS/MONTH', 'SALES/AMOUNT', 'sum');
   */
  async insertCrossTab(rowField, colField, dataField, aggregation = 'sum') {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Insert cross-tab tag
        const crossTabTag = `<?crosstab:rows='${rowField}' cols='${colField}' data='${dataField}' agg='${aggregation}'?>`;
        selection.insertParagraph(crossTabTag, Word.InsertLocation.before);

        // Insert placeholder
        const placeholder = `[Cross-tabulation: ${rowField} × ${colField}]`;
        const para = selection.insertParagraph(placeholder, Word.InsertLocation.after);
        para.font.italic = true;

        await context.sync();

        return {
          success: true,
          rowField,
          colField,
          dataField,
          message: 'Cross-tab inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertCrossTab', error);
      throw new Error(`Failed to insert cross-tab: ${error.message}`);
    }
  }

  /**
   * Insert barcode image placeholder with data reference
   * @async
   * @param {string} data - XPath or fixed data for barcode
   * @param {string} barcodeType - Barcode type (code39, code128, ean13, qrcode)
   * @param {Object} [options={}] - Additional barcode options
   * @returns {Promise<Object>} Barcode insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertBarcode('PRODUCTS/SKU', 'code128');
   */
  async insertBarcode(data, barcodeType, options = {}) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Insert barcode tag
        const barcodeTag = `<?barcode:data='${data}' type='${barcodeType}'?>`;
        selection.insertParagraph(barcodeTag, Word.InsertLocation.before);

        // Insert placeholder
        const placeholder = `[${barcodeType.toUpperCase()} Barcode: ${data}]`;
        const para = selection.insertParagraph(placeholder, Word.InsertLocation.after);
        para.font.name = 'Courier New';
        para.font.size = 9;

        await context.sync();

        return {
          success: true,
          data,
          barcodeType,
          message: 'Barcode placeholder inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertBarcode', error);
      throw new Error(`Failed to insert barcode: ${error.message}`);
    }
  }

  /**
   * Insert page break
   * @async
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertPageBreak();
   */
  async insertPageBreak() {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');
        selection.insertBreak(Word.BreakType.page, Word.InsertLocation.before);

        await context.sync();

        return {
          success: true,
          message: 'Page break inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertPageBreak', error);
      throw new Error(`Failed to insert page break: ${error.message}`);
    }
  }

  /**
   * Insert header or footer content
   * @async
   * @param {string} type - Header or Footer ('header' or 'footer')
   * @param {string} content - Content to insert
   * @param {Object} [options={}] - Formatting options
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertHeaderFooter('header', 'Sales Report - 2023');
   */
  async insertHeaderFooter(type, content, options = {}) {
    try {
      return await Word.run(async (context) => {
        const isHeader = type.toLowerCase() === 'header';
        const sections = context.document.sections;
        sections.load('items');

        await context.sync();

        sections.items.forEach((section) => {
          const headerOrFooter = isHeader ? section.header : section.footer;
          headerOrFooter.clear();
          const para = headerOrFooter.insertParagraph(
            content,
            Word.InsertLocation.start
          );

          // Apply formatting if specified
          if (options.fontSize) para.font.size = options.fontSize;
          if (options.alignment) para.alignment = options.alignment;
          if (options.italic) para.font.italic = true;
        });

        await context.sync();

        return {
          success: true,
          type,
          message: `${type} inserted successfully`
        };
      });
    } catch (error) {
      this._logError(`insertHeaderFooter(${type})`, error);
      throw new Error(`Failed to insert ${type}: ${error.message}`);
    }
  }

  /**
   * Insert page number field
   * @async
   * @param {string} [format='#'] - Page number format (#=current, ##=both, \Page # of \numpages=both)
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertPageNumber('# of \numpages');
   */
  async insertPageNumber(format = '#') {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        const fieldCode = format.includes('\\')
          ? format
          : `PAGE \\* MERGEFORMAT`;

        // Insert as complex field using field code
        const para = selection.insertParagraph('', Word.InsertLocation.before);
        const run = para.insertField(fieldCode, Word.InsertLocation.start);

        await context.sync();

        return {
          success: true,
          format,
          message: 'Page number field inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertPageNumber', error);
      throw new Error(`Failed to insert page number: ${error.message}`);
    }
  }

  /**
   * Insert table of contents
   * @async
   * @param {Object} [options={}] - TOC options
   * @param {number} [options.depth=3] - Heading depth to include
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertTOC({depth: 2});
   */
  async insertTOC(options = {}) {
    try {
      return await Word.run(async (context) => {
        const depth = options.depth || 3;
        const selection = context.document.body.getRange('End');

        // Insert TOC field code
        const fieldCode = `TOC \\o "1-${depth}" \\h \\z \\u`;
        selection.insertParagraph('Table of Contents', Word.InsertLocation.before);
        const para = selection.insertParagraph('', Word.InsertLocation.after);
        para.insertField(fieldCode, Word.InsertLocation.start);

        await context.sync();

        return {
          success: true,
          depth,
          message: 'Table of contents inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertTOC', error);
      throw new Error(`Failed to insert TOC: ${error.message}`);
    }
  }

  /**
   * Insert running total field
   * @async
   * @param {string} field - Field XPath to sum
   * @param {string} [resetGroup=''] - Group level to reset sum
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertRunningTotal('SALES/AMOUNT', 'REGIONS/REGION');
   */
  async insertRunningTotal(field, resetGroup = '') {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        const tag = resetGroup
          ? `<?sum:${field}:reset='${resetGroup}'?>`
          : `<?sum:${field}?>`;

        selection.insertParagraph(tag, Word.InsertLocation.before);

        await context.sync();

        return {
          success: true,
          field,
          resetGroup,
          message: 'Running total field inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertRunningTotal', error);
      throw new Error(`Failed to insert running total: ${error.message}`);
    }
  }

  /**
   * Set conditional formatting on a field
   * @async
   * @param {string} field - Field name/XPath
   * @param {Array<Object>} conditions - Condition rules
   * @param {string} [conditions.condition=''] - XPath condition
   * @param {Object} [conditions.format={}] - Format properties
   * @returns {Promise<Object>} Result
   * @throws {Error} If operation fails
   *
   * @example
   * await wordApi.setConditionalFormatting('PRICE', [
   *   {condition: 'PRICE > 1000', format: {bold: true, color: 'FF0000'}},
   *   {condition: 'PRICE <= 1000', format: {italic: true}}
   * ]);
   */
  async setConditionalFormatting(field, conditions) {
    try {
      return await Word.run(async (context) => {
        // This is primarily handled in template tags
        // Insert formatting tag before field
        const selection = context.document.body.getRange('End');

        conditions.forEach(cond => {
          const formatTag = `<?format:field='${field}' condition='${cond.condition}'?>`;
          selection.insertParagraph(formatTag, Word.InsertLocation.before);
        });

        await context.sync();

        return {
          success: true,
          field,
          conditionCount: conditions.length,
          message: 'Conditional formatting configured successfully'
        };
      });
    } catch (error) {
      this._logError('setConditionalFormatting', error);
      throw new Error(`Failed to set conditional formatting: ${error.message}`);
    }
  }

  /**
   * Insert image from URL
   * @async
   * @param {string} imageUrl - URL of image to insert
   * @param {Object} [options={}] - Image options
   * @param {number} [options.width=200] - Width in pixels
   * @param {number} [options.height=150] - Height in pixels
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertImage('https://example.com/logo.png', {width: 300, height: 200});
   */
  async insertImage(imageUrl, options = {}) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');
        const width = options.width || 200;
        const height = options.height || 150;

        // Insert image from URL
        const para = selection.insertParagraph('', Word.InsertLocation.before);
        const run = para.insertInlinePictureFromBase64(
          imageUrl,
          Word.InsertLocation.start
        );

        run.width = width;
        run.height = height;

        await context.sync();

        return {
          success: true,
          url: imageUrl,
          width,
          height,
          message: 'Image inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertImage', error);
      throw new Error(`Failed to insert image: ${error.message}`);
    }
  }

  /**
   * Insert hyperlink
   * @async
   * @param {string} url - Link target URL
   * @param {string} [text='Link'] - Display text
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertHyperlink('https://example.com', 'Click here');
   */
  async insertHyperlink(url, text = 'Link') {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');
        const para = selection.insertParagraph(text, Word.InsertLocation.before);
        const run = para.insertHyperlink(url, text, Word.InsertLocation.start);

        run.font.color = '#0563C1';
        run.font.underline = true;

        await context.sync();

        return {
          success: true,
          url,
          text,
          message: 'Hyperlink inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertHyperlink', error);
      throw new Error(`Failed to insert hyperlink: ${error.message}`);
    }
  }

  /**
   * Insert watermark text
   * @async
   * @param {string} text - Watermark text
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertWatermark('CONFIDENTIAL');
   */
  async insertWatermark(text) {
    try {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load('items');

        await context.sync();

        sections.items.forEach((section) => {
          const watermarkTag = `<?watermark:'${text}'?>`;
          section.header.clear();
          section.header.insertParagraph(watermarkTag, Word.InsertLocation.start);
        });

        await context.sync();

        return {
          success: true,
          text,
          message: 'Watermark inserted successfully'
        };
      });
    } catch (error) {
      this._logError('insertWatermark', error);
      throw new Error(`Failed to insert watermark: ${error.message}`);
    }
  }

  /**
   * Get full document content
   * @async
   * @returns {Promise<string>} Document content as text
   * @throws {Error} If retrieval fails
   *
   * @example
   * const content = await wordApi.getDocumentContent();
   */
  async getDocumentContent() {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');

        await context.sync();

        return body.text;
      });
    } catch (error) {
      this._logError('getDocumentContent', error);
      throw new Error(`Failed to get document content: ${error.message}`);
    }
  }

  /**
   * Get currently selected text
   * @async
   * @returns {Promise<string>} Selected text
   * @throws {Error} If retrieval fails
   *
   * @example
   * const selected = await wordApi.getSelectedText();
   */
  async getSelectedText() {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');

        await context.sync();

        return selection.text;
      });
    } catch (error) {
      this._logError('getSelectedText', error);
      throw new Error(`Failed to get selected text: ${error.message}`);
    }
  }

  /**
   * Replace bookmark content
   * @async
   * @param {string} name - Bookmark name
   * @param {string} content - New content
   * @returns {Promise<Object>} Replace result
   * @throws {Error} If operation fails
   *
   * @example
   * await wordApi.replaceBookmark('TitleBookmark', 'New Title');
   */
  async replaceBookmark(name, content) {
    try {
      return await Word.run(async (context) => {
        const range = context.document.bookmarks.getItem(name).getRange();
        range.clear();
        range.insertParagraph(content, Word.InsertLocation.start);

        await context.sync();

        return {
          success: true,
          bookmark: name,
          message: `Bookmark '${name}' replaced successfully`
        };
      });
    } catch (error) {
      this._logError(`replaceBookmark(${name})`, error);
      throw new Error(`Failed to replace bookmark: ${error.message}`);
    }
  }

  /**
   * Get all BI Publisher form fields in document
   * @async
   * @returns {Promise<Array>} Array of field definitions
   * @throws {Error} If retrieval fails
   *
   * @example
   * const fields = await wordApi.getAllFormFields();
   */
  async getAllFormFields() {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.load('paragraphs/items');

        await context.sync();

        const fields = [];
        body.paragraphs.items.forEach((para, idx) => {
          const match = para.text.match(/<?(\w+):([^>]+)?>/);
          if (match) {
            fields.push({
              id: idx,
              fieldType: match[1],
              fieldValue: match[2] || '',
              fullTag: match[0],
              paragraph: para
            });
          }
        });

        return fields;
      });
    } catch (error) {
      this._logError('getAllFormFields', error);
      throw new Error(`Failed to get form fields: ${error.message}`);
    }
  }

  /**
   * Update a form field's properties
   * @async
   * @param {string} fieldId - Field identifier
   * @param {Object} newProperties - New property values
   * @returns {Promise<Object>} Update result
   * @throws {Error} If update fails
   *
   * @example
   * await wordApi.updateFormField('field_0', {value: 'new_value'});
   */
  async updateFormField(fieldId, newProperties) {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.load('paragraphs/items');

        await context.sync();

        const para = body.paragraphs.items[parseInt(fieldId)];
        if (para) {
          const currentTag = para.text;
          let newTag = currentTag;

          // Update tag with new properties
          if (newProperties.value) {
            newTag = newTag.replace(
              /<?([^:]+):([^>]*)>/,
              `<?$1:${newProperties.value}?>`
            );
          }

          para.clear();
          para.insertParagraph(newTag, Word.InsertLocation.start);
        }

        await context.sync();

        return {
          success: true,
          fieldId,
          message: 'Field updated successfully'
        };
      });
    } catch (error) {
      this._logError(`updateFormField(${fieldId})`, error);
      throw new Error(`Failed to update form field: ${error.message}`);
    }
  }

  /**
   * Insert raw XSL/XPath tag
   * @async
   * @param {string} tag - Tag name
   * @param {Object} [attributes={}] - Tag attributes
   * @returns {Promise<Object>} Insertion result
   * @throws {Error} If insertion fails
   *
   * @example
   * await wordApi.insertXSLTag('value-of', {select: 'PRODUCT_NAME'});
   */
  async insertXSLTag(tag, attributes = {}) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.body.getRange('End');

        // Build tag string
        const attrStr = Object.entries(attributes)
          .map(([key, val]) => `${key}='${val}'`)
          .join(' ');

        const tagStr = attrStr ? `<?${tag}:${attrStr}?>` : `<?${tag}?>`;
        selection.insertParagraph(tagStr, Word.InsertLocation.before);

        await context.sync();

        return {
          success: true,
          tag,
          attributes,
          message: `XSL tag '${tag}' inserted successfully`
        };
      });
    } catch (error) {
      this._logError('insertXSLTag', error);
      throw new Error(`Failed to insert XSL tag: ${error.message}`);
    }
  }

  /**
   * Wrap current selection with tags
   * @async
   * @param {string} startTag - Opening tag
   * @param {string} endTag - Closing tag
   * @returns {Promise<Object>} Wrap result
   * @throws {Error} If wrapping fails
   *
   * @example
   * await wordApi.wrapSelectionWithTag('<?for-each:PRODUCTS?>', '<?end for-each?>');
   */
  async wrapSelectionWithTag(startTag, endTag) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertParagraph(endTag, Word.InsertLocation.after);
        selection.insertParagraph(startTag, Word.InsertLocation.before);

        await context.sync();

        return {
          success: true,
          startTag,
          endTag,
          message: 'Selection wrapped successfully'
        };
      });
    } catch (error) {
      this._logError('wrapSelectionWithTag', error);
      throw new Error(`Failed to wrap selection: ${error.message}`);
    }
  }

  /**
   * Get error log
   * @returns {Array} Error log entries
   *
   * @example
   * const errors = wordApi.getErrorLog();
   */
  getErrorLog() {
    return [...this.errorLog];
  }

  /**
   * Clear error log
   * @returns {void}
   */
  clearErrorLog() {
    this.errorLog = [];
  }

  // ============= Private Helper Methods =============

  /**
   * Generate BI Publisher field tag
   * @private
   * @param {string} fieldName
   * @param {string} xpath
   * @param {string} defaultValue
   * @param {Object} options
   * @returns {string} Field tag
   */
  _generateFieldTag(fieldName, xpath, defaultValue, options) {
    let tag = `<?${xpath}`;

    if (defaultValue) {
      tag += `:default='${defaultValue}'`;
    }

    if (options.format) {
      tag += `:format='${options.format}'`;
    }

    tag += '?>`;';

    if (options.placeholder) {
      tag += ` [${options.placeholder}]`;
    }

    return tag;
  }

  /**
   * Generate XPath tag from field
   * @private
   * @param {string} field
   * @returns {string} XPath tag
   */
  _generateXPathTag(field) {
    return `${field}`;
  }

  /**
   * Generate chart tag
   * @private
   * @param {string} chartType
   * @param {Object} dataMapping
   * @returns {string} Chart tag
   */
  _generateChartTag(chartType, dataMapping) {
    return `<?chart:type='${chartType}' categories='${dataMapping.categories || ''}' values='${dataMapping.values || ''}' series='${dataMapping.series || ''}'?>`;
  }

  /**
   * Store field reference for tracking
   * @private
   * @param {string} fieldName
   * @param {string} xpath
   * @param {Object} paragraph
   * @returns {void}
   */
  _storeFieldReference(fieldName, xpath, paragraph) {
    this.formFields.set(fieldName, {
      fieldName,
      xpath,
      paragraph,
      createdAt: new Date()
    });
  }

  /**
   * Log error
   * @private
   * @param {string} operation
   * @param {Error} error
   * @returns {void}
   */
  _logError(operation, error) {
    this.errorLog.push({
      operation,
      error: error.message,
      timestamp: new Date()
    });

    console.error(`WordApiHelper.${operation}: ${error.message}`);
  }
}

// Export
export default WordApiHelper;
export { WordApiHelper };
