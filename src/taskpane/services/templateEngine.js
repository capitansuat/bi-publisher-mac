/**
 * BI Publisher Template Tag Engine
 * Generates Oracle BI Publisher RTF template tags and validates template syntax
 * Supports all standard BI Publisher template features including grouping, sorting,
 * conditionals, aggregations, and cross-tabs
 *
 * @module templateEngine
 */

/**
 * TemplateEngine class generates and parses BI Publisher template tags
 * Implements full RTF template syntax for template building
 */
class TemplateEngine {
  /**
   * Creates a new template engine instance
   */
  constructor() {
    this.tags = new Map();
    this.validFunctions = [
      'SUM', 'AVG', 'MIN', 'MAX', 'COUNT', 'COUNTALL',
      'STDDEV', 'VARIANCE', 'RANK', 'DENSE_RANK', 'ROW_NUMBER'
    ];
    this.validChartTypes = [
      'column', 'bar', 'line', 'pie', 'area', 'bubble',
      'scatter', 'stock', 'gauge', 'funnel'
    ];
    this.tagReference = this._buildTagReference();
  }

  /**
   * Generate field tag for BI Publisher template
   * @param {string} fieldPath - XPath or field expression
   * @param {Object} [options={}] - Field options
   * @param {string} [options.defaultValue=''] - Default if data missing
   * @param {string} [options.format=''] - Number/date format
   * @param {boolean} [options.multiline=false] - Multi-line field
   * @param {string} [options.mask=''] - Input mask/validation
   * @returns {string} Generated field tag
   *
   * @example
   * const tag = templateEngine.generateFieldTag('PRODUCTS/PRODUCT_NAME', {
   *   defaultValue: 'N/A',
   *   format: '@[^]*'
   * });
   * // Returns: <?PRODUCTS/PRODUCT_NAME:default='N/A':format='@[^]*'?>
   */
  generateFieldTag(fieldPath, options = {}) {
    try {
      let tag = `<?${fieldPath}`;

      if (options.defaultValue) {
        tag += `:default='${this._escapeString(options.defaultValue)}'`;
      }

      if (options.format) {
        tag += `:format='${this._escapeString(options.format)}'`;
      }

      if (options.multiline) {
        tag += ':multiline';
      }

      if (options.mask) {
        tag += `:mask='${this._escapeString(options.mask)}'`;
      }

      tag += '?>';

      this._storeTag('field', fieldPath, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate field tag: ${error.message}`);
    }
  }

  /**
   * Generate for-each loop opening tag
   * @param {string} groupPath - XPath for repeating group
   * @param {Object} [options={}] - Loop options
   * @param {number} [options.instancesPerPage] - Break pages after N instances
   * @returns {string} Opening tag
   *
   * @example
   * const tag = templateEngine.generateForEachTag('PRODUCTS/PRODUCT');
   * // Returns: <?for-each:PRODUCTS/PRODUCT?>
   */
  generateForEachTag(groupPath, options = {}) {
    try {
      let tag = `<?for-each:${groupPath}`;

      if (options.instancesPerPage) {
        tag += `:instancesPerPage='${options.instancesPerPage}'`;
      }

      tag += '?>';

      this._storeTag('forEach', groupPath, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate for-each tag: ${error.message}`);
    }
  }

  /**
   * Generate for-each closing tag
   * @returns {string} Closing tag
   *
   * @example
   * const tag = templateEngine.generateEndForEach();
   * // Returns: <?end for-each?>
   */
  generateEndForEach() {
    const tag = '<?end for-each?>';
    this._storeTag('endForEach', '', tag);
    return tag;
  }

  /**
   * Generate if (conditional) opening tag
   * @param {string} condition - XPath boolean condition
   * @returns {string} Opening tag
   *
   * @example
   * const tag = templateEngine.generateIfTag('STATUS = "ACTIVE"');
   * // Returns: <?if:STATUS = "ACTIVE"?>
   */
  generateIfTag(condition) {
    try {
      const tag = `<?if:${condition}?>`;
      this._storeTag('if', condition, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate if tag: ${error.message}`);
    }
  }

  /**
   * Generate if closing tag
   * @returns {string} Closing tag
   *
   * @example
   * const tag = templateEngine.generateEndIf();
   * // Returns: <?end if?>
   */
  generateEndIf() {
    const tag = '<?end if?>';
    this._storeTag('endIf', '', tag);
    return tag;
  }

  /**
   * Generate choose/when/otherwise conditional structure
   * @param {Array<Object>} conditions - Condition specifications
   * @param {string} conditions[].when - XPath when condition
   * @param {string} conditions[].then - Text/template for true case
   * @param {string} [conditions[].otherwise=''] - Else text
   * @returns {string} Complete choose structure
   *
   * @example
   * const tag = templateEngine.generateChooseTag([
   *   {when: 'PRICE > 1000', then: 'Premium'},
   *   {when: 'PRICE > 100', then: 'Standard'},
   *   {otherwise: 'Budget'}
   * ]);
   */
  generateChooseTag(conditions) {
    try {
      let tag = '<?choose?>\n';

      for (let cond of conditions) {
        if (cond.when) {
          tag += `<?when:${cond.when}?>\n`;
          tag += cond.then + '\n';
          tag += '<?end when?>\n';
        }
      }

      if (conditions[conditions.length - 1]?.otherwise) {
        tag += '<?otherwise?>\n';
        tag += conditions[conditions.length - 1].otherwise + '\n';
        tag += '<?end otherwise?>\n';
      }

      tag += '<?end choose?>';

      this._storeTag('choose', '', tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate choose tag: ${error.message}`);
    }
  }

  /**
   * Generate sort specification tag
   * @param {string} field - Field/XPath to sort by
   * @param {string} [order='asc'] - Sort order (asc or desc)
   * @returns {string} Sort tag
   *
   * @example
   * const tag = templateEngine.generateSortTag('PRODUCT_NAME', 'asc');
   * // Returns: <?sort:PRODUCT_NAME:asc?>
   */
  generateSortTag(field, order = 'asc') {
    try {
      const validOrder = ['asc', 'desc'].includes(order?.toLowerCase())
        ? order.toLowerCase()
        : 'asc';

      const tag = `<?sort:${field}:${validOrder}?>`;
      this._storeTag('sort', field, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate sort tag: ${error.message}`);
    }
  }

  /**
   * Generate grouping tag for data grouping
   * @param {string} field - Field to group by
   * @param {Object} [options={}] - Group options
   * @param {boolean} [options.pageBreak=false] - Break pages on group change
   * @returns {string} Group tag
   *
   * @example
   * const tag = templateEngine.generateGroupTag('REGION', {pageBreak: true});
   * // Returns: <?group:REGION:pageBreak='true'?>
   */
  generateGroupTag(field, options = {}) {
    try {
      let tag = `<?group:${field}`;

      if (options.pageBreak) {
        tag += `:pageBreak='true'`;
      }

      tag += '?>';

      this._storeTag('group', field, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate group tag: ${error.message}`);
    }
  }

  /**
   * Generate running total formula
   * @param {string} field - Field to accumulate
   * @param {string} [resetGroup=''] - Group to reset sum on
   * @returns {string} Running total tag
   *
   * @example
   * const tag = templateEngine.generateRunningTotal('SALES_AMOUNT', 'REGION');
   * // Returns: <?sum:SALES_AMOUNT:reset='REGION'?>
   */
  generateRunningTotal(field, resetGroup = '') {
    try {
      let tag = `<?sum:${field}`;

      if (resetGroup) {
        tag += `:reset='${resetGroup}'`;
      }

      tag += '?>';

      this._storeTag('runningTotal', field, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate running total tag: ${error.message}`);
    }
  }

  /**
   * Generate subtotal aggregation function
   * @param {string} field - Field to aggregate
   * @param {string} [func='SUM'] - Aggregation function
   * @returns {string} Subtotal tag
   *
   * @example
   * const tag = templateEngine.generateSubtotal('AMOUNT', 'AVG');
   * // Returns: <?subtotal:AMOUNT:AVG?>
   */
  generateSubtotal(field, func = 'SUM') {
    try {
      const validFunc = this.validFunctions.includes(func?.toUpperCase())
        ? func.toUpperCase()
        : 'SUM';

      const tag = `<?subtotal:${field}:${validFunc}?>`;
      this._storeTag('subtotal', field, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate subtotal tag: ${error.message}`);
    }
  }

  /**
   * Generate cross-tabulation (pivot) XSL tags
   * @param {Object} config - Cross-tab configuration
   * @param {string} config.rowField - Row dimension XPath
   * @param {string} config.colField - Column dimension XPath
   * @param {string} config.dataField - Data value XPath
   * @param {string} [config.aggregation='SUM'] - Aggregation function
   * @returns {string} Complete cross-tab tag structure
   *
   * @example
   * const tag = templateEngine.generateCrossTabTags({
   *   rowField: 'REGIONS/REGION',
   *   colField: 'MONTHS/MONTH',
   *   dataField: 'SALES/AMOUNT',
   *   aggregation: 'SUM'
   * });
   */
  generateCrossTabTags(config) {
    try {
      const {
        rowField,
        colField,
        dataField,
        aggregation = 'SUM'
      } = config;

      let tag = `<?crosstab:rows='${rowField}' cols='${colField}' data='${dataField}' agg='${aggregation}'?>\n`;
      tag += `<?for-each:${rowField}?>\n`;
      tag += `  <?for-each:${colField}?>\n`;
      tag += `    ${this.generateSubtotal(dataField, aggregation)}\n`;
      tag += `  <?end for-each?>\n`;
      tag += `<?end for-each?>\n`;
      tag += `<?end crosstab?>`;

      this._storeTag('crossTab', `${rowField}:${colField}`, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate cross-tab tags: ${error.message}`);
    }
  }

  /**
   * Generate chart definition XML
   * @param {Object} chartConfig - Chart configuration
   * @param {string} chartConfig.type - Chart type (column, bar, line, pie, etc.)
   * @param {string} chartConfig.title - Chart title
   * @param {string} chartConfig.categories - X-axis/category XPath
   * @param {string} chartConfig.values - Y-axis/value XPath
   * @param {string} [chartConfig.series] - Series/legend XPath
   * @param {Object} [chartConfig.options] - Additional chart options
   * @returns {string} Chart XML definition
   *
   * @example
   * const xml = templateEngine.generateChartXML({
   *   type: 'column',
   *   title: 'Sales by Region',
   *   categories: 'REGIONS/REGION_NAME',
   *   values: 'REGIONS/SALES_AMOUNT',
   *   series: 'YEARS/YEAR'
   * });
   */
  generateChartXML(chartConfig) {
    try {
      const {
        type = 'column',
        title = 'Chart',
        categories = '',
        values = '',
        series = '',
        options = {}
      } = chartConfig;

      if (!this.validChartTypes.includes(type)) {
        throw new Error(`Invalid chart type: ${type}`);
      }

      let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
      xml += `<chart>\n`;
      xml += `  <chartType>${type}</chartType>\n`;
      xml += `  <chartTitle>${this._escapeXml(title)}</chartTitle>\n`;

      if (categories) {
        xml += `  <categoryAxis>\n`;
        xml += `    <dataPath>${this._escapeXml(categories)}</dataPath>\n`;
        xml += `  </categoryAxis>\n`;
      }

      if (values) {
        xml += `  <valueAxis>\n`;
        xml += `    <dataPath>${this._escapeXml(values)}</dataPath>\n`;
        xml += `  </valueAxis>\n`;
      }

      if (series) {
        xml += `  <seriesAxis>\n`;
        xml += `    <dataPath>${this._escapeXml(series)}</dataPath>\n`;
        xml += `  </seriesAxis>\n`;
      }

      // Add options
      if (Object.keys(options).length > 0) {
        xml += `  <options>\n`;
        for (let [key, value] of Object.entries(options)) {
          xml += `    <${key}>${this._escapeXml(String(value))}</${key}>\n`;
        }
        xml += `  </options>\n`;
      }

      xml += `</chart>`;

      this._storeTag('chart', title, xml);
      return xml;
    } catch (error) {
      throw new Error(`Failed to generate chart XML: ${error.message}`);
    }
  }

  /**
   * Generate barcode tag
   * @param {string} field - Data field/XPath for barcode content
   * @param {string} type - Barcode type (code39, code128, ean13, upca, qrcode)
   * @param {Object} [options={}] - Barcode options
   * @param {boolean} [options.displayValue=false] - Show text under barcode
   * @returns {string} Barcode tag
   *
   * @example
   * const tag = templateEngine.generateBarcodeTag('PRODUCT_ID', 'code128', {
   *   displayValue: true
   * });
   */
  generateBarcodeTag(field, type, options = {}) {
    try {
      const validTypes = [
        'code39', 'code128', 'ean13', 'upca', 'ean8',
        'qrcode', 'pdf417', 'datamatrix'
      ];

      if (!validTypes.includes(type)) {
        throw new Error(`Invalid barcode type: ${type}`);
      }

      let tag = `<?barcode:data='${field}' type='${type}'`;

      if (options.displayValue) {
        tag += `:displayValue='true'`;
      }

      tag += '?>';

      this._storeTag('barcode', field, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate barcode tag: ${error.message}`);
    }
  }

  /**
   * Generate dynamic column tags
   * @param {Object} config - Dynamic column configuration
   * @param {Array<string>} config.columns - Column XPaths
   * @param {string} [config.rowField] - Row grouping field
   * @returns {string} Dynamic column tags
   *
   * @example
   * const tag = templateEngine.generateDynamicColumnTags({
   *   columns: ['MONTH_1', 'MONTH_2', 'MONTH_3'],
   *   rowField: 'REGION'
   * });
   */
  generateDynamicColumnTags(config) {
    try {
      const { columns = [], rowField = '' } = config;

      let tag = '<?dynamicColumns>\n';

      if (rowField) {
        tag += `<?for-each:${rowField}?>\n`;
      }

      for (let col of columns) {
        tag += `  <?column:${col}?>\n`;
      }

      if (rowField) {
        tag += `<?end for-each?>\n`;
      }

      tag += '<?end dynamicColumns?>';

      this._storeTag('dynamicColumn', columns.join(','), tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate dynamic column tags: ${error.message}`);
    }
  }

  /**
   * Generate conditional page break tag
   * @param {string} condition - XPath condition for page break
   * @returns {string} Page break tag
   *
   * @example
   * const tag = templateEngine.generatePageBreakTag('REGION != PREVIOUS_REGION');
   * // Returns: <?pagebreak:condition='REGION != PREVIOUS_REGION'?>
   */
  generatePageBreakTag(condition) {
    try {
      const tag = `<?pagebreak:condition='${condition}'?>`;
      this._storeTag('pageBreak', condition, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate page break tag: ${error.message}`);
    }
  }

  /**
   * Generate variable definition tag
   * @param {string} name - Variable name
   * @param {string} value - Variable value or expression
   * @returns {string} Variable tag
   *
   * @example
   * const tag = templateEngine.generateVariableTag('total_sales', 'SUM(SALES/AMOUNT)');
   * // Returns: <?variable:name='total_sales' value='SUM(SALES/AMOUNT)'?>
   */
  generateVariableTag(name, value) {
    try {
      const tag = `<?variable:name='${name}' value='${this._escapeString(value)}'?>`;
      this._storeTag('variable', name, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate variable tag: ${error.message}`);
    }
  }

  /**
   * Generate parameter tag for report parameters
   * @param {string} name - Parameter name
   * @param {string} [dataType='string'] - Parameter data type
   * @param {string} [defaultValue=''] - Default value
   * @returns {string} Parameter tag
   *
   * @example
   * const tag = templateEngine.generateParameterTag('p_year', 'integer', '2023');
   */
  generateParameterTag(name, dataType = 'string', defaultValue = '') {
    try {
      let tag = `<?parameter:name='${name}' dataType='${dataType}'`;

      if (defaultValue) {
        tag += ` default='${this._escapeString(defaultValue)}'`;
      }

      tag += '?>';

      this._storeTag('parameter', name, tag);
      return tag;
    } catch (error) {
      throw new Error(`Failed to generate parameter tag: ${error.message}`);
    }
  }

  /**
   * Parse and extract existing tags from document content
   * @param {string} documentContent - Full document text
   * @returns {Array<Object>} Extracted tags with positions
   *
   * @example
   * const tags = templateEngine.parseExistingTags(documentText);
   */
  parseExistingTags(documentContent) {
    try {
      if (!documentContent || typeof documentContent !== 'string') {
        return [];
      }

      const tags = [];
      const tagRegex = /<\?([^?]+)\?>/g;
      let match;

      while ((match = tagRegex.exec(documentContent)) !== null) {
        const tagContent = match[1];
        const tagType = tagContent.split(':')[0];

        tags.push({
          type: tagType,
          content: tagContent,
          fullTag: match[0],
          position: match.index,
          endPosition: match.index + match[0].length
        });
      }

      return tags;
    } catch (error) {
      throw new Error(`Tag parsing failed: ${error.message}`);
    }
  }

  /**
   * Validate template syntax and tag correctness
   * @param {string} documentContent - Full document text
   * @returns {Object} Validation result with {isValid: boolean, errors: Array}
   *
   * @example
   * const result = templateEngine.validateTemplate(documentText);
   */
  validateTemplate(documentContent) {
    try {
      const errors = [];
      const tags = this.parseExistingTags(documentContent);

      // Check for mismatched for-each/end for-each
      let forEachCount = 0;
      let ifCount = 0;
      let chooseCount = 0;

      for (let tag of tags) {
        switch (tag.type) {
          case 'for-each':
            forEachCount++;
            break;
          case 'end':
            if (tag.content.includes('for-each')) forEachCount--;
            else if (tag.content.includes('if')) ifCount--;
            else if (tag.content.includes('choose')) chooseCount--;
            break;
          case 'if':
            ifCount++;
            break;
          case 'choose':
            chooseCount++;
            break;
          default:
            break;
        }
      }

      if (forEachCount !== 0) {
        errors.push('Unmatched for-each tags: missing end for-each');
      }
      if (ifCount !== 0) {
        errors.push('Unmatched if tags: missing end if');
      }
      if (chooseCount !== 0) {
        errors.push('Unmatched choose tags: missing end choose');
      }

      // Check for valid tag syntax
      for (let tag of tags) {
        if (!this._isValidTag(tag.type, tag.content)) {
          errors.push(`Invalid tag syntax: ${tag.fullTag}`);
        }
      }

      return {
        isValid: errors.length === 0,
        errors,
        tagCount: tags.length
      };
    } catch (error) {
      return {
        isValid: false,
        errors: [error.message],
        tagCount: 0
      };
    }
  }

  /**
   * Get human-readable description of a tag
   * @param {string} tagName - Tag name (without brackets)
   * @returns {string} Description
   *
   * @example
   * const desc = templateEngine.getTagDescription('for-each');
   */
  getTagDescription(tagName) {
    const descriptions = {
      'for-each': 'Repeating group - loops through collection elements',
      'if': 'Conditional block - executes content if condition is true',
      'choose': 'Multiple conditions - like switch/case statement',
      'sort': 'Sorts data by specified field',
      'group': 'Groups data by specified field',
      'sum': 'Running total of numeric values',
      'subtotal': 'Aggregation function (sum, avg, count, etc.)',
      'crosstab': 'Pivot table - cross-tabulation',
      'chart': 'Chart/graph insertion',
      'barcode': 'Barcode image generation',
      'pagebreak': 'Conditional page break',
      'variable': 'Variable definition and assignment',
      'parameter': 'Report parameter definition'
    };

    return descriptions[tagName] || 'Unknown tag';
  }

  /**
   * Get list of supported aggregation functions
   * @returns {Array<string>} Function names
   */
  getSupportedFunctions() {
    return [...this.validFunctions];
  }

  /**
   * Get list of supported chart types
   * @returns {Array<string>} Chart type names
   */
  getSupportedChartTypes() {
    return [...this.validChartTypes];
  }

  // ============= Private Helper Methods =============

  /**
   * Store tag reference
   * @private
   * @param {string} type
   * @param {string} name
   * @param {string} tag
   * @returns {void}
   */
  _storeTag(type, name, tag) {
    const key = `${type}:${name}`;
    this.tags.set(key, {
      type,
      name,
      tag,
      createdAt: new Date()
    });
  }

  /**
   * Escape string for use in tags
   * @private
   * @param {string} str
   * @returns {string}
   */
  _escapeString(str) {
    if (!str) return '';
    return str
      .replace(/'/g, "''")
      .replace(/"/g, '\\"')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r');
  }

  /**
   * Escape string for XML
   * @private
   * @param {string} str
   * @returns {string}
   */
  _escapeXml(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Validate tag syntax
   * @private
   * @param {string} type
   * @param {string} content
   * @returns {boolean}
   */
  _isValidTag(type, content) {
    // Basic validation - can be extended
    const validTagTypes = [
      'for-each', 'end', 'if', 'else', 'otherwise',
      'choose', 'when', 'sort', 'group', 'sum', 'subtotal',
      'crosstab', 'chart', 'barcode', 'pagebreak',
      'variable', 'parameter', 'format', 'watermark'
    ];

    const baseType = type.split(':')[0];
    return validTagTypes.includes(baseType);
  }

  /**
   * Build tag reference documentation
   * @private
   * @returns {Object}
   */
  _buildTagReference() {
    return {
      field: {
        syntax: `<?fieldPath:options?>`,
        example: `<?PRODUCT_NAME:default='N/A'?>`,
        description: 'Insert data field value'
      },
      'for-each': {
        syntax: `<?for-each:groupPath?> ... <?end for-each?>`,
        example: `<?for-each:PRODUCTS/PRODUCT?> ... <?end for-each?>`,
        description: 'Repeat content for each group element'
      },
      if: {
        syntax: `<?if:condition?> ... <?end if?>`,
        example: `<?if:STATUS = 'ACTIVE'?> ... <?end if?>`,
        description: 'Conditional content'
      },
      sort: {
        syntax: `<?sort:field:order?>`,
        example: `<?sort:PRODUCT_NAME:asc?>`,
        description: 'Sort data by field'
      }
    };
  }
}

// Export
export default TemplateEngine;
export { TemplateEngine };
