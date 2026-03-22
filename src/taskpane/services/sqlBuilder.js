/**
 * SQL Query Builder
 * Constructs and validates SQL queries for Oracle databases
 * Supports SELECT, JOIN, WHERE, GROUP BY, ORDER BY, HAVING clauses
 * Generates proper Oracle SQL syntax with parameter binding
 *
 * @module sqlBuilder
 */

/**
 * SQLBuilder class constructs and manages SQL queries
 * Primarily supports Oracle SQL dialect with safe query building
 */
class SQLBuilder {
  /**
   * Creates a new SQL query builder instance
   */
  constructor() {
    this.query = {
      select: [],
      from: null,
      joins: [],
      where: [],
      groupBy: [],
      having: [],
      orderBy: [],
      limit: null,
      offset: null
    };
    this.parameters = [];
    this.parameterIndex = 0;
    this.dialect = 'ORACLE';
    this.queryHistory = [];
    this.maxQueryLength = 32000; // Oracle limit
  }

  /**
   * Build SELECT query with specified columns and conditions
   * @param {string} table - Table name
   * @param {Array<string>|string} columns - Column names to select
   * @param {Array<string>|string} [conditions=''] - WHERE conditions
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If parameters are invalid
   *
   * @example
   * builder.buildSelect('PRODUCTS', ['ID', 'NAME', 'PRICE'], 'STATUS = "ACTIVE"')
   *   .buildQuery();
   */
  buildSelect(table, columns, conditions = '') {
    try {
      // Validate inputs
      if (!table || typeof table !== 'string') {
        throw new Error('Table name is required and must be a string');
      }

      this.query.from = this._validateTableName(table);

      // Parse columns
      if (typeof columns === 'string') {
        this.query.select = this._parseColumns(columns);
      } else if (Array.isArray(columns)) {
        this.query.select = columns.map(col => this._validateColumnName(col));
      } else if (columns === '*') {
        this.query.select = ['*'];
      } else {
        throw new Error('Columns must be array, string, or "*"');
      }

      // Add conditions if provided
      if (conditions) {
        this.addWhere(conditions);
      }

      return this;
    } catch (error) {
      throw new Error(`buildSelect failed: ${error.message}`);
    }
  }

  /**
   * Add JOIN clause to query
   * @param {string} type - JOIN type (INNER, LEFT, RIGHT, FULL, CROSS)
   * @param {string} table - Table to join
   * @param {string} condition - JOIN condition (ON clause)
   * @param {string} [alias=''] - Table alias
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If JOIN parameters invalid
   *
   * @example
   * builder.addJoin('INNER', 'ORDERS', 'PRODUCTS.ID = ORDERS.PRODUCT_ID', 'O')
   */
  addJoin(type, table, condition, alias = '') {
    try {
      const validTypes = ['INNER', 'LEFT', 'RIGHT', 'FULL', 'CROSS'];
      const joinType = type?.toUpperCase();

      if (!validTypes.includes(joinType)) {
        throw new Error(`Invalid JOIN type: ${type}`);
      }

      const join = {
        type: joinType,
        table: this._validateTableName(table),
        condition: this._validateCondition(condition),
        alias: alias ? this._validateAlias(alias) : null
      };

      this.query.joins.push(join);
      return this;
    } catch (error) {
      throw new Error(`addJoin failed: ${error.message}`);
    }
  }

  /**
   * Add WHERE clause conditions
   * @param {string|Array<string>} condition - WHERE condition(s)
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If condition is invalid
   *
   * @example
   * builder.addWhere('STATUS = "ACTIVE"')
   *   .addWhere('PRICE > 100')
   */
  addWhere(condition) {
    try {
      if (!condition) {
        return this;
      }

      if (typeof condition === 'string') {
        const validatedCondition = this._validateCondition(condition);
        this.query.where.push(validatedCondition);
      } else if (Array.isArray(condition)) {
        condition.forEach(cond => {
          this.query.where.push(this._validateCondition(cond));
        });
      } else {
        throw new Error('WHERE condition must be string or array');
      }

      return this;
    } catch (error) {
      throw new Error(`addWhere failed: ${error.message}`);
    }
  }

  /**
   * Add GROUP BY clause
   * @param {string|Array<string>} columns - Column(s) to group by
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If columns invalid
   *
   * @example
   * builder.addGroupBy(['REGION', 'YEAR'])
   */
  addGroupBy(columns) {
    try {
      if (!columns) {
        return this;
      }

      if (typeof columns === 'string') {
        this.query.groupBy.push(this._validateColumnName(columns));
      } else if (Array.isArray(columns)) {
        columns.forEach(col => {
          this.query.groupBy.push(this._validateColumnName(col));
        });
      } else {
        throw new Error('GROUP BY must be string or array');
      }

      return this;
    } catch (error) {
      throw new Error(`addGroupBy failed: ${error.message}`);
    }
  }

  /**
   * Add ORDER BY clause with direction
   * @param {Array<Object>|string} columns - Column(s) and direction
   * @param {Array<string>} [directions=['ASC']] - Sort directions (ASC/DESC)
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If parameters invalid
   *
   * @example
   * builder.addOrderBy(['NAME', 'DATE'], ['ASC', 'DESC'])
   * // or
   * builder.addOrderBy('CREATED_DATE', 'DESC')
   */
  addOrderBy(columns, directions = []) {
    try {
      if (!columns) {
        return this;
      }

      // Normalize inputs
      const cols = typeof columns === 'string'
        ? [columns]
        : Array.isArray(columns) ? columns : [columns];

      const dirs = directions.length > 0
        ? directions.map(d => d?.toUpperCase() === 'DESC' ? 'DESC' : 'ASC')
        : Array(cols.length).fill('ASC');

      // Add order by clauses
      cols.forEach((col, idx) => {
        const validCol = this._validateColumnName(col);
        const direction = dirs[idx] || 'ASC';

        this.query.orderBy.push(`${validCol} ${direction}`);
      });

      return this;
    } catch (error) {
      throw new Error(`addOrderBy failed: ${error.message}`);
    }
  }

  /**
   * Add HAVING clause for aggregate conditions
   * @param {string} condition - HAVING condition
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If condition invalid
   *
   * @example
   * builder.addHaving('SUM(AMOUNT) > 1000')
   */
  addHaving(condition) {
    try {
      if (!condition) {
        return this;
      }

      const validCondition = this._validateCondition(condition);
      this.query.having.push(validCondition);
      return this;
    } catch (error) {
      throw new Error(`addHaving failed: ${error.message}`);
    }
  }

  /**
   * Set LIMIT and OFFSET for pagination
   * @param {number} limit - Maximum rows to return
   * @param {number} [offset=0] - Starting row offset
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If values invalid
   *
   * @example
   * builder.setLimit(50, 100) // LIMIT 50 OFFSET 100
   */
  setLimit(limit, offset = 0) {
    try {
      if (!Number.isInteger(limit) || limit <= 0) {
        throw new Error('Limit must be positive integer');
      }

      if (!Number.isInteger(offset) || offset < 0) {
        throw new Error('Offset must be non-negative integer');
      }

      this.query.limit = limit;
      this.query.offset = offset;
      return this;
    } catch (error) {
      throw new Error(`setLimit failed: ${error.message}`);
    }
  }

  /**
   * Build and return complete SQL query string
   * @returns {string} Complete SQL query
   * @throws {Error} If query construction fails
   *
   * @example
   * const sql = builder
   *   .buildSelect('PRODUCTS', ['ID', 'NAME', 'PRICE'])
   *   .addWhere('STATUS = "ACTIVE"')
   *   .addOrderBy('NAME', 'ASC')
   *   .buildQuery();
   */
  buildQuery() {
    try {
      if (!this.query.from) {
        throw new Error('No table specified for SELECT');
      }

      if (this.query.select.length === 0) {
        throw new Error('No columns specified for SELECT');
      }

      let sql = this._buildSelectClause();
      sql += this._buildFromClause();
      sql += this._buildJoinClauses();
      sql += this._buildWhereClause();
      sql += this._buildGroupByClause();
      sql += this._buildHavingClause();
      sql += this._buildOrderByClause();
      sql += this._buildLimitClause();

      // Validate query length
      if (sql.length > this.maxQueryLength) {
        throw new Error(
          `Query exceeds maximum length of ${this.maxQueryLength} characters`
        );
      }

      // Store in history
      this.queryHistory.push({
        query: sql,
        timestamp: new Date(),
        parameters: [...this.parameters]
      });

      return sql;
    } catch (error) {
      throw new Error(`buildQuery failed: ${error.message}`);
    }
  }

  /**
   * Parse existing SQL query into builder components
   * @param {string} sql - SQL query string to parse
   * @returns {SQLBuilder} Returns this for chaining
   * @throws {Error} If SQL cannot be parsed
   *
   * @example
   * builder.parseQuery('SELECT ID, NAME FROM PRODUCTS WHERE STATUS = "ACTIVE"')
   */
  parseQuery(sql) {
    try {
      if (!sql || typeof sql !== 'string') {
        throw new Error('SQL must be non-empty string');
      }

      // Simple regex-based parsing
      const upperSql = sql.toUpperCase();

      // Extract SELECT
      const selectMatch = sql.match(/SELECT\s+(.*?)\s+FROM/i);
      if (selectMatch) {
        const columns = selectMatch[1].split(',').map(col => col.trim());
        this.query.select = columns;
      }

      // Extract FROM
      const fromMatch = sql.match(/FROM\s+(\w+)/i);
      if (fromMatch) {
        this.query.from = fromMatch[1];
      }

      // Extract WHERE
      const whereMatch = sql.match(/WHERE\s+(.*?)(?:GROUP BY|ORDER BY|LIMIT|$)/i);
      if (whereMatch) {
        this.query.where = [whereMatch[1].trim()];
      }

      // Extract GROUP BY
      const groupMatch = sql.match(/GROUP BY\s+(.*?)(?:HAVING|ORDER BY|LIMIT|$)/i);
      if (groupMatch) {
        this.query.groupBy = groupMatch[1].split(',').map(col => col.trim());
      }

      // Extract HAVING
      const havingMatch = sql.match(/HAVING\s+(.*?)(?:ORDER BY|LIMIT|$)/i);
      if (havingMatch) {
        this.query.having = [havingMatch[1].trim()];
      }

      // Extract ORDER BY
      const orderMatch = sql.match(/ORDER BY\s+(.*?)(?:LIMIT|$)/i);
      if (orderMatch) {
        this.query.orderBy = orderMatch[1].split(',').map(col => col.trim());
      }

      // Extract LIMIT
      const limitMatch = sql.match(/LIMIT\s+(\d+)(?:\s+OFFSET\s+(\d+))?/i);
      if (limitMatch) {
        this.query.limit = parseInt(limitMatch[1]);
        this.query.offset = limitMatch[2] ? parseInt(limitMatch[2]) : 0;
      }

      return this;
    } catch (error) {
      throw new Error(`parseQuery failed: ${error.message}`);
    }
  }

  /**
   * Validate SQL query for syntax errors
   * @param {string} sql - SQL query to validate
   * @returns {Object} Validation result {isValid: boolean, errors: Array}
   *
   * @example
   * const result = sqlBuilder.validateQuery(sql);
   */
  validateQuery(sql) {
    try {
      const errors = [];

      if (!sql || typeof sql !== 'string') {
        return { isValid: false, errors: ['SQL must be non-empty string'] };
      }

      // Check for required keywords
      if (!sql.match(/SELECT/i)) {
        errors.push('Missing SELECT keyword');
      }

      if (!sql.match(/FROM/i)) {
        errors.push('Missing FROM clause');
      }

      // Check for matching parentheses
      const openParens = (sql.match(/\(/g) || []).length;
      const closeParens = (sql.match(/\)/g) || []).length;

      if (openParens !== closeParens) {
        errors.push('Unmatched parentheses');
      }

      // Check for suspicious patterns
      if (sql.match(/;\s*(?:DROP|DELETE|TRUNCATE|UPDATE)/i)) {
        errors.push('Potentially dangerous SQL detected (multiple statements)');
      }

      // Check query length
      if (sql.length > this.maxQueryLength) {
        errors.push(
          `Query exceeds maximum length of ${this.maxQueryLength} characters`
        );
      }

      return {
        isValid: errors.length === 0,
        errors,
        length: sql.length
      };
    } catch (error) {
      return {
        isValid: false,
        errors: [error.message]
      };
    }
  }

  /**
   * Format/beautify SQL query with proper indentation
   * @param {string} sql - SQL query to format
   * @returns {string} Formatted SQL
   *
   * @example
   * const formatted = sqlBuilder.formatQuery(sql);
   */
  formatQuery(sql) {
    try {
      if (!sql || typeof sql !== 'string') {
        return '';
      }

      let formatted = sql;

      // Add newlines before major keywords
      const keywords = [
        'SELECT', 'FROM', 'WHERE', 'GROUP BY', 'HAVING',
        'ORDER BY', 'LIMIT', 'JOIN', 'LEFT JOIN', 'INNER JOIN'
      ];

      keywords.forEach(keyword => {
        const regex = new RegExp(`\\s+${keyword}\\s+`, 'gi');
        formatted = formatted.replace(regex, `\n${keyword} `);
      });

      // Indent continuation lines
      const lines = formatted.split('\n');
      formatted = lines
        .map((line, idx) => {
          if (idx === 0) return line;
          return '  ' + line.trim();
        })
        .join('\n');

      return formatted;
    } catch (error) {
      console.warn('Format query failed:', error.message);
      return sql;
    }
  }

  /**
   * Extract metadata from SQL query
   * @param {string} sql - SQL query
   * @returns {Object} Metadata with tables, columns, conditions
   *
   * @example
   * const meta = sqlBuilder.getQueryMetadata(sql);
   */
  getQueryMetadata(sql) {
    try {
      const metadata = {
        tables: [],
        columns: [],
        whereConditions: [],
        groupByColumns: [],
        orderByColumns: [],
        joins: []
      };

      if (!sql || typeof sql !== 'string') {
        return metadata;
      }

      // Extract tables
      const fromMatch = sql.match(/FROM\s+(\w+)/gi);
      if (fromMatch) {
        metadata.tables.push(...fromMatch.map(m => m.replace(/FROM\s+/i, '')));
      }

      // Extract columns from SELECT
      const selectMatch = sql.match(/SELECT\s+(.*?)\s+FROM/i);
      if (selectMatch) {
        const cols = selectMatch[1].split(',').map(col => col.trim());
        metadata.columns = cols;
      }

      // Extract WHERE conditions
      const whereMatch = sql.match(/WHERE\s+(.*?)(?:GROUP BY|ORDER BY|LIMIT|$)/i);
      if (whereMatch) {
        metadata.whereConditions = whereMatch[1]
          .split(/AND|OR/i)
          .map(cond => cond.trim());
      }

      // Extract GROUP BY
      const groupMatch = sql.match(/GROUP BY\s+(.*?)(?:HAVING|ORDER BY|$)/i);
      if (groupMatch) {
        metadata.groupByColumns = groupMatch[1]
          .split(',')
          .map(col => col.trim());
      }

      // Extract ORDER BY
      const orderMatch = sql.match(/ORDER BY\s+(.*?)(?:LIMIT|$)/i);
      if (orderMatch) {
        metadata.orderByColumns = orderMatch[1]
          .split(',')
          .map(col => col.trim());
      }

      // Extract JOINs
      const joinMatches = sql.match(/(?:LEFT|RIGHT|INNER|FULL)?\s*JOIN\s+(\w+)/gi);
      if (joinMatches) {
        metadata.joins = joinMatches.map(m => m.replace(/JOIN\s+/i, ''));
      }

      return metadata;
    } catch (error) {
      console.warn('Get metadata failed:', error.message);
      return null;
    }
  }

  /**
   * Get list of supported SQL functions
   * @returns {Array<string>} Available functions
   *
   * @example
   * const functions = sqlBuilder.getSupportedFunctions();
   */
  getSupportedFunctions() {
    return [
      // Aggregate Functions
      'SUM', 'AVG', 'MIN', 'MAX', 'COUNT', 'STDDEV', 'VARIANCE',
      'MEDIAN', 'MODE',
      // String Functions
      'SUBSTR', 'INSTR', 'LENGTH', 'UPPER', 'LOWER', 'INITCAP',
      'TRIM', 'LTRIM', 'RTRIM', 'CONCAT', 'REPLACE', 'TRANSLATE',
      // Numeric Functions
      'ABS', 'CEIL', 'FLOOR', 'ROUND', 'TRUNC', 'MOD', 'POWER',
      'SQRT', 'SIGN', 'GREATEST', 'LEAST',
      // Date Functions
      'SYSDATE', 'TRUNC', 'ROUND', 'ADD_MONTHS', 'MONTHS_BETWEEN',
      'NEXT_DAY', 'LAST_DAY', 'EXTRACT',
      // Conversion Functions
      'TO_DATE', 'TO_CHAR', 'TO_NUMBER', 'CAST',
      // Other
      'DECODE', 'CASE', 'COALESCE', 'NVL', 'NVL2'
    ];
  }

  /**
   * Reset builder to initial state
   * @returns {void}
   *
   * @example
   * builder.reset();
   */
  reset() {
    this.query = {
      select: [],
      from: null,
      joins: [],
      where: [],
      groupBy: [],
      having: [],
      orderBy: [],
      limit: null,
      offset: null
    };
    this.parameters = [];
    this.parameterIndex = 0;
  }

  /**
   * Get query history
   * @returns {Array} Previous queries
   */
  getHistory() {
    return [...this.queryHistory];
  }

  /**
   * Clear query history
   * @returns {void}
   */
  clearHistory() {
    this.queryHistory = [];
  }

  // ============= Private Helper Methods =============

  /**
   * Validate table name
   * @private
   * @param {string} tableName
   * @returns {string} Validated name
   */
  _validateTableName(tableName) {
    const trimmed = tableName.trim();

    if (!/^[a-zA-Z_][a-zA-Z0-9_$#]*$/.test(trimmed)) {
      throw new Error(`Invalid table name: ${tableName}`);
    }

    return trimmed;
  }

  /**
   * Validate column name
   * @private
   * @param {string} columnName
   * @returns {string} Validated name
   */
  _validateColumnName(columnName) {
    const trimmed = columnName.trim();

    if (trimmed === '*') return '*';

    // Allow schema.table.column notation
    if (!/^[a-zA-Z_][a-zA-Z0-9_$.#]*$/.test(trimmed)) {
      throw new Error(`Invalid column name: ${columnName}`);
    }

    return trimmed;
  }

  /**
   * Validate alias
   * @private
   * @param {string} alias
   * @returns {string} Validated alias
   */
  _validateAlias(alias) {
    const trimmed = alias.trim();

    if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(trimmed)) {
      throw new Error(`Invalid alias: ${alias}`);
    }

    return trimmed;
  }

  /**
   * Validate condition
   * @private
   * @param {string} condition
   * @returns {string} Validated condition
   */
  _validateCondition(condition) {
    if (!condition || typeof condition !== 'string') {
      throw new Error('Condition must be non-empty string');
    }

    const trimmed = condition.trim();

    // Check for dangerous SQL injection patterns
    if (trimmed.match(/;\s*(?:DROP|DELETE|TRUNCATE|UPDATE|INSERT)/i)) {
      throw new Error('Potentially dangerous SQL detected in condition');
    }

    return trimmed;
  }

  /**
   * Parse column list
   * @private
   * @param {string} columns
   * @returns {Array}
   */
  _parseColumns(columns) {
    return columns.split(',').map(col => this._validateColumnName(col.trim()));
  }

  /**
   * Build SELECT clause
   * @private
   * @returns {string}
   */
  _buildSelectClause() {
    const columns = this.query.select.join(', ');
    return `SELECT ${columns}`;
  }

  /**
   * Build FROM clause
   * @private
   * @returns {string}
   */
  _buildFromClause() {
    return `\nFROM ${this.query.from}`;
  }

  /**
   * Build JOIN clauses
   * @private
   * @returns {string}
   */
  _buildJoinClauses() {
    if (this.query.joins.length === 0) return '';

    return this.query.joins
      .map(join => {
        let clause = `\n${join.type} JOIN ${join.table}`;
        if (join.alias) clause += ` ${join.alias}`;
        clause += ` ON ${join.condition}`;
        return clause;
      })
      .join('');
  }

  /**
   * Build WHERE clause
   * @private
   * @returns {string}
   */
  _buildWhereClause() {
    if (this.query.where.length === 0) return '';

    const conditions = this.query.where.join('\n  AND ');
    return `\nWHERE ${conditions}`;
  }

  /**
   * Build GROUP BY clause
   * @private
   * @returns {string}
   */
  _buildGroupByClause() {
    if (this.query.groupBy.length === 0) return '';

    const columns = this.query.groupBy.join(', ');
    return `\nGROUP BY ${columns}`;
  }

  /**
   * Build HAVING clause
   * @private
   * @returns {string}
   */
  _buildHavingClause() {
    if (this.query.having.length === 0) return '';

    const conditions = this.query.having.join('\n  AND ');
    return `\nHAVING ${conditions}`;
  }

  /**
   * Build ORDER BY clause
   * @private
   * @returns {string}
   */
  _buildOrderByClause() {
    if (this.query.orderBy.length === 0) return '';

    const columns = this.query.orderBy.join(', ');
    return `\nORDER BY ${columns}`;
  }

  /**
   * Build LIMIT clause (Oracle specific)
   * @private
   * @returns {string}
   */
  _buildLimitClause() {
    if (this.query.limit === null && this.query.offset === null) {
      return '';
    }

    // Oracle uses OFFSET/FETCH in 12c+
    let clause = '';

    if (this.query.offset !== null && this.query.offset > 0) {
      clause += `\nOFFSET ${this.query.offset} ROWS`;
    }

    if (this.query.limit !== null) {
      clause += `\nFETCH NEXT ${this.query.limit} ROWS ONLY`;
    }

    return clause;
  }
}

// Export
export default SQLBuilder;
export { SQLBuilder };
