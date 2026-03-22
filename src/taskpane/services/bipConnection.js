/**
 * BI Publisher Server Connection Manager
 * Handles authentication, communication, and data retrieval from Oracle BI Publisher
 * REST API v12c and later versions
 *
 * @module bipConnection
 * @requires axios or native fetch
 */

/**
 * BIPConnection class manages connections to BI Publisher servers
 * Supports online and offline modes with caching and session management
 */
class BIPConnection {
  /**
   * Creates a new BI Publisher connection instance
   */
  constructor() {
    this.serverUrl = null;
    this.username = null;
    this.authToken = null;
    this.sessionId = null;
    this.isConnected = false;
    this.isOfflineMode = false;
    this.tokenExpiresAt = null;
    this.cache = new Map();
    this.retryConfig = {
      maxRetries: 3,
      retryDelay: 1000,
      backoffMultiplier: 2
    };
    this.timeout = 30000; // 30 seconds
  }

  /**
   * Connect to BI Publisher server with credentials
   * @async
   * @param {string} serverUrl - Base URL of BI Publisher (e.g., http://bipserver:9502/xmlpserver)
   * @param {string} username - Username for authentication
   * @param {string} password - Password for authentication
   * @returns {Promise<Object>} Connection result with status and server info
   * @throws {Error} Connection or authentication errors
   *
   * @example
   * const bipConn = new BIPConnection();
   * await bipConn.connect('http://localhost:9502/xmlpserver', 'admin', 'admin123');
   */
  async connect(serverUrl, username, password) {
    try {
      // Validate inputs
      if (!serverUrl || !username || !password) {
        throw new Error('Server URL, username, and password are required');
      }

      this.serverUrl = this._normalizeServerUrl(serverUrl);
      this.username = username;

      // Attempt authentication
      const authResponse = await this._authenticate(username, password);

      this.authToken = authResponse.token;
      this.sessionId = authResponse.sessionId;
      this.isConnected = true;
      this.isOfflineMode = false;

      // Set token expiration (typically 30 minutes)
      this.tokenExpiresAt = new Date(Date.now() + 30 * 60 * 1000);

      // Clear cache on new connection
      this.cache.clear();

      return {
        success: true,
        message: 'Successfully connected to BI Publisher',
        sessionId: this.sessionId,
        serverUrl: this.serverUrl
      };
    } catch (error) {
      this.isConnected = false;
      throw new Error(`Connection failed: ${error.message}`);
    }
  }

  /**
   * Disconnect from BI Publisher server
   * @async
   * @returns {Promise<void>}
   */
  async disconnect() {
    if (this.isConnected && this.sessionId) {
      try {
        const endpoint = `/admin/v2/session`;
        await this._makeRequest('DELETE', endpoint);
      } catch (error) {
        console.warn('Disconnect failed (non-critical):', error.message);
      }
    }

    this.authToken = null;
    this.sessionId = null;
    this.isConnected = false;
    this.cache.clear();
  }

  /**
   * Test connection to BI Publisher server
   * @async
   * @returns {Promise<Object>} Test result with server status
   */
  async testConnection() {
    try {
      if (!this.isConnected) {
        throw new Error('Not connected to server');
      }

      // Refresh token if near expiration
      if (this._isTokenNearExpiration()) {
        await this._refreshToken();
      }

      const serverInfo = await this.getServerInfo();
      return {
        success: true,
        connected: true,
        serverInfo
      };
    } catch (error) {
      return {
        success: false,
        connected: false,
        error: error.message
      };
    }
  }

  /**
   * Get list of available reports on the server
   * @async
   * @param {string} [folder='/'] - Optional folder path to list reports from
   * @returns {Promise<Array>} Array of report objects with paths and properties
   * @throws {Error} If not connected or API call fails
   *
   * @example
   * const reports = await bipConn.getReports('/Reports/Financial');
   */
  async getReports(folder = '/') {
    try {
      this._checkConnection();

      const cacheKey = `reports:${folder}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      const endpoint = `/api/v2/jobs`;
      const params = new URLSearchParams({
        startIndex: 0,
        fetchSize: 500,
        jobType: 'REPORT',
        reportPath: folder
      });

      const response = await this._makeRequest('GET', `${endpoint}?${params}`);
      const reports = response.jobs || [];

      this.cache.set(cacheKey, reports);
      return reports;
    } catch (error) {
      throw new Error(`Failed to get reports: ${error.message}`);
    }
  }

  /**
   * Get report data in XML format
   * @async
   * @param {string} reportPath - Path to the report (e.g., /Reports/Sales)
   * @param {Object} [parameters={}] - Report parameters and values
   * @param {string} [format='xml'] - Output format (xml, pdf, html, xls)
   * @returns {Promise<string>} Report data as XML string
   * @throws {Error} If report execution fails
   *
   * @example
   * const data = await bipConn.getReportData('/Reports/Sales', {year: 2023, region: 'North'});
   */
  async getReportData(reportPath, parameters = {}, format = 'xml') {
    try {
      this._checkConnection();

      const cacheKey = `reportData:${reportPath}:${JSON.stringify(parameters)}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      // First, run the report
      const jobId = await this._runReportJob(reportPath, parameters, format);

      // Poll for completion
      const reportOutput = await this._pollJobCompletion(jobId, format);

      this.cache.set(cacheKey, reportOutput);
      return reportOutput;
    } catch (error) {
      throw new Error(`Failed to get report data: ${error.message}`);
    }
  }

  /**
   * Get data model definition for a report
   * @async
   * @param {string} reportPath - Path to the report
   * @returns {Promise<Object>} Data model structure with elements and attributes
   * @throws {Error} If unable to retrieve data model
   *
   * @example
   * const model = await bipConn.getDataModel('/Reports/Sales');
   * console.log(model.elements); // Array of data elements
   */
  async getDataModel(reportPath) {
    try {
      this._checkConnection();

      const cacheKey = `dataModel:${reportPath}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      // Get sample data first
      const sampleData = await this.getReportData(reportPath, {}, 'xml');

      // Parse XML to extract structure
      const dataModel = this._parseDataModel(sampleData);

      this.cache.set(cacheKey, dataModel);
      return dataModel;
    } catch (error) {
      throw new Error(`Failed to get data model: ${error.message}`);
    }
  }

  /**
   * Upload a template file to BI Publisher
   * @async
   * @param {string} reportPath - Path where template will be stored
   * @param {File|Blob} templateFile - Template file (RTF or DOCX)
   * @param {string} templateName - Name of the template
   * @returns {Promise<Object>} Upload result
   * @throws {Error} If upload fails
   *
   * @example
   * const result = await bipConn.uploadTemplate('/Reports/Sales', templateBlob, 'SalesReport');
   */
  async uploadTemplate(reportPath, templateFile, templateName) {
    try {
      this._checkConnection();

      const formData = new FormData();
      formData.append('reportPath', reportPath);
      formData.append('templateName', templateName);
      formData.append('templateContent', templateFile);

      const endpoint = `/api/v2/templates/upload`;
      const response = await this._makeRequest('POST', endpoint, formData, true);

      return {
        success: true,
        templateId: response.templateId,
        message: 'Template uploaded successfully'
      };
    } catch (error) {
      throw new Error(`Template upload failed: ${error.message}`);
    }
  }

  /**
   * Download template file from BI Publisher
   * @async
   * @param {string} reportPath - Path to the report
   * @param {string} templateName - Name of the template to download
   * @returns {Promise<Blob>} Template file as Blob
   * @throws {Error} If download fails
   *
   * @example
   * const templateBlob = await bipConn.downloadTemplate('/Reports/Sales', 'SalesReport');
   */
  async downloadTemplate(reportPath, templateName) {
    try {
      this._checkConnection();

      const endpoint = `/api/v2/templates/download`;
      const params = new URLSearchParams({
        reportPath,
        templateName
      });

      const response = await this._makeRequest(
        'GET',
        `${endpoint}?${params}`,
        null,
        false,
        true
      );

      return response;
    } catch (error) {
      throw new Error(`Template download failed: ${error.message}`);
    }
  }

  /**
   * Get report parameters and their configuration
   * @async
   * @param {string} reportPath - Path to the report
   * @returns {Promise<Array>} Array of parameter definitions
   * @throws {Error} If unable to retrieve parameters
   *
   * @example
   * const params = await bipConn.getParameters('/Reports/Sales');
   * // Returns: [{ name: 'Year', type: 'integer', required: true, lov: {...} }, ...]
   */
  async getParameters(reportPath) {
    try {
      this._checkConnection();

      const cacheKey = `parameters:${reportPath}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      const endpoint = `/api/v2/catalogs/parameters`;
      const params = new URLSearchParams({ reportPath });

      const response = await this._makeRequest('GET', `${endpoint}?${params}`);
      const parameters = response.parameters || [];

      this.cache.set(cacheKey, parameters);
      return parameters;
    } catch (error) {
      throw new Error(`Failed to get parameters: ${error.message}`);
    }
  }

  /**
   * Get list of values (LOV) for a report parameter
   * @async
   * @param {string} parameterName - Name of the parameter
   * @param {string} [reportPath] - Optional report path for context
   * @returns {Promise<Array>} Array of valid values for the parameter
   * @throws {Error} If LOV retrieval fails
   *
   * @example
   * const values = await bipConn.getLOV('Region');
   * // Returns: ['North', 'South', 'East', 'West']
   */
  async getLOV(parameterName, reportPath) {
    try {
      this._checkConnection();

      const cacheKey = `lov:${parameterName}:${reportPath || ''}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      const endpoint = `/api/v2/catalogs/lov`;
      const params = new URLSearchParams({ parameterName });
      if (reportPath) {
        params.append('reportPath', reportPath);
      }

      const response = await this._makeRequest('GET', `${endpoint}?${params}`);
      const values = response.values || [];

      this.cache.set(cacheKey, values);
      return values;
    } catch (error) {
      throw new Error(`Failed to get LOV: ${error.message}`);
    }
  }

  /**
   * Run a report with specified parameters and get output
   * @async
   * @param {string} reportPath - Path to the report
   * @param {Object} parameters - Report parameters
   * @param {string} format - Output format (pdf, html, xls, xml, json)
   * @returns {Promise<string|Blob>} Report output in specified format
   * @throws {Error} If report execution fails
   *
   * @example
   * const pdf = await bipConn.runReport('/Reports/Sales', {year: 2023}, 'pdf');
   */
  async runReport(reportPath, parameters, format) {
    try {
      this._checkConnection();

      const jobId = await this._runReportJob(reportPath, parameters, format);
      return await this._pollJobCompletion(jobId, format);
    } catch (error) {
      throw new Error(`Report execution failed: ${error.message}`);
    }
  }

  /**
   * Get server information and version
   * @async
   * @returns {Promise<Object>} Server info with version, build, and capabilities
   * @throws {Error} If unable to retrieve server info
   *
   * @example
   * const info = await bipConn.getServerInfo();
   * // Returns: { version: '12.2.1.2', build: 1234, supportedFormats: [...] }
   */
  async getServerInfo() {
    try {
      this._checkConnection();

      const cacheKey = 'serverInfo';
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }

      const endpoint = `/admin/v2/about`;
      const response = await this._makeRequest('GET', endpoint);

      const serverInfo = {
        version: response.version || 'Unknown',
        build: response.build || 'Unknown',
        supportedFormats: response.supportedFormats || ['pdf', 'html', 'xls', 'xml'],
        features: response.features || []
      };

      this.cache.set(cacheKey, serverInfo);
      return serverInfo;
    } catch (error) {
      throw new Error(`Failed to get server info: ${error.message}`);
    }
  }

  /**
   * Clear the connection cache
   * @returns {void}
   */
  clearCache() {
    this.cache.clear();
  }

  /**
   * Check if currently connected
   * @returns {boolean}
   */
  isConnectedToServer() {
    return this.isConnected && !this._isTokenNearExpiration();
  }

  /**
   * Get current session info
   * @returns {Object} Session information
   */
  getSessionInfo() {
    return {
      isConnected: this.isConnected,
      sessionId: this.sessionId,
      username: this.username,
      serverUrl: this.serverUrl,
      tokenExpiresAt: this.tokenExpiresAt,
      isOfflineMode: this.isOfflineMode
    };
  }

  // ============= Private Helper Methods =============

  /**
   * Authenticate with server and get session token
   * @private
   * @async
   * @param {string} username
   * @param {string} password
   * @returns {Promise<Object>} Token and session info
   */
  async _authenticate(username, password) {
    try {
      const endpoint = `/admin/v2/login`;
      const payload = { username, password };

      const response = await this._makeRequest('POST', endpoint, payload);

      return {
        token: response.sessionID || response.token,
        sessionId: response.sessionID || response.id
      };
    } catch (error) {
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }

  /**
   * Refresh authentication token
   * @private
   * @async
   * @returns {Promise<void>}
   */
  async _refreshToken() {
    try {
      const endpoint = `/admin/v2/session/refresh`;
      const response = await this._makeRequest('POST', endpoint);

      this.authToken = response.sessionID || response.token;
      this.tokenExpiresAt = new Date(Date.now() + 30 * 60 * 1000);
    } catch (error) {
      console.warn('Token refresh failed:', error.message);
    }
  }

  /**
   * Run a report job on the server
   * @private
   * @async
   * @param {string} reportPath
   * @param {Object} parameters
   * @param {string} format
   * @returns {Promise<string>} Job ID
   */
  async _runReportJob(reportPath, parameters, format) {
    const endpoint = `/api/v2/jobs`;
    const payload = {
      jobType: 'REPORT',
      reportPath,
      outputFormat: format,
      parameters
    };

    const response = await this._makeRequest('POST', endpoint, payload);
    return response.jobId || response.id;
  }

  /**
   * Poll for job completion and retrieve output
   * @private
   * @async
   * @param {string} jobId
   * @param {string} format
   * @returns {Promise<any>} Job output
   */
  async _pollJobCompletion(jobId, format) {
    const maxAttempts = 60;
    const pollInterval = 500;

    for (let attempt = 0; attempt < maxAttempts; attempt++) {
      const endpoint = `/api/v2/jobs/${jobId}`;
      const response = await this._makeRequest('GET', endpoint);

      if (response.status === 'COMPLETED') {
        return await this._getJobOutput(jobId, format);
      } else if (response.status === 'FAILED') {
        throw new Error(`Job failed: ${response.errorMessage}`);
      }

      await new Promise(resolve => setTimeout(resolve, pollInterval));
    }

    throw new Error('Job execution timeout');
  }

  /**
   * Get output from completed job
   * @private
   * @async
   * @param {string} jobId
   * @param {string} format
   * @returns {Promise<any>} Job output
   */
  async _getJobOutput(jobId, format) {
    const endpoint = `/api/v2/jobs/${jobId}/output`;
    return await this._makeRequest(
      'GET',
      endpoint,
      null,
      false,
      format !== 'xml' && format !== 'json'
    );
  }

  /**
   * Parse XML data to extract data model structure
   * @private
   * @param {string} xmlData
   * @returns {Object} Data model
   */
  _parseDataModel(xmlData) {
    try {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlData, 'text/xml');

      if (xmlDoc.parseError && xmlDoc.parseError.errorCode !== 0) {
        throw new Error('XML parse error');
      }

      const elements = [];
      const processNode = (node, path = '') => {
        if (node.nodeType === 1) { // Element node
          const nodePath = path ? `${path}/${node.nodeName}` : node.nodeName;
          elements.push({
            name: node.nodeName,
            path: nodePath,
            type: this._inferDataType(node.textContent),
            attributes: Array.from(node.attributes || []).map(attr => attr.name)
          });

          Array.from(node.childNodes).forEach(child => {
            processNode(child, nodePath);
          });
        }
      };

      processNode(xmlDoc.documentElement);

      return {
        rootElement: xmlDoc.documentElement.nodeName,
        elements,
        namespaces: this._extractNamespaces(xmlDoc)
      };
    } catch (error) {
      throw new Error(`Data model parsing failed: ${error.message}`);
    }
  }

  /**
   * Extract namespace declarations from XML
   * @private
   * @param {Document} xmlDoc
   * @returns {Object} Namespace map
   */
  _extractNamespaces(xmlDoc) {
    const namespaces = {};
    const root = xmlDoc.documentElement;

    for (let attr of root.attributes || []) {
      if (attr.name.startsWith('xmlns')) {
        const prefix = attr.name === 'xmlns' ? '' : attr.name.split(':')[1];
        namespaces[prefix || 'default'] = attr.value;
      }
    }

    return namespaces;
  }

  /**
   * Infer data type from value
   * @private
   * @param {string} value
   * @returns {string} Data type (string, number, date, boolean)
   */
  _inferDataType(value) {
    if (!value) return 'string';
    if (/^\d+$/.test(value)) return 'number';
    if (/^\d{4}-\d{2}-\d{2}/.test(value)) return 'date';
    if (/^(true|false)$/i.test(value)) return 'boolean';
    return 'string';
  }

  /**
   * Make HTTP request with retry logic and error handling
   * @private
   * @async
   * @param {string} method - HTTP method (GET, POST, etc.)
   * @param {string} endpoint - API endpoint
   * @param {any} [data=null] - Request payload
   * @param {boolean} [isFormData=false] - Is request body FormData
   * @param {boolean} [returnBlob=false] - Return response as Blob
   * @returns {Promise<any>} Response data
   */
  async _makeRequest(
    method,
    endpoint,
    data = null,
    isFormData = false,
    returnBlob = false
  ) {
    let lastError;

    for (let attempt = 0; attempt < this.retryConfig.maxRetries; attempt++) {
      try {
        const url = `${this.serverUrl}${endpoint}`;
        const options = {
          method,
          headers: this._getRequestHeaders(isFormData),
          timeout: this.timeout
        };

        if (this.authToken) {
          options.headers['Authorization'] = `Bearer ${this.authToken}`;
        }

        if (data && method !== 'GET') {
          if (isFormData) {
            options.body = data;
          } else {
            options.body = JSON.stringify(data);
          }
        }

        const response = await fetch(url, options);

        if (!response.ok) {
          if (response.status === 401) {
            // Token expired, refresh and retry
            if (attempt < this.retryConfig.maxRetries - 1) {
              await this._refreshToken();
              continue;
            }
          }
          throw new Error(
            `HTTP ${response.status}: ${response.statusText}`
          );
        }

        if (returnBlob) {
          return await response.blob();
        }

        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('application/json')) {
          return await response.json();
        }

        return await response.text();
      } catch (error) {
        lastError = error;

        if (attempt < this.retryConfig.maxRetries - 1) {
          const delay = this.retryConfig.retryDelay *
            Math.pow(this.retryConfig.backoffMultiplier, attempt);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }

    throw lastError || new Error('Request failed');
  }

  /**
   * Get request headers for API calls
   * @private
   * @param {boolean} isFormData
   * @returns {Object} Headers object
   */
  _getRequestHeaders(isFormData = false) {
    const headers = {};

    if (!isFormData) {
      headers['Content-Type'] = 'application/json';
    }

    headers['Accept'] = 'application/json';
    headers['User-Agent'] = 'BIPTemplateBuilder/1.0';

    return headers;
  }

  /**
   * Normalize server URL
   * @private
   * @param {string} url
   * @returns {string} Normalized URL
   */
  _normalizeServerUrl(url) {
    let normalized = url.trim();

    if (!normalized.startsWith('http')) {
      normalized = `http://${normalized}`;
    }

    if (normalized.endsWith('/')) {
      normalized = normalized.slice(0, -1);
    }

    return normalized;
  }

  /**
   * Check if token is near expiration
   * @private
   * @returns {boolean}
   */
  _isTokenNearExpiration() {
    if (!this.tokenExpiresAt) return false;
    const fiveMinutesFromNow = new Date(Date.now() + 5 * 60 * 1000);
    return this.tokenExpiresAt < fiveMinutesFromNow;
  }

  /**
   * Check if connection is active
   * @private
   * @throws {Error} If not connected
   */
  _checkConnection() {
    if (!this.isConnected) {
      throw new Error('Not connected to BI Publisher server');
    }
  }
}

// Export
export default BIPConnection;
export { BIPConnection };
