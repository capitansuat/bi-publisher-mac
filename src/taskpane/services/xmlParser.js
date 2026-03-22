/**
 * XML Data Parser and XPath Utilities
 * Parses XML structures, extracts field information, and provides XPath helpers
 * for BI Publisher template building
 *
 * @module xmlParser
 */

/**
 * XMLDataParser class handles XML parsing and XPath operations
 * Supports XML namespaces, large files, and schema validation
 */
class XMLDataParser {
  /**
   * Creates a new XML parser instance
   */
  constructor() {
    this.parsedXML = null;
    this.parsedXSD = null;
    this.namespaces = {};
    this.fieldCache = new Map();
  }

  /**
   * Parse XML string into DOM tree structure
   * @param {string} xmlString - XML content as string
   * @returns {Document} Parsed XML document
   * @throws {Error} If XML is malformed
   *
   * @example
   * const xmlDoc = xmlParser.parseXML('<root><item>value</item></root>');
   */
  parseXML(xmlString) {
    try {
      if (!xmlString || typeof xmlString !== 'string') {
        throw new Error('Invalid XML input: must be a non-empty string');
      }

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'text/xml');

      // Check for parse errors
      if (xmlDoc.parseError && xmlDoc.parseError.errorCode !== 0) {
        throw new Error(
          `XML Parse Error: ${xmlDoc.parseError.reason} at line ${xmlDoc.parseError.line}`
        );
      }

      // Check for error elements (alternative error detection)
      if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
        throw new Error('XML parsing failed: malformed XML content');
      }

      this.parsedXML = xmlDoc;
      this._extractNamespaces(xmlDoc.documentElement);

      return xmlDoc;
    } catch (error) {
      throw new Error(`XML parsing failed: ${error.message}`);
    }
  }

  /**
   * Parse XML Schema (XSD) definition
   * @param {string} xsdString - XSD content as string
   * @returns {Object} Schema structure with elements and types
   * @throws {Error} If XSD is invalid
   *
   * @example
   * const schema = xmlParser.parseXSD('<xs:schema>...</xs:schema>');
   */
  parseXSD(xsdString) {
    try {
      const parser = new DOMParser();
      const xsdDoc = parser.parseFromString(xsdString, 'text/xml');

      if (xsdDoc.parseError && xsdDoc.parseError.errorCode !== 0) {
        throw new Error('XSD Parse Error');
      }

      const schema = {
        targetNamespace: xsdDoc.documentElement.getAttribute('targetNamespace'),
        elements: [],
        complexTypes: [],
        simpleTypes: []
      };

      // Extract element definitions
      const elements = xsdDoc.getElementsByTagName('xs:element');
      for (let el of elements) {
        schema.elements.push({
          name: el.getAttribute('name'),
          type: el.getAttribute('type'),
          minOccurs: el.getAttribute('minOccurs') || '1',
          maxOccurs: el.getAttribute('maxOccurs') || '1'
        });
      }

      // Extract complex types
      const complexTypes = xsdDoc.getElementsByTagName('xs:complexType');
      for (let ct of complexTypes) {
        schema.complexTypes.push({
          name: ct.getAttribute('name'),
          sequence: this._parseSequence(ct)
        });
      }

      // Extract simple types
      const simpleTypes = xsdDoc.getElementsByTagName('xs:simpleType');
      for (let st of simpleTypes) {
        schema.simpleTypes.push({
          name: st.getAttribute('name'),
          restriction: st.querySelector('xs:restriction')?.getAttribute('base')
        });
      }

      this.parsedXSD = schema;
      return schema;
    } catch (error) {
      throw new Error(`XSD parsing failed: ${error.message}`);
    }
  }

  /**
   * Extract all XPath field expressions from XML
   * @param {Document|Element} xmlTree - XML document or element
   * @returns {Array<string>} Array of XPath expressions
   * @throws {Error} If extraction fails
   *
   * @example
   * const paths = xmlParser.getFieldPaths(xmlDoc);
   * // Returns: ['ROOT/PRODUCT', 'ROOT/PRODUCT/ID', 'ROOT/PRODUCT/NAME', ...]
   */
  getFieldPaths(xmlTree) {
    try {
      const paths = [];
      const root = xmlTree.documentElement || xmlTree;
      const rootPath = root.nodeName;

      const traverse = (node, currentPath) => {
        if (node.nodeType === 1) { // Element node
          const nodePath = currentPath ? `${currentPath}/${node.nodeName}` : node.nodeName;

          // Add path if element contains data
          if (node.childNodes.length === 1 && node.childNodes[0].nodeType === 3) {
            paths.push(nodePath);
          } else if (node.childNodes.length === 0) {
            // Empty element
            paths.push(nodePath);
          }

          // Traverse children
          for (let i = 0; i < node.childNodes.length; i++) {
            traverse(node.childNodes[i], nodePath);
          }
        }
      };

      traverse(root, '');
      return [...new Set(paths)]; // Remove duplicates
    } catch (error) {
      throw new Error(`Field path extraction failed: ${error.message}`);
    }
  }

  /**
   * Build hierarchical field tree for UI display
   * @param {Document|Element} xmlTree - XML document or element
   * @returns {Object} Nested tree structure
   * @throws {Error} If tree building fails
   *
   * @example
   * const tree = xmlParser.getFieldTree(xmlDoc);
   * // Returns: {name: 'ROOT', children: [{name: 'PRODUCT', children: [...]}]}
   */
  getFieldTree(xmlTree) {
    try {
      const root = xmlTree.documentElement || xmlTree;

      const buildNode = (element, depth = 0) => {
        // Prevent deeply nested trees
        if (depth > 50) return null;

        const node = {
          name: element.nodeName,
          xpath: this._getElementXPath(element),
          type: this._inferElementType(element),
          attributes: Array.from(element.attributes || []).map(attr => ({
            name: attr.name,
            value: attr.value
          })),
          isRepeating: this._isRepeatingElement(element),
          children: []
        };

        // Add sample value if leaf node
        if (element.childNodes.length === 1 &&
            element.childNodes[0].nodeType === 3) {
          node.sampleValue = element.textContent.substring(0, 100);
        }

        // Build children
        const childMap = new Map();
        for (let child of element.childNodes) {
          if (child.nodeType === 1) { // Element node
            const childName = child.nodeName;
            if (!childMap.has(childName)) {
              childMap.set(childName, buildNode(child, depth + 1));
            }
          }
        }

        node.children = Array.from(childMap.values()).filter(c => c !== null);
        return node;
      };

      return buildNode(root);
    } catch (error) {
      throw new Error(`Field tree building failed: ${error.message}`);
    }
  }

  /**
   * Get inferred data type of a field
   * @param {string} xpath - XPath expression
   * @returns {string} Data type (string, number, date, boolean, decimal)
   * @throws {Error} If field not found
   *
   * @example
   * const type = xmlParser.getFieldType('PRODUCTS/PRICE');
   */
  getFieldType(xpath) {
    try {
      if (this.fieldCache.has(`type:${xpath}`)) {
        return this.fieldCache.get(`type:${xpath}`);
      }

      const element = this._getElementByXPath(xpath);
      if (!element) {
        return 'unknown';
      }

      const type = this._inferElementType(element);
      this.fieldCache.set(`type:${xpath}`, type);

      return type;
    } catch (error) {
      console.warn(`Failed to determine type for ${xpath}:`, error.message);
      return 'unknown';
    }
  }

  /**
   * Get sample values for a field from XML
   * @param {Document|Element} xmlTree - XML document
   * @param {string} xpath - XPath expression
   * @param {number} [limit=5] - Maximum samples to return
   * @returns {Array<any>} Sample values
   * @throws {Error} If retrieval fails
   *
   * @example
   * const samples = xmlParser.getSampleData(xmlDoc, 'PRODUCTS/PRICE', 3);
   */
  getSampleData(xmlTree, xpath, limit = 5) {
    try {
      const samples = [];
      const parts = xpath.split('/').filter(p => p);

      if (parts.length === 0) {
        return samples;
      }

      const root = xmlTree.documentElement || xmlTree;
      const traverse = (node, pathIndex) => {
        if (pathIndex === parts.length) {
          const value = node.textContent.trim();
          if (value && !samples.includes(value) && samples.length < limit) {
            samples.push(value);
          }
          return;
        }

        const currentPart = parts[pathIndex];
        for (let child of node.childNodes) {
          if (child.nodeType === 1 && child.nodeName === currentPart) {
            traverse(child, pathIndex + 1);
          }
        }
      };

      traverse(root, 0);
      return samples;
    } catch (error) {
      throw new Error(`Sample data retrieval failed: ${error.message}`);
    }
  }

  /**
   * Validate XPath expression for correctness
   * @param {string} xpath - XPath expression to validate
   * @returns {Object} Validation result with {valid: boolean, error?: string}
   *
   * @example
   * const result = xmlParser.validateXPath('PRODUCTS/PRODUCT/NAME');
   */
  validateXPath(xpath) {
    try {
      if (!xpath || typeof xpath !== 'string') {
        return { valid: false, error: 'XPath must be a non-empty string' };
      }

      // Basic validation rules
      const parts = xpath.split('/').filter(p => p);

      if (parts.length === 0) {
        return { valid: false, error: 'XPath cannot be empty' };
      }

      // Check for valid XML element names
      const validNameRegex = /^[a-zA-Z_:][a-zA-Z0-9_:.\-]*$/;
      for (let part of parts) {
        if (!validNameRegex.test(part)) {
          return {
            valid: false,
            error: `Invalid element name: '${part}'`
          };
        }
      }

      // Check for balanced brackets and quotes
      const bracketCount = (xpath.match(/\[/g) || []).length;
      const closeBracketCount = (xpath.match(/\]/g) || []).length;

      if (bracketCount !== closeBracketCount) {
        return { valid: false, error: 'Unbalanced brackets in XPath' };
      }

      return { valid: true };
    } catch (error) {
      return { valid: false, error: error.message };
    }
  }

  /**
   * Build XPath from tree node selection
   * @param {Object} treeNode - Node from field tree
   * @param {string} [parentPath=''] - Parent XPath (used recursively)
   * @returns {string} Generated XPath
   *
   * @example
   * const xpath = xmlParser.buildXPathFromTree(treeNode);
   */
  buildXPathFromTree(treeNode, parentPath = '') {
    if (!treeNode || !treeNode.name) {
      return parentPath;
    }

    const currentPath = parentPath
      ? `${parentPath}/${treeNode.name}`
      : treeNode.name;

    // If node has children, optionally append first child
    if (treeNode.children && treeNode.children.length > 0) {
      return currentPath;
    }

    return currentPath;
  }

  /**
   * Transform XML tree into flat field list
   * @param {Document|Element} xmlTree - XML document
   * @returns {Array<Object>} Flat array of fields with labels
   *
   * @example
   * const fields = xmlParser.transformToDataFields(xmlDoc);
   */
  transformToDataFields(xmlTree) {
    try {
      const fields = [];
      const paths = this.getFieldPaths(xmlTree);

      paths.forEach(xpath => {
        const type = this.getFieldType(xpath);
        const sampleData = this.getSampleData(xmlTree, xpath, 1);

        fields.push({
          name: this._xpathToFieldName(xpath),
          xpath,
          type,
          label: this._xpathToLabel(xpath),
          sampleValue: sampleData[0] || '',
          isRequired: false,
          description: ``
        });
      });

      return fields;
    } catch (error) {
      throw new Error(`Field transformation failed: ${error.message}`);
    }
  }

  /**
   * Get repeating elements (groups) in XML
   * @param {Document|Element} xmlTree - XML document
   * @returns {Array<Object>} Repeating element definitions
   *
   * @example
   * const groups = xmlParser.getRepeatingElements(xmlDoc);
   */
  getRepeatingElements(xmlTree) {
    try {
      const repeating = [];
      const root = xmlTree.documentElement || xmlTree;
      const elementCounts = new Map();

      // Count element occurrences
      const countElements = (node, path) => {
        if (node.nodeType === 1) {
          const nodePath = path ? `${path}/${node.nodeName}` : node.nodeName;
          const currentCount = elementCounts.get(nodePath) || 0;
          elementCounts.set(nodePath, currentCount + 1);

          for (let child of node.childNodes) {
            countElements(child, nodePath);
          }
        }
      };

      countElements(root, '');

      // Find paths with count > 1 (repeating)
      for (let [path, count] of elementCounts) {
        if (count > 1) {
          repeating.push({
            xpath: path,
            occurrences: count,
            isRepeating: true,
            label: this._xpathToLabel(path)
          });
        }
      }

      return repeating.sort((a, b) => b.occurrences - a.occurrences);
    } catch (error) {
      throw new Error(`Repeating element detection failed: ${error.message}`);
    }
  }

  /**
   * Get attributes of an element
   * @param {string} xpath - XPath to element
   * @returns {Array<Object>} Attribute definitions
   *
   * @example
   * const attrs = xmlParser.getAttributes('PRODUCTS/PRODUCT');
   */
  getAttributes(xpath) {
    try {
      const element = this._getElementByXPath(xpath);
      if (!element) {
        return [];
      }

      const attributes = [];
      for (let attr of element.attributes || []) {
        attributes.push({
          name: attr.name,
          value: attr.value,
          xpath: `${xpath}/@${attr.name}`
        });
      }

      return attributes;
    } catch (error) {
      throw new Error(`Attribute retrieval failed: ${error.message}`);
    }
  }

  /**
   * Clear internal cache
   * @returns {void}
   */
  clearCache() {
    this.fieldCache.clear();
  }

  // ============= Private Helper Methods =============

  /**
   * Extract namespace declarations from element
   * @private
   * @param {Element} element
   * @returns {void}
   */
  _extractNamespaces(element) {
    for (let attr of element.attributes || []) {
      if (attr.name.startsWith('xmlns')) {
        const prefix = attr.name === 'xmlns' ? 'default' : attr.name.split(':')[1];
        this.namespaces[prefix] = attr.value;
      }
    }
  }

  /**
   * Get element by XPath (simple implementation)
   * @private
   * @param {string} xpath
   * @returns {Element|null}
   */
  _getElementByXPath(xpath) {
    if (!this.parsedXML) {
      return null;
    }

    try {
      // Create evaluator
      const evaluator = new XPathEvaluator();
      const result = evaluator.evaluate(
        xpath,
        this.parsedXML,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null
      );

      return result.singleNodeValue;
    } catch (error) {
      // Fallback to simple path traversal
      const parts = xpath.split('/').filter(p => p);
      let node = this.parsedXML.documentElement;

      for (let part of parts) {
        let found = false;
        for (let child of node.childNodes) {
          if (child.nodeType === 1 && child.nodeName === part) {
            node = child;
            found = true;
            break;
          }
        }
        if (!found) return null;
      }

      return node;
    }
  }

  /**
   * Get XPath of an element
   * @private
   * @param {Element} element
   * @returns {string}
   */
  _getElementXPath(element) {
    const paths = [];
    let node = element;

    while (node && node.nodeType === 1) {
      let index = 1;
      let sibling = node.previousSibling;

      while (sibling) {
        if (sibling.nodeType === 1 && sibling.nodeName === node.nodeName) {
          index++;
        }
        sibling = sibling.previousSibling;
      }

      paths.unshift(node.nodeName);
      node = node.parentNode;
    }

    return paths.join('/');
  }

  /**
   * Infer data type from element
   * @private
   * @param {Element} element
   * @returns {string}
   */
  _inferElementType(element) {
    const value = element.textContent.trim();

    if (!value) return 'string';
    if (/^\d+$/.test(value)) return 'integer';
    if (/^\d+\.\d+$/.test(value)) return 'decimal';
    if (/^\d{4}-\d{2}-\d{2}/.test(value)) return 'date';
    if (/^\d{2}:\d{2}:\d{2}/.test(value)) return 'time';
    if (/^(true|false)$/i.test(value)) return 'boolean';

    return 'string';
  }

  /**
   * Check if element repeats in parent
   * @private
   * @param {Element} element
   * @returns {boolean}
   */
  _isRepeatingElement(element) {
    const parent = element.parentNode;
    if (!parent) return false;

    let count = 0;
    for (let child of parent.childNodes) {
      if (child.nodeType === 1 && child.nodeName === element.nodeName) {
        count++;
        if (count > 1) return true;
      }
    }

    return false;
  }

  /**
   * Convert XPath to field name
   * @private
   * @param {string} xpath
   * @returns {string}
   */
  _xpathToFieldName(xpath) {
    const parts = xpath.split('/');
    return parts[parts.length - 1] || xpath;
  }

  /**
   * Convert XPath to display label
   * @private
   * @param {string} xpath
   * @returns {string}
   */
  _xpathToLabel(xpath) {
    const fieldName = this._xpathToFieldName(xpath);
    return fieldName
      .replace(/_/g, ' ')
      .replace(/([A-Z])/g, ' $1')
      .trim()
      .split(/\s+/)
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }

  /**
   * Parse XSD sequence element
   * @private
   * @param {Element} complexType
   * @returns {Array}
   */
  _parseSequence(complexType) {
    const sequence = complexType.querySelector('xs:sequence');
    if (!sequence) return [];

    const elements = [];
    for (let el of sequence.querySelectorAll('xs:element')) {
      elements.push({
        name: el.getAttribute('name'),
        type: el.getAttribute('type'),
        minOccurs: el.getAttribute('minOccurs') || '0',
        maxOccurs: el.getAttribute('maxOccurs') || '1'
      });
    }

    return elements;
  }
}

// Export
export default XMLDataParser;
export { XMLDataParser };
