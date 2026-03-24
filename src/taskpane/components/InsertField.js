/**
 * InsertField Component
 * Hierarchical tree view of XML data fields with search, properties panel,
 * and batch insert capabilities for BI Publisher templates.
 */

class InsertField {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.dataFields = [];
    this.selectedFields = [];
    this.batchMode = false;
    this.expandedNodes = new Set();
    this.filteredFields = [];
    this.currentFilter = '';
  }

  /**
   * Render the InsertField component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'insert-field-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    `;

    // Header
    const header = document.createElement('h3');
    header.textContent = 'Insert Data Fields';
    header.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(header);

    // Search and Filter
    const searchContainer = document.createElement('div');
    searchContainer.style.cssText = 'display: flex; gap: 8px;';
    
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search fields...';
    searchInput.className = 'field-search-input';
    searchInput.style.cssText = `
      flex: 1;
      padding: 8px 12px;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      font-size: 13px;
      outline: none;
    `;
    searchContainer.appendChild(searchInput);

    // Batch mode toggle
    const batchLabel = document.createElement('label');
    batchLabel.style.cssText = 'display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: 12px;';
    const batchCheckbox = document.createElement('input');
    batchCheckbox.type = 'checkbox';
    batchCheckbox.className = 'batch-mode-toggle';
    batchCheckbox.style.cssText = 'cursor: pointer;';
    batchLabel.appendChild(batchCheckbox);
    batchLabel.appendChild(document.createTextNode('Batch Mode'));
    searchContainer.appendChild(batchLabel);

    wrapper.appendChild(searchContainer);

    // Main content area - split view
    const contentArea = document.createElement('div');
    contentArea.style.cssText = `
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      flex: 1;
      overflow: hidden;
      min-height: 0;
    `;

    // Left: Tree view
    const treeContainer = document.createElement('div');
    treeContainer.className = 'field-tree-container';
    treeContainer.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      overflow-y: auto;
      overflow-x: hidden;
      background: #f9fafb;
    `;
    contentArea.appendChild(treeContainer);

    // Right: Properties panel
    const propsContainer = document.createElement('div');
    propsContainer.className = 'field-properties-container';
    propsContainer.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      padding: 12px;
      overflow-y: auto;
      background: #f9fafb;
      display: flex;
      flex-direction: column;
      gap: 12px;
    `;

    const propsHeader = document.createElement('h4');
    propsHeader.textContent = 'Field Properties';
    propsHeader.style.cssText = 'margin: 0; font-size: 12px; font-weight: 600; color: #374151;';
    propsContainer.appendChild(propsHeader);

    const propsContent = document.createElement('div');
    propsContent.className = 'properties-content';
    propsContent.style.cssText = 'flex: 1; display: flex; flex-direction: column; gap: 8px; font-size: 12px;';
    propsContent.innerHTML = '<p style="color: #9ca3af; margin: 0;">Select a field to view properties</p>';
    propsContainer.appendChild(propsContent);

    contentArea.appendChild(propsContainer);
    wrapper.appendChild(contentArea);

    // Buttons area
    const buttonArea = document.createElement('div');
    buttonArea.style.cssText = `
      display: flex;
      gap: 8px;
      padding-top: 12px;
      border-top: 1px solid #d1d5db;
    `;

    const insertBtn = document.createElement('button');
    insertBtn.textContent = 'Insert Selected';
    insertBtn.className = 'insert-field-btn';
    insertBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.2s;
    `;
    insertBtn.onmouseover = () => insertBtn.style.background = '#2563eb';
    insertBtn.onmouseout = () => insertBtn.style.background = '#3b82f6';
    buttonArea.appendChild(insertBtn);

    const insertBatchBtn = document.createElement('button');
    insertBatchBtn.textContent = 'Insert All Batch';
    insertBatchBtn.className = 'insert-batch-btn';
    insertBatchBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.2s;
      display: ${this.batchMode ? 'block' : 'none'};
    `;
    insertBatchBtn.onmouseover = () => insertBatchBtn.style.background = '#059669';
    insertBatchBtn.onmouseout = () => insertBatchBtn.style.background = '#10b981';
    buttonArea.appendChild(insertBatchBtn);

    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    closeBtn.style.cssText = `
      padding: 8px 16px;
      background: #e5e7eb;
      color: #374151;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.2s;
    `;
    closeBtn.onmouseover = () => closeBtn.style.background = '#d1d5db';
    closeBtn.onmouseout = () => closeBtn.style.background = '#e5e7eb';
    buttonArea.appendChild(closeBtn);

    wrapper.appendChild(buttonArea);
    container.appendChild(wrapper);

    // Load data and build tree
    this.loadDataFields();
    this.renderFieldTree(treeContainer);

    // Bind events
    this.bindEvents({
      container,
      searchInput,
      batchCheckbox,
      insertBtn,
      insertBatchBtn,
      closeBtn,
      treeContainer,
      propsContainer,
      propsContent
    });
  }

  /**
   * Load data fields from template or data source
   */
  loadDataFields() {
    try {
      // Use AppState.fieldTree if available (from loaded XML/XSD)
      const appState = this.services.AppState || (typeof AppState !== 'undefined' ? AppState : null);
      console.log('[InsertField] loadDataFields - AppState:', !!appState, 'fieldTree:', !!(appState && appState.fieldTree), 'loadedData:', !!(appState && appState.loadedData));

      if (appState && appState.fieldTree) {
        this.dataFields = Array.isArray(appState.fieldTree) ? appState.fieldTree : [appState.fieldTree];
        console.log('[InsertField] Using fieldTree, fields:', this.dataFields.length, JSON.stringify(this.dataFields[0]?.name));
        this.filteredFields = JSON.parse(JSON.stringify(this.dataFields));
        return;
      }

      // Fallback: if raw XML data is loaded, build tree from it
      if (appState && appState.loadedData && this.xmlParser) {
        try {
          const tree = this.xmlParser.getFieldTree(appState.loadedData);
          if (tree) {
            this.dataFields = [tree];
            console.log('[InsertField] Built tree from loadedData:', tree.name);
            this.filteredFields = JSON.parse(JSON.stringify(this.dataFields));
            return;
          }
        } catch (e) {
          console.warn('[InsertField] Failed to build tree from loadedData', e);
        }
      }

      const template = this.templateEngine ? this.templateEngine.getCurrentTemplate() : null;
      if (template && template.dataSource) {
        this.dataFields = this.xmlParser.parseToFieldTree(template.dataSource);
      } else {
        // Sample data for demo
        this.dataFields = [
          {
            name: 'Company',
            xpath: '/Company',
            type: 'group',
            children: [
              { name: 'CompanyName', xpath: '/Company/CompanyName', type: 'text', sample: 'Acme Corp' },
              { name: 'Address', xpath: '/Company/Address', type: 'text', sample: '123 Main St' },
              { name: 'Employees', xpath: '/Company/Employees', type: 'group', children: [
                { name: 'Employee', xpath: '/Company/Employees/Employee', type: 'group', isRepeating: true, children: [
                  { name: 'Name', xpath: '/Company/Employees/Employee/Name', type: 'text', sample: 'John Doe' },
                  { name: 'Salary', xpath: '/Company/Employees/Employee/Salary', type: 'number', sample: '50000', format: 'currency' },
                  { name: 'HireDate', xpath: '/Company/Employees/Employee/HireDate', type: 'date', sample: '2024-01-15', format: 'MM/DD/YYYY' }
                ]}
              ]}
            ]
          }
        ];
      }
      this.filteredFields = JSON.parse(JSON.stringify(this.dataFields));
    } catch (error) {
      console.error('Error loading data fields:', error);
      this.dataFields = [];
    }
  }

  /**
   * Render field tree structure
   */
  renderFieldTree(container, fields = null, depth = 0) {
    if (!fields) fields = this.filteredFields;
    
    fields.forEach((field, index) => {
      const fieldItem = document.createElement('div');
      fieldItem.className = 'field-tree-item';
      fieldItem.style.cssText = `
        padding-left: ${depth * 20}px;
        border-bottom: 1px solid #e5e7eb;
        user-select: none;
      `;

      const itemContent = document.createElement('div');
      itemContent.style.cssText = `
        display: flex;
        align-items: center;
        padding: 8px 8px;
        gap: 6px;
        cursor: pointer;
        transition: background 0.15s;
      `;
      itemContent.onmouseover = () => itemContent.style.background = '#f3f4f6';
      itemContent.onmouseout = () => itemContent.style.background = 'transparent';

      // Expand/collapse arrow
      const arrow = document.createElement('span');
      arrow.className = 'field-expand-arrow';
      arrow.textContent = field.children ? (this.expandedNodes.has(field.xpath) ? '▼' : '▶') : '•';
      arrow.style.cssText = `
        width: 18px;
        display: inline-flex;
        justify-content: center;
        font-size: 11px;
        color: #6b7280;
        cursor: pointer;
      `;
      itemContent.appendChild(arrow);

      // Field label
      const label = document.createElement('span');
      label.textContent = field.name;
      label.style.cssText = `
        flex: 1;
        font-size: 12px;
        color: #1f2937;
        font-weight: ${field.type === 'group' ? '500' : '400'};
      `;
      itemContent.appendChild(label);

      // Field type badge
      const typeBadge = document.createElement('span');
      const typeColors = {
        'text': '#d1d5db',
        'number': '#dbeafe',
        'date': '#fcd34d',
        'group': '#e0e7ff'
      };
      typeBadge.textContent = field.type;
      typeBadge.style.cssText = `
        font-size: 10px;
        padding: 2px 6px;
        background: ${typeColors[field.type] || '#e5e7eb'};
        color: #374151;
        border-radius: 3px;
        white-space: nowrap;
      `;
      if (field.type !== 'group') itemContent.appendChild(typeBadge);

      // Checkbox for batch mode
      if (this.batchMode && field.type !== 'group') {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'field-batch-checkbox';
        checkbox.style.cssText = 'cursor: pointer; margin-left: 4px;';
        checkbox.dataset.xpath = field.xpath;
        itemContent.appendChild(checkbox);
      }

      fieldItem.appendChild(itemContent);

      // Handle expand/collapse
      arrow.addEventListener('click', (e) => {
        e.stopPropagation();
        if (field.children) {
          if (this.expandedNodes.has(field.xpath)) {
            this.expandedNodes.delete(field.xpath);
            arrow.textContent = '▶';
            const childContainer = fieldItem.nextElementSibling;
            if (childContainer && childContainer.classList.contains('field-children')) {
              childContainer.remove();
            }
          } else {
            this.expandedNodes.add(field.xpath);
            arrow.textContent = '▼';
            const childContainer = document.createElement('div');
            childContainer.className = 'field-children';
            this.renderFieldTree(childContainer, field.children, depth + 1);
            fieldItem.parentNode.insertBefore(childContainer, fieldItem.nextSibling);
          }
        }
      });

      // Handle field selection
      itemContent.addEventListener('click', () => {
        if (field.type !== 'group') {
          this.displayFieldProperties(field);
          if (this.batchMode) {
            const checkbox = itemContent.querySelector('.field-batch-checkbox');
            if (checkbox) checkbox.checked = !checkbox.checked;
          }
        }
      });

      // Context menu
      itemContent.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        this.showFieldContextMenu(e, field);
      });

      container.appendChild(fieldItem);
    });
  }

  /**
   * Display field properties
   */
  displayFieldProperties(field) {
    const propsContent = document.querySelector('.properties-content');
    if (!propsContent) return;

    let html = '';
    html += `<div><strong>Name:</strong> ${field.name}</div>`;
    html += `<div><strong>XPath:</strong> <code style="background: #e5e7eb; padding: 2px 4px; border-radius: 2px; font-size: 11px; word-break: break-all;">${field.xpath}</code></div>`;
    html += `<div><strong>Type:</strong> ${field.type}</div>`;
    
    if (field.sample) {
      html += `<div><strong>Sample:</strong> ${field.sample}</div>`;
    }
    if (field.format) {
      html += `<div><strong>Format:</strong> ${field.format}</div>`;
    }
    if (field.isRepeating) {
      html += `<div style="padding: 6px; background: #fef3c7; border-radius: 3px; color: #92400e; font-size: 11px;">⚠ Repeating element</div>`;
    }

    // Format options
    if (field.type === 'number' || field.type === 'date') {
      html += `<div style="margin-top: 8px; padding-top: 8px; border-top: 1px solid #d1d5db;">`;
      html += `<label style="display: block; margin-bottom: 4px;"><strong>Format:</strong></label>`;
      html += `<select style="width: 100%; padding: 4px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;">`;
      
      if (field.type === 'number') {
        html += `<option>Standard</option><option>Currency</option><option>Percent</option><option>Scientific</option>`;
      } else if (field.type === 'date') {
        html += `<option>MM/DD/YYYY</option><option>DD/MM/YYYY</option><option>YYYY-MM-DD</option><option>Full Date</option>`;
      }
      
      html += `</select></div>`;
    }

    // Default value and null handling
    html += `<div style="margin-top: 8px; padding-top: 8px; border-top: 1px solid #d1d5db;">`;
    html += `<label style="display: block; font-size: 11px; margin-bottom: 4px;">Default Value:</label>`;
    html += `<input type="text" style="width: 100%; padding: 4px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;" placeholder="Leave blank if none">`;
    html += `</div>`;

    html += `<div style="margin-top: 8px;">`;
    html += `<label style="display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: 11px;">`;
    html += `<input type="checkbox"> Display blank if null`;
    html += `</label>`;
    html += `</div>`;

    propsContent.innerHTML = html;
  }

  /**
   * Show context menu for field
   */
  showFieldContextMenu(event, field) {
    const menu = document.createElement('div');
    menu.style.cssText = `
      position: fixed;
      top: ${event.clientY}px;
      left: ${event.clientX}px;
      background: white;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      z-index: 1000;
      min-width: 180px;
    `;

    const items = [
      { label: 'Insert Field', action: () => this.insertField(field) },
      { label: 'Copy XPath', action: () => this.copyToClipboard(field.xpath) },
      { label: 'View Properties', action: () => this.displayFieldProperties(field) }
    ];

    items.forEach(item => {
      const menuItem = document.createElement('div');
      menuItem.textContent = item.label;
      menuItem.style.cssText = `
        padding: 8px 12px;
        cursor: pointer;
        font-size: 12px;
        color: #1f2937;
        transition: background 0.15s;
      `;
      menuItem.onmouseover = () => menuItem.style.background = '#f3f4f6';
      menuItem.onmouseout = () => menuItem.style.background = 'transparent';
      menuItem.onclick = () => {
        item.action();
        menu.remove();
      };
      menu.appendChild(menuItem);
    });

    document.body.appendChild(menu);
    setTimeout(() => {
      document.addEventListener('click', () => menu.remove(), { once: true });
    }, 0);
  }

  /**
   * Insert a field into the document
   */
  async insertField(field) {
    try {
      const fieldTag = `<?${field.xpath}?>`;
      await this.wordApi.insertText(fieldTag);
    } catch (error) {
      console.error('Error inserting field:', error);
    }
  }

  /**
   * Copy text to clipboard
   */
  copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
      console.log('Copied to clipboard:', text);
    });
  }

  /**
   * Get selected fields for batch insert
   */
  getSelectedFields() {
    const checkboxes = document.querySelectorAll('.field-batch-checkbox:checked');
    const selected = [];
    checkboxes.forEach(cb => {
      selected.push(cb.dataset.xpath);
    });
    return selected;
  }

  /**
   * Bind events
   */
  bindEvents(elements) {
    const {
      searchInput,
      batchCheckbox,
      insertBtn,
      insertBatchBtn,
      closeBtn,
      treeContainer,
      propsContainer
    } = elements;

    // Search/filter
    searchInput.addEventListener('input', (e) => {
      this.currentFilter = e.target.value.toLowerCase();
      this.applyFilter();
      treeContainer.innerHTML = '';
      this.renderFieldTree(treeContainer);
    });

    // Batch mode toggle
    batchCheckbox.addEventListener('change', (e) => {
      this.batchMode = e.target.checked;
      insertBatchBtn.style.display = this.batchMode ? 'block' : 'none';
      treeContainer.innerHTML = '';
      this.renderFieldTree(treeContainer);
    });

    // Insert selected
    insertBtn.addEventListener('click', async () => {
      if (this.batchMode) {
        const selected = this.getSelectedFields();
        for (const xpath of selected) {
          const field = this.findFieldByXpath(xpath);
          if (field) await this.insertField(field);
        }
      } else {
        const selected = document.querySelector('.field-tree-item.selected');
        if (selected) {
          console.log('Insert selected field');
        }
      }
    });

    // Insert all batch
    insertBatchBtn.addEventListener('click', async () => {
      const selected = this.getSelectedFields();
      for (const xpath of selected) {
        const field = this.findFieldByXpath(xpath);
        if (field) await this.insertField(field);
      }
    });

    // Close
    closeBtn.addEventListener('click', () => {
      elements.container.innerHTML = '';
    });
  }

  /**
   * Apply filter to fields
   */
  applyFilter() {
    this.filteredFields = this.filterFields(this.dataFields, this.currentFilter);
  }

  /**
   * Recursively filter fields
   */
  filterFields(fields, filter) {
    if (!filter) return JSON.parse(JSON.stringify(fields));
    
    return fields
      .filter(f => f.name.toLowerCase().includes(filter) || f.xpath.toLowerCase().includes(filter))
      .map(f => {
        const copy = { ...f };
        if (f.children) {
          copy.children = this.filterFields(f.children, filter);
        }
        return copy;
      });
  }

  /**
   * Find field by XPath
   */
  findFieldByXpath(xpath) {
    const search = (fields) => {
      for (const field of fields) {
        if (field.xpath === xpath) return field;
        if (field.children) {
          const found = search(field.children);
          if (found) return found;
        }
      }
      return null;
    };
    return search(this.dataFields);
  }
}

export default InsertField;
export { InsertField };
