/**
 * RepeatingGroup Component
 * Component for inserting repeating groups with sorting, grouping headers/footers,
 * page breaks, and running counts/totals for BI Publisher templates.
 */

class RepeatingGroup {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.selectedElement = null;
    this.groupByField = null;
    this.sortConfig = [];
    this.headerOptions = {
      showHeader: false,
      headerText: ''
    };
    this.footerOptions = {
      showFooter: false,
      footerText: ''
    };
    this.pageBreakBetween = false;
    this.runningCount = false;
    this.runningTotal = null;
    this.dataElements = [];
    this.nestedGroups = [];
  }

  /**
   * Render the RepeatingGroup component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'repeating-group-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Repeating Group';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Element selection
    const selectionSection = document.createElement('fieldset');
    selectionSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const selectionLegend = document.createElement('legend');
    selectionLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    selectionLegend.textContent = 'Repeating Element';
    selectionSection.appendChild(selectionLegend);

    const selectionContent = document.createElement('div');
    selectionContent.style.cssText = 'display: flex; gap: 8px; font-size: 12px;';

    const elementSelect = document.createElement('select');
    elementSelect.style.cssText = 'flex: 1; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px;';
    this.loadDataElements();
    elementSelect.innerHTML = '<option value="">-- Select element --</option>';
    this.dataElements.forEach(el => {
      elementSelect.innerHTML += `<option value="${el.xpath}">${el.name}</option>`;
    });
    elementSelect.addEventListener('change', (e) => {
      this.selectedElement = e.target.value;
    });
    selectionContent.appendChild(elementSelect);

    const browseBtn = document.createElement('button');
    browseBtn.textContent = 'Browse';
    browseBtn.style.cssText = `
      padding: 8px 12px;
      background: #e5e7eb;
      color: #374151;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
    `;
    selectionContent.appendChild(browseBtn);

    selectionSection.appendChild(selectionContent);
    wrapper.appendChild(selectionSection);

    // Group by field
    const groupSection = document.createElement('fieldset');
    groupSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const groupLegend = document.createElement('legend');
    groupLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    groupLegend.textContent = 'Group By';
    groupSection.appendChild(groupLegend);

    const groupSelect = document.createElement('select');
    groupSelect.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
    groupSelect.innerHTML = '<option value="">-- No Grouping --</option>';
    groupSelect.addEventListener('change', (e) => {
      this.groupByField = e.target.value;
    });
    groupSection.appendChild(groupSelect);

    wrapper.appendChild(groupSection);

    // Sort configuration
    const sortSection = document.createElement('fieldset');
    sortSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const sortLegend = document.createElement('legend');
    sortLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    sortLegend.textContent = 'Sort Order';
    sortSection.appendChild(sortLegend);

    const sortContent = document.createElement('div');
    sortContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const sortFields = [
      { label: 'Primary Sort', placeholder: 'Field 1' },
      { label: 'Secondary Sort', placeholder: 'Field 2' },
      { label: 'Tertiary Sort', placeholder: 'Field 3' }
    ];

    sortFields.forEach((field, index) => {
      const row = document.createElement('div');
      row.style.cssText = 'display: grid; grid-template-columns: 100px 1fr auto; gap: 8px; align-items: center;';

      const label = document.createElement('label');
      label.textContent = field.label + ':';
      label.style.cssText = 'font-size: 11px; color: #6b7280; white-space: nowrap;';
      row.appendChild(label);

      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = field.placeholder;
      input.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
      input.dataset.sortIndex = index;
      row.appendChild(input);

      const orderSelect = document.createElement('select');
      orderSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px; width: 80px;';
      orderSelect.innerHTML = '<option value="asc">Ascending</option><option value="desc">Descending</option>';
      orderSelect.dataset.sortIndex = index;
      row.appendChild(orderSelect);

      sortContent.appendChild(row);
    });

    sortSection.appendChild(sortContent);
    wrapper.appendChild(sortSection);

    // Header and Footer options
    const contentSection = document.createElement('fieldset');
    contentSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const contentLegend = document.createElement('legend');
    contentLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    contentLegend.textContent = 'Group Header & Footer';
    contentSection.appendChild(contentLegend);

    const contentDiv = document.createElement('div');
    contentDiv.style.cssText = 'display: flex; flex-direction: column; gap: 10px;';

    const headerCheckbox = document.createElement('label');
    headerCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
    const headerInput = document.createElement('input');
    headerInput.type = 'checkbox';
    headerInput.addEventListener('change', (e) => {
      this.headerOptions.showHeader = e.target.checked;
    });
    headerCheckbox.appendChild(headerInput);
    headerCheckbox.appendChild(document.createTextNode('Show group header'));
    contentDiv.appendChild(headerCheckbox);

    const footerCheckbox = document.createElement('label');
    footerCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
    const footerInput = document.createElement('input');
    footerInput.type = 'checkbox';
    footerInput.addEventListener('change', (e) => {
      this.footerOptions.showFooter = e.target.checked;
    });
    footerCheckbox.appendChild(footerInput);
    footerCheckbox.appendChild(document.createTextNode('Show group footer'));
    contentDiv.appendChild(footerCheckbox);

    const pageBreakCheckbox = document.createElement('label');
    pageBreakCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
    const pageBreakInput = document.createElement('input');
    pageBreakInput.type = 'checkbox';
    pageBreakInput.addEventListener('change', (e) => {
      this.pageBreakBetween = e.target.checked;
    });
    pageBreakCheckbox.appendChild(pageBreakInput);
    pageBreakCheckbox.appendChild(document.createTextNode('Page break between groups'));
    contentDiv.appendChild(pageBreakCheckbox);

    const runningCountCheckbox = document.createElement('label');
    runningCountCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
    const runningCountInput = document.createElement('input');
    runningCountInput.type = 'checkbox';
    runningCountInput.addEventListener('change', (e) => {
      this.runningCount = e.target.checked;
    });
    runningCountCheckbox.appendChild(runningCountInput);
    runningCountCheckbox.appendChild(document.createTextNode('Show running count'));
    contentDiv.appendChild(runningCountCheckbox);

    contentSection.appendChild(contentDiv);
    wrapper.appendChild(contentSection);

    // Action buttons
    const buttonArea = document.createElement('div');
    buttonArea.style.cssText = `
      display: flex;
      gap: 8px;
      padding-top: 12px;
      border-top: 1px solid #d1d5db;
    `;

    const wrapBtn = document.createElement('button');
    wrapBtn.textContent = 'Wrap Selection';
    wrapBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
    `;
    wrapBtn.addEventListener('click', () => this.wrapSelection());
    buttonArea.appendChild(wrapBtn);

    const insertBtn = document.createElement('button');
    insertBtn.textContent = 'Insert Group';
    insertBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
    `;
    insertBtn.addEventListener('click', () => this.insertGroup());
    buttonArea.appendChild(insertBtn);

    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    closeBtn.style.cssText = `
      padding: 8px 16px;
      background: #e5e7eb;
      color: #374151;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    `;
    closeBtn.addEventListener('click', () => container.innerHTML = '');
    buttonArea.appendChild(closeBtn);

    wrapper.appendChild(buttonArea);
    container.appendChild(wrapper);
  }

  /**
   * Load available data elements
   */
  loadDataElements() {
    // Load repeating elements from AppState.fieldTree (loaded XML)
    this.dataElements = [];
    const tree = this.services.AppState ? this.services.AppState.fieldTree : null;
    if (!tree) return;

    // Find all nodes that have children (potential repeating groups)
    const collectGroups = (node, path) => {
      if (!node) return;
      const currentPath = path ? `${path}/${node.name}` : node.name;
      if (node.children && node.children.length > 0) {
        this.dataElements.push({
          name: node.name,
          xpath: currentPath,
          isRepeating: true
        });
        node.children.forEach(child => collectGroups(child, currentPath));
      }
    };

    if (Array.isArray(tree)) {
      tree.forEach(node => collectGroups(node, ''));
    } else {
      collectGroups(tree, '');
    }
  }

  /**
   * Wrap selected content
   */
  async wrapSelection() {
    try {
      if (!this.selectedElement) {
        alert('Please select a repeating element');
        return;
      }

      let groupXml = `<?for-each:${this.selectedElement}?>\n`;
      groupXml += '  <!-- Group content -->\n';
      groupXml += '<?end for-each?>\n';

      await this.wordApi.insertText(groupXml);
    } catch (error) {
      console.error('Error wrapping selection:', error);
    }
  }

  /**
   * Insert a new repeating group
   */
  async insertGroup() {
    try {
      if (!this.selectedElement) {
        alert('Please select a repeating element');
        return;
      }

      let groupXml = `<?for-each:${this.selectedElement}?>\n`;

      // Add header if enabled
      if (this.headerOptions.showHeader) {
        groupXml += '  <w:p>\n';
        groupXml += '    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>\n';
        groupXml += `    <w:r><w:t>Group: <?${this.groupByField || 'value'}?></w:t></w:r>\n`;
        groupXml += '  </w:p>\n';
      }

      groupXml += '  <!-- Repeating content goes here -->\n';

      // Add footer if enabled
      if (this.footerOptions.showFooter) {
        groupXml += '  <w:p>\n';
        groupXml += '    <w:r><w:t>End of group</w:t></w:r>\n';
        groupXml += '  </w:p>\n';
      }

      if (this.pageBreakBetween) {
        groupXml += '  <w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>\n';
      }

      if (this.runningCount) {
        groupXml += '  <!-- Running count available via BI Publisher -->\n';
      }

      groupXml += '<?end for-each?>\n';

      await this.wordApi.insertText(groupXml);
    } catch (error) {
      console.error('Error inserting group:', error);
      alert('Error inserting group: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default RepeatingGroup;
export { RepeatingGroup };
