/**
 * HelpPanel Component
 * Comprehensive help system with searchable tag reference, function reference,
 * template patterns, keyboard shortcuts, quick start guide, and version info.
 */

class HelpPanel {
  constructor(services) {
    this.services = services;
    this.activeSection = 'quickstart';
    this.searchQuery = '';
    this.version = '1.0.0';
    this.tagReference = [
      { tag: '<?for-each:path?>', syntax: '<?for-each:path?> ... <?end for-each?>', category: 'loops', description: 'Repeating group for iterating over data elements' },
      { tag: '<?if:condition?>', syntax: '<?if:condition?> ... <?end if?>', category: 'conditional', description: 'Conditional block for showing content based on conditions' },
      { tag: '<?choose?>', syntax: '<?choose?> <?when:cond?> ... <?otherwise?> ... <?end choose?>', category: 'conditional', description: 'Switch-like conditional with multiple cases' },
      { tag: '<?end for-each?>', syntax: 'End tag for for-each', category: 'loops', description: 'Closes a repeating group block' },
      { tag: '<?end if?>', syntax: 'End tag for if block', category: 'conditional', description: 'Closes a conditional block' },
      { tag: '<?end choose?>', syntax: 'End tag for choose', category: 'conditional', description: 'Closes a choose block' },
      { tag: '<?else?>', syntax: 'Else clause in if block', category: 'conditional', description: 'Alternative block executed when condition is false' },
      { tag: '<?when:condition?>', syntax: 'Case in choose block', category: 'conditional', description: 'A condition case in choose block' },
      { tag: '<?otherwise?>', syntax: 'Default case in choose', category: 'conditional', description: 'Default block executed when no conditions match' }
    ];
    this.functionReference = [
      { name: 'count()', returns: 'number', description: 'Count number of items in node set' },
      { name: 'sum()', returns: 'number', description: 'Sum numeric values' },
      { name: 'avg()', returns: 'number', description: 'Calculate average of numeric values' },
      { name: 'max()', returns: 'number', description: 'Find maximum value' },
      { name: 'min()', returns: 'number', description: 'Find minimum value' },
      { name: 'substring()', returns: 'text', description: 'Extract substring from text' },
      { name: 'concat()', returns: 'text', description: 'Concatenate multiple text strings' },
      { name: 'contains()', returns: 'boolean', description: 'Check if text contains substring' },
      { name: 'starts-with()', returns: 'boolean', description: 'Check if text starts with substring' },
      { name: 'ends-with()', returns: 'boolean', description: 'Check if text ends with substring' },
      { name: 'upper-case()', returns: 'text', description: 'Convert text to uppercase' },
      { name: 'lower-case()', returns: 'text', description: 'Convert text to lowercase' }
    ];
    this.patterns = [
      { name: 'Simple Repeating Table', code: '<?for-each:employees?><tr><td><?name?></td><td><?salary?></td></tr><?end for-each?>' },
      { name: 'Conditional Field', code: '<?if:status="Active"?><?employee_name?><?end if?>' },
      { name: 'Grouped Data', code: '<?for-each:departments?><h2><?dept_name?></h2><?for-each:employees?><?name?><?end for-each?><?end for-each?>' }
    ];
    this.shortcuts = [
      { key: 'Ctrl+Shift+F', action: 'Insert field' },
      { key: 'Ctrl+Shift+T', action: 'Insert table' },
      { key: 'Ctrl+Shift+C', action: 'Insert chart' },
      { key: 'Ctrl+Shift+I', action: 'Insert image' },
      { key: 'Ctrl+/', action: 'Toggle help panel' }
    ];
  }

  /**
   * Render the HelpPanel component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'help-panel-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Help & Documentation';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Search bar
    const searchContainer = document.createElement('div');
    searchContainer.style.cssText = 'display: flex; gap: 8px;';
    
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search documentation...';
    searchInput.style.cssText = `
      flex: 1;
      padding: 8px 12px;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      font-size: 12px;
    `;
    searchInput.addEventListener('input', (e) => {
      this.searchQuery = e.target.value;
      this.render(container);
    });
    searchContainer.appendChild(searchInput);

    wrapper.appendChild(searchContainer);

    // Section tabs
    const tabBar = document.createElement('div');
    tabBar.style.cssText = `
      display: flex;
      gap: 4px;
      border-bottom: 1px solid #e5e7eb;
      overflow-x: auto;
    `;

    const sections = [
      { id: 'quickstart', label: 'Quick Start' },
      { id: 'tags', label: 'Tag Reference' },
      { id: 'functions', label: 'Functions' },
      { id: 'patterns', label: 'Patterns' },
      { id: 'shortcuts', label: 'Shortcuts' },
      { id: 'about', label: 'About' }
    ];

    sections.forEach(section => {
      const tabBtn = document.createElement('button');
      tabBtn.textContent = section.label;
      tabBtn.style.cssText = `
        padding: 8px 12px;
        background: transparent;
        border: none;
        border-bottom: 2px solid ${this.activeSection === section.id ? '#3b82f6' : 'transparent'};
        color: ${this.activeSection === section.id ? '#3b82f6' : '#6b7280'};
        font-size: 11px;
        font-weight: ${this.activeSection === section.id ? '600' : '500'};
        cursor: pointer;
        white-space: nowrap;
      `;
      tabBtn.addEventListener('click', () => {
        this.activeSection = section.id;
        this.render(container);
      });
      tabBar.appendChild(tabBtn);
    });

    wrapper.appendChild(tabBar);

    // Content area
    const contentArea = document.createElement('div');
    contentArea.style.cssText = `
      flex: 1;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 12px;
      min-height: 0;
    `;

    switch (this.activeSection) {
      case 'quickstart':
        this.renderQuickStart(contentArea);
        break;
      case 'tags':
        this.renderTagReference(contentArea);
        break;
      case 'functions':
        this.renderFunctionReference(contentArea);
        break;
      case 'patterns':
        this.renderPatterns(contentArea);
        break;
      case 'shortcuts':
        this.renderShortcuts(contentArea);
        break;
      case 'about':
        this.renderAbout(contentArea);
        break;
    }

    wrapper.appendChild(contentArea);

    // Close button
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
    wrapper.appendChild(closeBtn);

    container.appendChild(wrapper);
  }

  /**
   * Render quick start guide
   */
  renderQuickStart(container) {
    const steps = [
      { title: 'Step 1: Insert Fields', desc: 'Use Insert Field to add data fields from your XML data source' },
      { title: 'Step 2: Create Tables', desc: 'Use Table Wizard to create data tables with sorting and grouping' },
      { title: 'Step 3: Add Visuals', desc: 'Insert charts, barcodes, and images using the respective wizards' },
      { title: 'Step 4: Set Conditions', desc: 'Add conditional regions to show/hide content based on data' },
      { title: 'Step 5: Format Report', desc: 'Use Format Helper to apply number, date, and conditional formatting' },
      { title: 'Step 6: Preview & Export', desc: 'Preview your report in different formats (PDF, Excel, HTML, etc.)' }
    ];

    steps.forEach((step, index) => {
      const stepDiv = document.createElement('div');
      stepDiv.style.cssText = `
        padding: 12px;
        border-left: 3px solid #3b82f6;
        background: #f0f9ff;
        border-radius: 4px;
      `;

      const stepTitle = document.createElement('div');
      stepTitle.textContent = step.title;
      stepTitle.style.cssText = 'font-weight: 600; color: #1f2937; font-size: 12px; margin-bottom: 4px;';
      stepDiv.appendChild(stepTitle);

      const stepDesc = document.createElement('div');
      stepDesc.textContent = step.desc;
      stepDesc.style.cssText = 'color: #6b7280; font-size: 11px;';
      stepDiv.appendChild(stepDesc);

      container.appendChild(stepDiv);
    });
  }

  /**
   * Render tag reference
   */
  renderTagReference(container) {
    let filtered = this.tagReference;
    if (this.searchQuery) {
      filtered = this.tagReference.filter(tag =>
        tag.tag.toLowerCase().includes(this.searchQuery.toLowerCase()) ||
        tag.description.toLowerCase().includes(this.searchQuery.toLowerCase())
      );
    }

    if (filtered.length === 0) {
      const noResults = document.createElement('div');
      noResults.style.cssText = 'color: #9ca3af; font-size: 12px; text-align: center; padding: 20px;';
      noResults.textContent = 'No tags found matching your search.';
      container.appendChild(noResults);
      return;
    }

    // Group by category
    const byCategory = {};
    filtered.forEach(tag => {
      if (!byCategory[tag.category]) byCategory[tag.category] = [];
      byCategory[tag.category].push(tag);
    });

    Object.entries(byCategory).forEach(([category, tags]) => {
      const categoryTitle = document.createElement('div');
      categoryTitle.style.cssText = 'font-weight: 600; color: #1f2937; font-size: 12px; margin-top: 8px; text-transform: capitalize;';
      categoryTitle.textContent = category;
      container.appendChild(categoryTitle);

      tags.forEach(tag => {
        const tagDiv = document.createElement('div');
        tagDiv.style.cssText = `
          padding: 10px;
          border: 1px solid #d1d5db;
          border-radius: 4px;
          margin-bottom: 6px;
          font-size: 11px;
        `;

        const tagName = document.createElement('div');
        tagName.textContent = tag.tag;
        tagName.style.cssText = 'font-family: monospace; background: #f3f4f6; padding: 4px; border-radius: 2px; margin-bottom: 4px; font-weight: 600;';
        tagDiv.appendChild(tagName);

        const syntax = document.createElement('div');
        syntax.textContent = 'Syntax: ' + tag.syntax;
        syntax.style.cssText = 'color: #6b7280; margin-bottom: 4px;';
        tagDiv.appendChild(syntax);

        const desc = document.createElement('div');
        desc.textContent = tag.description;
        desc.style.cssText = 'color: #374151;';
        tagDiv.appendChild(desc);

        container.appendChild(tagDiv);
      });
    });
  }

  /**
   * Render function reference
   */
  renderFunctionReference(container) {
    let filtered = this.functionReference;
    if (this.searchQuery) {
      filtered = this.functionReference.filter(fn =>
        fn.name.toLowerCase().includes(this.searchQuery.toLowerCase()) ||
        fn.description.toLowerCase().includes(this.searchQuery.toLowerCase())
      );
    }

    filtered.forEach(fn => {
      const fnDiv = document.createElement('div');
      fnDiv.style.cssText = `
        padding: 10px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        margin-bottom: 6px;
        font-size: 11px;
      `;

      const fnName = document.createElement('div');
      fnName.textContent = fn.name;
      fnName.style.cssText = 'font-family: monospace; font-weight: 600; background: #f3f4f6; padding: 4px; border-radius: 2px; margin-bottom: 4px;';
      fnDiv.appendChild(fnName);

      const returns = document.createElement('div');
      returns.textContent = 'Returns: ' + fn.returns;
      returns.style.cssText = 'color: #6b7280; margin-bottom: 4px;';
      fnDiv.appendChild(returns);

      const desc = document.createElement('div');
      desc.textContent = fn.description;
      desc.style.cssText = 'color: #374151;';
      fnDiv.appendChild(desc);

      container.appendChild(fnDiv);
    });
  }

  /**
   * Render pattern examples
   */
  renderPatterns(container) {
    this.patterns.forEach(pattern => {
      const patternDiv = document.createElement('div');
      patternDiv.style.cssText = `
        padding: 12px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        margin-bottom: 8px;
      `;

      const patternTitle = document.createElement('div');
      patternTitle.textContent = pattern.name;
      patternTitle.style.cssText = 'font-weight: 600; color: #1f2937; font-size: 12px; margin-bottom: 6px;';
      patternDiv.appendChild(patternTitle);

      const patternCode = document.createElement('div');
      patternCode.textContent = pattern.code;
      patternCode.style.cssText = `
        font-family: monospace;
        font-size: 10px;
        background: #f3f4f6;
        padding: 8px;
        border-radius: 3px;
        word-break: break-all;
        color: #374151;
      `;
      patternDiv.appendChild(patternCode);

      const copyBtn = document.createElement('button');
      copyBtn.textContent = 'Copy';
      copyBtn.style.cssText = `
        margin-top: 6px;
        padding: 4px 8px;
        background: #3b82f6;
        color: white;
        border: none;
        border-radius: 3px;
        font-size: 10px;
        cursor: pointer;
      `;
      copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(pattern.code);
        copyBtn.textContent = 'Copied!';
        setTimeout(() => copyBtn.textContent = 'Copy', 2000);
      });
      patternDiv.appendChild(copyBtn);

      container.appendChild(patternDiv);
    });
  }

  /**
   * Render shortcuts
   */
  renderShortcuts(container) {
    const shortcutsTable = document.createElement('table');
    shortcutsTable.style.cssText = 'width: 100%; border-collapse: collapse; font-size: 11px;';

    const header = document.createElement('tr');
    header.style.cssText = 'background: #f3f4f6; border-bottom: 2px solid #d1d5db;';
    header.innerHTML = '<th style="padding: 8px; text-align: left; font-weight: 600;">Shortcut</th><th style="padding: 8px; text-align: left; font-weight: 600;">Action</th>';
    shortcutsTable.appendChild(header);

    this.shortcuts.forEach((shortcut, index) => {
      const row = document.createElement('tr');
      row.style.cssText = `border-bottom: 1px solid #e5e7eb; ${index % 2 === 0 ? 'background: #f9fafb;' : ''}`;
      row.innerHTML = `
        <td style="padding: 8px; font-family: monospace; color: #374151;">${shortcut.key}</td>
        <td style="padding: 8px; color: #6b7280;">${shortcut.action}</td>
      `;
      shortcutsTable.appendChild(row);
    });

    container.appendChild(shortcutsTable);
  }

  /**
   * Render about section
   */
  renderAbout(container) {
    const aboutDiv = document.createElement('div');
    aboutDiv.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    const versionDiv = document.createElement('div');
    versionDiv.style.cssText = 'padding: 12px; background: #f0f9ff; border-radius: 4px;';
    versionDiv.innerHTML = `
      <div style="font-weight: 600; color: #1f2937; margin-bottom: 4px;">BI Publisher Template Builder Add-in</div>
      <div style="color: #6b7280; font-size: 11px;">Version ${this.version}</div>
    `;
    aboutDiv.appendChild(versionDiv);

    const descDiv = document.createElement('div');
    descDiv.style.cssText = 'color: #374151; font-size: 12px; line-height: 1.6;';
    descDiv.innerHTML = `
      <p style="margin: 0 0 8px 0;">A comprehensive Word Add-in for creating and managing Oracle BI Publisher templates with advanced features including:</p>
      <ul style="margin: 0; padding-left: 20px; font-size: 11px;">
        <li>Data field insertion and management</li>
        <li>Interactive table and chart wizards</li>
        <li>Cross-tabulation (pivot) support</li>
        <li>Conditional regions and repeating groups</li>
        <li>Barcode and QR code generation</li>
        <li>Multi-language support with XLIFF</li>
        <li>Accessibility compliance checking</li>
        <li>Preview and export in multiple formats</li>
      </ul>
    `;
    aboutDiv.appendChild(descDiv);

    const licenseDiv = document.createElement('div');
    licenseDiv.style.cssText = 'padding: 8px; background: #f3f4f6; border-radius: 4px; color: #6b7280; font-size: 10px;';
    licenseDiv.textContent = 'For support and documentation, visit the project repository.';
    aboutDiv.appendChild(licenseDiv);

    container.appendChild(aboutDiv);
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default HelpPanel;
export { HelpPanel };
