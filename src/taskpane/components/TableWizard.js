/**
 * TableWizard Component
 * 4-step wizard for creating data tables with column selection, sorting,
 * grouping, and styling options for BI Publisher templates.
 */

class TableWizard {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.currentStep = 1;
    this.dataGroups = [];
    this.selectedColumns = [];
    this.sortConfig = {};
    this.groupConfig = {};
    this.totalsConfig = {};
    this.styleConfig = { style: 'grid', alternating: true, borders: true };
    this.previewData = [];
  }

  /**
   * Render the TableWizard component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'table-wizard-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 16px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    // Header with step indicator
    const header = document.createElement('div');
    header.style.cssText = 'display: flex; align-items: center; gap: 12px; margin-bottom: 8px;';
    
    const title = document.createElement('h3');
    title.textContent = 'Table Wizard';
    title.style.cssText = 'margin: 0; font-size: 16px; font-weight: 700; color: #1f2937; flex: 1;';
    header.appendChild(title);

    const stepIndicator = document.createElement('div');
    stepIndicator.className = 'step-indicator';
    stepIndicator.style.cssText = 'font-size: 12px; color: #6b7280; font-weight: 500;';
    stepIndicator.textContent = `Step ${this.currentStep} of 4`;
    header.appendChild(stepIndicator);

    wrapper.appendChild(header);

    // Step progress bar
    const progressBar = document.createElement('div');
    progressBar.style.cssText = `
      height: 4px;
      background: #e5e7eb;
      border-radius: 2px;
      overflow: hidden;
    `;
    const progressFill = document.createElement('div');
    progressFill.style.cssText = `
      height: 100%;
      background: #3b82f6;
      width: ${(this.currentStep / 4) * 100}%;
      transition: width 0.3s;
    `;
    progressBar.appendChild(progressFill);
    wrapper.appendChild(progressBar);

    // Content area
    const contentArea = document.createElement('div');
    contentArea.className = 'wizard-content';
    contentArea.style.cssText = `
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 16px;
      overflow-y: auto;
      min-height: 0;
    `;

    // Render appropriate step
    switch (this.currentStep) {
      case 1:
        this.renderStep1(contentArea);
        break;
      case 2:
        this.renderStep2(contentArea);
        break;
      case 3:
        this.renderStep3(contentArea);
        break;
      case 4:
        this.renderStep4(contentArea);
        break;
    }

    wrapper.appendChild(contentArea);

    // Bottom buttons
    const buttonArea = document.createElement('div');
    buttonArea.style.cssText = `
      display: flex;
      gap: 8px;
      padding-top: 12px;
      border-top: 1px solid #d1d5db;
    `;

    const backBtn = document.createElement('button');
    backBtn.textContent = 'Back';
    backBtn.style.cssText = `
      padding: 8px 16px;
      background: #e5e7eb;
      color: #374151;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
      display: ${this.currentStep === 1 ? 'none' : 'block'};
    `;
    backBtn.addEventListener('click', () => this.previousStep(container));
    buttonArea.appendChild(backBtn);

    const nextBtn = document.createElement('button');
    nextBtn.textContent = this.currentStep === 4 ? 'Insert Table' : 'Next';
    nextBtn.className = 'next-step-btn';
    nextBtn.style.cssText = `
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
    nextBtn.addEventListener('click', () => {
      if (this.currentStep === 4) {
        this.insertTable(container);
      } else {
        this.nextStep(container);
      }
    });
    buttonArea.appendChild(nextBtn);

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = `
      padding: 8px 16px;
      background: #f3f4f6;
      color: #6b7280;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    `;
    cancelBtn.addEventListener('click', () => container.innerHTML = '');
    buttonArea.appendChild(cancelBtn);

    wrapper.appendChild(buttonArea);
    container.appendChild(wrapper);
  }

  /**
   * Step 1: Select repeating data group
   */
  renderStep1(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    const label = document.createElement('label');
    label.textContent = 'Select Repeating Data Group:';
    label.style.cssText = 'font-size: 13px; font-weight: 600; color: #1f2937;';
    section.appendChild(label);

    // Load repeating groups
    this.loadDataGroups();

    const groupList = document.createElement('div');
    groupList.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; overflow: hidden;';

    this.dataGroups.forEach((group, index) => {
      const item = document.createElement('div');
      item.style.cssText = `
        padding: 12px;
        border-bottom: ${index < this.dataGroups.length - 1 ? '1px solid #e5e7eb' : 'none'};
        cursor: pointer;
        transition: background 0.15s;
        display: flex;
        align-items: center;
        gap: 12px;
      `;
      item.onmouseover = () => item.style.background = '#f9fafb';
      item.onmouseout = () => item.style.background = 'transparent';

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'data-group';
      radio.value = group.xpath;
      radio.style.cssText = 'cursor: pointer;';
      radio.addEventListener('change', () => {
        this.dataGroups.forEach(g => g.selected = false);
        group.selected = true;
      });
      item.appendChild(radio);

      const info = document.createElement('div');
      info.style.cssText = 'flex: 1;';
      info.innerHTML = `
        <div style="font-size: 13px; font-weight: 500; color: #1f2937;">${group.name}</div>
        <div style="font-size: 11px; color: #6b7280; margin-top: 2px;">Path: ${group.xpath}</div>
      `;
      item.appendChild(info);

      groupList.appendChild(item);
    });

    section.appendChild(groupList);

    // Description
    const desc = document.createElement('div');
    desc.style.cssText = `
      padding: 12px;
      background: #f0fdf4;
      border: 1px solid #dcfce7;
      border-radius: 4px;
      font-size: 12px;
      color: #166534;
    `;
    desc.textContent = 'Select the repeating group that will provide rows for your table. This group must contain the data records you want to display.';
    section.appendChild(desc);

    container.appendChild(section);
  }

  /**
   * Step 2: Select and reorder columns
   */
  renderStep2(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    const label = document.createElement('label');
    label.textContent = 'Select Columns (drag to reorder):';
    label.style.cssText = 'font-size: 13px; font-weight: 600; color: #1f2937;';
    section.appendChild(label);

    const columnList = document.createElement('div');
    columnList.className = 'column-list';
    columnList.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      max-height: 300px;
      overflow-y: auto;
      background: white;
    `;

    const selectedGroup = this.dataGroups.find(g => g.selected);
    const fields = selectedGroup ? selectedGroup.fields : [];

    fields.forEach((field, index) => {
      const item = document.createElement('div');
      item.className = 'column-item';
      item.draggable = true;
      item.style.cssText = `
        padding: 10px 12px;
        border-bottom: ${index < fields.length - 1 ? '1px solid #e5e7eb' : 'none'};
        display: flex;
        align-items: center;
        gap: 10px;
        cursor: move;
        transition: background 0.15s;
      `;
      item.onmouseover = () => item.style.background = '#f9fafb';
      item.onmouseout = () => item.style.background = 'transparent';

      const dragHandle = document.createElement('span');
      dragHandle.textContent = '≡';
      dragHandle.style.cssText = 'color: #9ca3af; font-size: 14px; cursor: grab;';
      item.appendChild(dragHandle);

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'column-checkbox';
      checkbox.value = field.name;
      checkbox.checked = true;
      checkbox.style.cssText = 'cursor: pointer;';
      item.appendChild(checkbox);

      const fieldName = document.createElement('span');
      fieldName.textContent = field.name;
      fieldName.style.cssText = 'flex: 1; font-size: 13px; color: #1f2937;';
      item.appendChild(fieldName);

      const fieldType = document.createElement('span');
      fieldType.textContent = field.type;
      fieldType.style.cssText = `
        font-size: 10px;
        padding: 2px 6px;
        background: #dbeafe;
        color: #1e40af;
        border-radius: 3px;
      `;
      item.appendChild(fieldType);

      columnList.appendChild(item);
    });

    section.appendChild(columnList);

    // Help text
    const help = document.createElement('div');
    help.style.cssText = `
      padding: 10px;
      background: #eff6ff;
      border: 1px solid #bfdbfe;
      border-radius: 4px;
      font-size: 12px;
      color: #1e40af;
    `;
    help.textContent = 'Check the columns you want to include. Drag to reorder them.';
    section.appendChild(help);

    container.appendChild(section);
  }

  /**
   * Step 3: Configure sorting, grouping, and totals
   */
  renderStep3(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 16px;';

    // Sorting section
    const sortingSection = document.createElement('fieldset');
    sortingSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const sortingLegend = document.createElement('legend');
    sortingLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    sortingLegend.textContent = 'Sort Options';
    sortingSection.appendChild(sortingLegend);

    const sortingContent = document.createElement('div');
    sortingContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px; font-size: 12px;';
    
    const sortFieldSelect = document.createElement('select');
    sortFieldSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    sortFieldSelect.innerHTML = '<option>-- No Sort --</option>';
    const selectedGroup = this.dataGroups.find(g => g.selected);
    if (selectedGroup) {
      selectedGroup.fields.forEach(f => {
        sortFieldSelect.innerHTML += `<option value="${f.name}">${f.name}</option>`;
      });
    }
    
    const sortLabel = document.createElement('label');
    sortLabel.style.cssText = 'display: flex; align-items: center; gap: 8px;';
    sortLabel.appendChild(document.createTextNode('Sort by: '));
    sortLabel.appendChild(sortFieldSelect);
    sortingContent.appendChild(sortLabel);

    const sortOrder = document.createElement('label');
    sortOrder.style.cssText = 'display: flex; align-items: center; gap: 8px;';
    const sortRadio = document.createElement('input');
    sortRadio.type = 'radio';
    sortRadio.name = 'sort-order';
    sortRadio.checked = true;
    sortOrder.appendChild(sortRadio);
    sortOrder.appendChild(document.createTextNode('Ascending'));
    sortingContent.appendChild(sortOrder);

    const sortOrderDesc = document.createElement('label');
    sortOrderDesc.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-left: 24px;';
    const sortRadioDesc = document.createElement('input');
    sortRadioDesc.type = 'radio';
    sortRadioDesc.name = 'sort-order';
    sortOrderDesc.appendChild(sortRadioDesc);
    sortOrderDesc.appendChild(document.createTextNode('Descending'));
    sortingContent.appendChild(sortOrderDesc);

    sortingSection.appendChild(sortingContent);
    section.appendChild(sortingSection);

    // Grouping section
    const groupingSection = document.createElement('fieldset');
    groupingSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const groupingLegend = document.createElement('legend');
    groupingLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    groupingLegend.textContent = 'Group Options';
    groupingSection.appendChild(groupingLegend);

    const groupingContent = document.createElement('div');
    groupingContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px; font-size: 12px;';
    
    const groupFieldSelect = document.createElement('select');
    groupFieldSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    groupFieldSelect.innerHTML = '<option>-- No Grouping --</option>';
    if (selectedGroup) {
      selectedGroup.fields.forEach(f => {
        groupFieldSelect.innerHTML += `<option value="${f.name}">${f.name}</option>`;
      });
    }
    
    const groupLabel = document.createElement('label');
    groupLabel.style.cssText = 'display: flex; align-items: center; gap: 8px;';
    groupLabel.appendChild(document.createTextNode('Group by: '));
    groupLabel.appendChild(groupFieldSelect);
    groupingContent.appendChild(groupLabel);

    const groupHeaderCheckbox = document.createElement('label');
    groupHeaderCheckbox.style.cssText = 'display: flex; align-items: center; gap: 6px; cursor: pointer;';
    const headerInput = document.createElement('input');
    headerInput.type = 'checkbox';
    headerInput.checked = true;
    groupHeaderCheckbox.appendChild(headerInput);
    groupHeaderCheckbox.appendChild(document.createTextNode('Show group header'));
    groupingContent.appendChild(groupHeaderCheckbox);

    const groupFooterCheckbox = document.createElement('label');
    groupFooterCheckbox.style.cssText = 'display: flex; align-items: center; gap: 6px; cursor: pointer;';
    const footerInput = document.createElement('input');
    footerInput.type = 'checkbox';
    groupFooterCheckbox.appendChild(footerInput);
    groupFooterCheckbox.appendChild(document.createTextNode('Show group footer'));
    groupingContent.appendChild(groupFooterCheckbox);

    groupingSection.appendChild(groupingContent);
    section.appendChild(groupingSection);

    // Totals section
    const totalsSection = document.createElement('fieldset');
    totalsSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const totalsLegend = document.createElement('legend');
    totalsLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    totalsLegend.textContent = 'Totals Per Column';
    totalsSection.appendChild(totalsLegend);

    const totalsContent = document.createElement('div');
    totalsContent.style.cssText = 'display: flex; flex-direction: column; gap: 6px; font-size: 11px;';
    
    if (selectedGroup) {
      selectedGroup.fields.forEach(field => {
        const rowDiv = document.createElement('div');
        rowDiv.style.cssText = 'display: flex; align-items: center; gap: 8px;';
        
        const fieldLabel = document.createElement('span');
        fieldLabel.textContent = field.name;
        fieldLabel.style.cssText = 'min-width: 100px;';
        rowDiv.appendChild(fieldLabel);

        const totalSelect = document.createElement('select');
        totalSelect.style.cssText = 'flex: 1; padding: 4px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
        totalSelect.innerHTML = `
          <option value="">None</option>
          <option value="sum">Sum</option>
          <option value="count">Count</option>
          <option value="avg">Average</option>
          <option value="min">Minimum</option>
          <option value="max">Maximum</option>
        `;
        rowDiv.appendChild(totalSelect);

        totalsContent.appendChild(rowDiv);
      });
    }

    totalsSection.appendChild(totalsContent);
    section.appendChild(totalsSection);

    container.appendChild(section);
  }

  /**
   * Step 4: Configure table style and appearance
   */
  renderStep4(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 16px;';

    // Style selection
    const styleSection = document.createElement('fieldset');
    styleSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const styleLegend = document.createElement('legend');
    styleLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    styleLegend.textContent = 'Table Style';
    styleSection.appendChild(styleLegend);

    const styles = [
      { value: 'grid', label: 'Grid (borders on all sides)' },
      { value: 'list', label: 'List (horizontal lines only)' },
      { value: 'form', label: 'Form (no borders)' }
    ];

    styles.forEach(style => {
      const radio = document.createElement('label');
      radio.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 6px; cursor: pointer; font-size: 12px;';
      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'table-style';
      input.value = style.value;
      input.checked = style.value === this.styleConfig.style;
      input.addEventListener('change', () => this.styleConfig.style = style.value);
      radio.appendChild(input);
      radio.appendChild(document.createTextNode(style.label));
      styleSection.appendChild(radio);
    });

    section.appendChild(styleSection);

    // Options
    const optionsSection = document.createElement('fieldset');
    optionsSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const optionsLegend = document.createElement('legend');
    optionsLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    optionsLegend.textContent = 'Options';
    optionsSection.appendChild(optionsLegend);

    const alternatingCheckbox = document.createElement('label');
    alternatingCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 8px; cursor: pointer; font-size: 12px;';
    const altInput = document.createElement('input');
    altInput.type = 'checkbox';
    altInput.checked = this.styleConfig.alternating;
    altInput.addEventListener('change', () => this.styleConfig.alternating = altInput.checked);
    alternatingCheckbox.appendChild(altInput);
    alternatingCheckbox.appendChild(document.createTextNode('Alternating row colors'));
    optionsSection.appendChild(alternatingCheckbox);

    const bordersCheckbox = document.createElement('label');
    bordersCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const borderInput = document.createElement('input');
    borderInput.type = 'checkbox';
    borderInput.checked = this.styleConfig.borders;
    borderInput.addEventListener('change', () => this.styleConfig.borders = borderInput.checked);
    bordersCheckbox.appendChild(borderInput);
    bordersCheckbox.appendChild(document.createTextNode('Show borders'));
    optionsSection.appendChild(bordersCheckbox);

    section.appendChild(optionsSection);

    // Preview
    const previewSection = document.createElement('fieldset');
    previewSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const previewLegend = document.createElement('legend');
    previewLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    previewLegend.textContent = 'Preview';
    previewSection.appendChild(previewLegend);

    const preview = document.createElement('div');
    preview.style.cssText = `
      max-height: 200px;
      overflow: auto;
      background: white;
      border-radius: 3px;
      padding: 8px;
    `;
    preview.innerHTML = '<p style="color: #9ca3af; font-size: 12px;">Table preview will appear here</p>';
    previewSection.appendChild(preview);

    section.appendChild(previewSection);
    container.appendChild(section);
  }

  /**
   * Load available data groups
   */
  loadDataGroups() {
    this.dataGroups = [
      {
        name: 'Employees',
        xpath: '/Company/Employees/Employee',
        selected: true,
        fields: [
          { name: 'Name', type: 'text' },
          { name: 'Department', type: 'text' },
          { name: 'Salary', type: 'number' },
          { name: 'HireDate', type: 'date' }
        ]
      },
      {
        name: 'Orders',
        xpath: '/Company/Orders/Order',
        selected: false,
        fields: [
          { name: 'OrderID', type: 'text' },
          { name: 'Customer', type: 'text' },
          { name: 'Amount', type: 'number' },
          { name: 'OrderDate', type: 'date' }
        ]
      }
    ];
  }

  /**
   * Move to next step
   */
  nextStep(container) {
    if (this.currentStep < 4) {
      this.currentStep++;
      this.render(container);
    }
  }

  /**
   * Move to previous step
   */
  previousStep(container) {
    if (this.currentStep > 1) {
      this.currentStep--;
      this.render(container);
    }
  }

  /**
   * Insert the configured table into the document
   */
  async insertTable(container) {
    try {
      const selectedGroup = this.dataGroups.find(g => g.selected);
      if (!selectedGroup) {
        alert('Please select a data group');
        return;
      }

      const columns = document.querySelectorAll('.column-checkbox:checked');
      const columnNames = Array.from(columns).map(cb => cb.value);

      // Generate BI Publisher table XML
      let tableXml = '<?for-each:' + selectedGroup.xpath + '?>\n';
      tableXml += '  <w:tbl>\n';
      tableXml += '    <w:tblPr><w:tblW w:w="5000" w:type="auto"/></w:tblPr>\n';
      
      // Header row
      tableXml += '    <w:tr>\n';
      columnNames.forEach(col => {
        tableXml += `      <w:tc><w:p><w:r><w:t>${col}</w:t></w:r></w:p></w:tc>\n`;
      });
      tableXml += '    </w:tr>\n';

      // Data rows
      tableXml += '    <w:tr>\n';
      columnNames.forEach(col => {
        const field = selectedGroup.fields.find(f => f.name === col);
        const fieldPath = selectedGroup.xpath.replace(/\/[^\/]+$/, '') + '/' + col;
        tableXml += `      <w:tc><w:p><w:r><w:t><?${fieldPath}?></w:t></w:r></w:p></w:tc>\n`;
      });
      tableXml += '    </w:tr>\n';

      tableXml += '  </w:tbl>\n';
      tableXml += '<?end for-each?>\n';

      await this.wordApi.insertText(tableXml);
      container.innerHTML = '';
    } catch (error) {
      console.error('Error inserting table:', error);
      alert('Error inserting table: ' + error.message);
    }
  }

  /**
   * Bind component events
   */
  bindEvents() {
    // Events are bound during render
  }
}

export default TableWizard;
export { TableWizard };
