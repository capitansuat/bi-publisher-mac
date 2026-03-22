/**
 * CrossTabWizard Component
 * 3-step wizard for creating cross-tabulation (pivot table) layouts
 * with row fields, column fields, and measure fields with aggregation.
 */

class CrossTabWizard {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.currentStep = 1;
    this.dataFields = [];
    this.rowFields = [];
    this.columnFields = [];
    this.measureFields = [];
    this.totalsConfig = {
      showRowTotals: true,
      showColumnTotals: true,
      showGrandTotal: true
    };
  }

  /**
   * Render the CrossTabWizard component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'crosstab-wizard-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 16px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    // Header
    const header = document.createElement('div');
    header.style.cssText = 'display: flex; align-items: center; justify-content: space-between;';
    
    const title = document.createElement('h3');
    title.textContent = 'Cross-Tab Wizard';
    title.style.cssText = 'margin: 0; font-size: 16px; font-weight: 700; color: #1f2937;';
    header.appendChild(title);

    const stepIndicator = document.createElement('div');
    stepIndicator.style.cssText = 'font-size: 12px; color: #6b7280;';
    stepIndicator.textContent = `Step ${this.currentStep} of 3`;
    header.appendChild(stepIndicator);

    wrapper.appendChild(header);

    // Progress bar
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
      width: ${(this.currentStep / 3) * 100}%;
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
    }

    wrapper.appendChild(contentArea);

    // Buttons
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
    nextBtn.textContent = this.currentStep === 3 ? 'Insert Cross-Tab' : 'Next';
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
      if (this.currentStep === 3) {
        this.insertCrossTab(container);
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
   * Step 1: Select row fields
   */
  renderStep1(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 16px; min-height: 0;';

    // Available fields
    const availSection = document.createElement('div');
    availSection.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
    
    const availLabel = document.createElement('label');
    availLabel.textContent = 'Available Fields:';
    availLabel.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
    availSection.appendChild(availLabel);

    this.loadDataFields();

    const availList = document.createElement('div');
    availList.className = 'available-fields';
    availList.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      overflow-y: auto;
      background: #f9fafb;
      flex: 1;
      min-height: 0;
    `;

    this.dataFields.forEach(field => {
      if (!this.rowFields.includes(field.name)) {
        const item = document.createElement('div');
        item.className = 'field-item';
        item.textContent = field.name;
        item.draggable = true;
        item.style.cssText = `
          padding: 8px 12px;
          border-bottom: 1px solid #e5e7eb;
          cursor: move;
          user-select: none;
          transition: background 0.15s;
        `;
        item.onmouseover = () => item.style.background = '#f0fdf4';
        item.onmouseout = () => item.style.background = 'transparent';
        
        item.addEventListener('dragstart', (e) => {
          e.dataTransfer.setData('field', field.name);
          e.dataTransfer.setData('source', 'available');
        });

        availList.appendChild(item);
      }
    });

    availSection.appendChild(availList);
    section.appendChild(availSection);

    // Selected row fields
    const rowSection = document.createElement('div');
    rowSection.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
    
    const rowLabel = document.createElement('label');
    rowLabel.textContent = 'Row Fields:';
    rowLabel.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
    rowSection.appendChild(rowLabel);

    const rowList = document.createElement('div');
    rowList.className = 'row-fields-drop';
    rowList.style.cssText = `
      border: 2px dashed #d1d5db;
      border-radius: 4px;
      padding: 12px;
      min-height: 100px;
      background: #fafafa;
      flex: 1;
      overflow-y: auto;
      transition: border-color 0.2s;
    `;

    this.rowFields.forEach((field, index) => {
      const item = document.createElement('div');
      item.style.cssText = `
        padding: 8px 12px;
        background: #dbeafe;
        border-radius: 4px;
        margin-bottom: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 12px;
        color: #1e40af;
      `;
      
      const label = document.createElement('span');
      label.textContent = field;
      item.appendChild(label);

      const removeBtn = document.createElement('button');
      removeBtn.textContent = '✕';
      removeBtn.style.cssText = `
        background: transparent;
        border: none;
        color: #1e40af;
        cursor: pointer;
        font-weight: bold;
        padding: 0 4px;
      `;
      removeBtn.addEventListener('click', () => {
        this.rowFields = this.rowFields.filter(f => f !== field);
        this.render(container);
      });
      item.appendChild(removeBtn);

      rowList.appendChild(item);
    });

    // Drag and drop handlers
    rowList.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      rowList.style.borderColor = '#3b82f6';
      rowList.style.backgroundColor = '#eff6ff';
    });

    rowList.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      rowList.style.borderColor = '#d1d5db';
      rowList.style.backgroundColor = '#fafafa';
    });

    rowList.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const field = e.dataTransfer.getData('field');
      if (field && !this.rowFields.includes(field)) {
        this.rowFields.push(field);
        this.render(container);
      }
      rowList.style.borderColor = '#d1d5db';
      rowList.style.backgroundColor = '#fafafa';
    });

    rowSection.appendChild(rowList);
    section.appendChild(rowSection);

    container.appendChild(section);
  }

  /**
   * Step 2: Select column fields
   */
  renderStep2(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 16px; min-height: 0;';

    // Available fields
    const availSection = document.createElement('div');
    availSection.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
    
    const availLabel = document.createElement('label');
    availLabel.textContent = 'Available Fields:';
    availLabel.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
    availSection.appendChild(availLabel);

    const availList = document.createElement('div');
    availList.className = 'available-fields';
    availList.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      overflow-y: auto;
      background: #f9fafb;
      flex: 1;
      min-height: 0;
    `;

    this.dataFields.forEach(field => {
      if (!this.columnFields.includes(field.name)) {
        const item = document.createElement('div');
        item.className = 'field-item';
        item.textContent = field.name;
        item.draggable = true;
        item.style.cssText = `
          padding: 8px 12px;
          border-bottom: 1px solid #e5e7eb;
          cursor: move;
          user-select: none;
          transition: background 0.15s;
        `;
        item.onmouseover = () => item.style.background = '#fef3c7';
        item.onmouseout = () => item.style.background = 'transparent';
        
        item.addEventListener('dragstart', (e) => {
          e.dataTransfer.setData('field', field.name);
          e.dataTransfer.setData('source', 'available');
        });

        availList.appendChild(item);
      }
    });

    availSection.appendChild(availList);
    section.appendChild(availSection);

    // Selected column fields
    const colSection = document.createElement('div');
    colSection.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
    
    const colLabel = document.createElement('label');
    colLabel.textContent = 'Column Fields:';
    colLabel.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
    colSection.appendChild(colLabel);

    const colList = document.createElement('div');
    colList.className = 'column-fields-drop';
    colList.style.cssText = `
      border: 2px dashed #d1d5db;
      border-radius: 4px;
      padding: 12px;
      min-height: 100px;
      background: #fafafa;
      flex: 1;
      overflow-y: auto;
      transition: border-color 0.2s;
    `;

    this.columnFields.forEach((field, index) => {
      const item = document.createElement('div');
      item.style.cssText = `
        padding: 8px 12px;
        background: #fef3c7;
        border-radius: 4px;
        margin-bottom: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 12px;
        color: #92400e;
      `;
      
      const label = document.createElement('span');
      label.textContent = field;
      item.appendChild(label);

      const removeBtn = document.createElement('button');
      removeBtn.textContent = '✕';
      removeBtn.style.cssText = `
        background: transparent;
        border: none;
        color: #92400e;
        cursor: pointer;
        font-weight: bold;
        padding: 0 4px;
      `;
      removeBtn.addEventListener('click', () => {
        this.columnFields = this.columnFields.filter(f => f !== field);
        this.render(container);
      });
      item.appendChild(removeBtn);

      colList.appendChild(item);
    });

    // Drag and drop
    colList.addEventListener('dragover', (e) => {
      e.preventDefault();
      colList.style.borderColor = '#f59e0b';
      colList.style.backgroundColor = '#fffbeb';
    });

    colList.addEventListener('dragleave', (e) => {
      e.preventDefault();
      colList.style.borderColor = '#d1d5db';
      colList.style.backgroundColor = '#fafafa';
    });

    colList.addEventListener('drop', (e) => {
      e.preventDefault();
      const field = e.dataTransfer.getData('field');
      if (field && !this.columnFields.includes(field)) {
        this.columnFields.push(field);
        this.render(container);
      }
      colList.style.borderColor = '#d1d5db';
      colList.style.backgroundColor = '#fafafa';
    });

    colSection.appendChild(colList);
    section.appendChild(colSection);

    container.appendChild(section);
  }

  /**
   * Step 3: Select measure fields with aggregation
   */
  renderStep3(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 16px;';

    // Measure fields configuration
    const measSection = document.createElement('fieldset');
    measSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const measLegend = document.createElement('legend');
    measLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    measLegend.textContent = 'Measure Fields & Aggregation';
    measSection.appendChild(measLegend);

    const measContent = document.createElement('div');
    measContent.style.cssText = 'display: flex; flex-direction: column; gap: 10px;';

    this.dataFields.forEach(field => {
      if (field.type === 'number') {
        const row = document.createElement('div');
        row.style.cssText = 'display: grid; grid-template-columns: 1fr auto; gap: 8px; align-items: center;';

        const checkbox = document.createElement('label');
        checkbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.value = field.name;
        input.checked = this.measureFields.some(m => m.name === field.name);
        input.addEventListener('change', (e) => {
          if (e.target.checked) {
            if (!this.measureFields.find(m => m.name === field.name)) {
              this.measureFields.push({ name: field.name, aggregation: 'sum' });
            }
          } else {
            this.measureFields = this.measureFields.filter(m => m.name !== field.name);
          }
          this.render(container);
        });
        checkbox.appendChild(input);
        checkbox.appendChild(document.createTextNode(field.name));
        row.appendChild(checkbox);

        if (input.checked) {
          const aggSelect = document.createElement('select');
          aggSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
          aggSelect.innerHTML = `
            <option value="sum">Sum</option>
            <option value="count">Count</option>
            <option value="avg">Average</option>
            <option value="min">Minimum</option>
            <option value="max">Maximum</option>
            <option value="stddev">Std Dev</option>
          `;
          const measure = this.measureFields.find(m => m.name === field.name);
          if (measure) aggSelect.value = measure.aggregation;
          aggSelect.addEventListener('change', (e) => {
            const m = this.measureFields.find(mf => mf.name === field.name);
            if (m) m.aggregation = e.target.value;
          });
          row.appendChild(aggSelect);
        }

        measContent.appendChild(row);
      }
    });

    measSection.appendChild(measContent);
    section.appendChild(measSection);

    // Totals options
    const totalsSection = document.createElement('fieldset');
    totalsSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const totalsLegend = document.createElement('legend');
    totalsLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    totalsLegend.textContent = 'Totals';
    totalsSection.appendChild(totalsLegend);

    const totalsContent = document.createElement('div');
    totalsContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const rowTotalsCheckbox = document.createElement('label');
    rowTotalsCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const rowTotalsInput = document.createElement('input');
    rowTotalsInput.type = 'checkbox';
    rowTotalsInput.checked = this.totalsConfig.showRowTotals;
    rowTotalsInput.addEventListener('change', (e) => this.totalsConfig.showRowTotals = e.target.checked);
    rowTotalsCheckbox.appendChild(rowTotalsInput);
    rowTotalsCheckbox.appendChild(document.createTextNode('Show row totals'));
    totalsContent.appendChild(rowTotalsCheckbox);

    const colTotalsCheckbox = document.createElement('label');
    colTotalsCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const colTotalsInput = document.createElement('input');
    colTotalsInput.type = 'checkbox';
    colTotalsInput.checked = this.totalsConfig.showColumnTotals;
    colTotalsInput.addEventListener('change', (e) => this.totalsConfig.showColumnTotals = e.target.checked);
    colTotalsCheckbox.appendChild(colTotalsInput);
    colTotalsCheckbox.appendChild(document.createTextNode('Show column totals'));
    totalsContent.appendChild(colTotalsCheckbox);

    const grandTotalCheckbox = document.createElement('label');
    grandTotalCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const grandTotalInput = document.createElement('input');
    grandTotalInput.type = 'checkbox';
    grandTotalInput.checked = this.totalsConfig.showGrandTotal;
    grandTotalInput.addEventListener('change', (e) => this.totalsConfig.showGrandTotal = e.target.checked);
    grandTotalCheckbox.appendChild(grandTotalInput);
    grandTotalCheckbox.appendChild(document.createTextNode('Show grand total'));
    totalsContent.appendChild(grandTotalCheckbox);

    totalsSection.appendChild(totalsContent);
    section.appendChild(totalsSection);

    // Matrix preview
    const previewSection = document.createElement('fieldset');
    previewSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const previewLegend = document.createElement('legend');
    previewLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    previewLegend.textContent = 'Preview';
    previewSection.appendChild(previewLegend);

    const preview = document.createElement('div');
    preview.style.cssText = `
      font-size: 11px;
      color: #6b7280;
      line-height: 1.5;
    `;
    preview.innerHTML = `
      <div><strong>Rows:</strong> ${this.rowFields.length > 0 ? this.rowFields.join(', ') : 'None selected'}</div>
      <div style="margin-top: 4px;"><strong>Columns:</strong> ${this.columnFields.length > 0 ? this.columnFields.join(', ') : 'None selected'}</div>
      <div style="margin-top: 4px;"><strong>Measures:</strong> ${this.measureFields.length > 0 ? this.measureFields.map(m => `${m.name} (${m.aggregation})`).join(', ') : 'None selected'}</div>
    `;
    previewSection.appendChild(preview);

    section.appendChild(previewSection);
    container.appendChild(section);
  }

  /**
   * Load data fields
   */
  loadDataFields() {
    this.dataFields = [
      { name: 'Region', type: 'text' },
      { name: 'Product', type: 'text' },
      { name: 'Year', type: 'text' },
      { name: 'Quarter', type: 'text' },
      { name: 'Sales', type: 'number' },
      { name: 'Quantity', type: 'number' },
      { name: 'Profit', type: 'number' }
    ];
  }

  /**
   * Move to next step
   */
  nextStep(container) {
    if (this.currentStep < 3) {
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
   * Insert cross-tab into document
   */
  async insertCrossTab(container) {
    try {
      if (this.rowFields.length === 0 || this.columnFields.length === 0 || this.measureFields.length === 0) {
        alert('Please select row fields, column fields, and measure fields');
        return;
      }

      // Generate BI Publisher cross-tab XML
      let crossTabXml = '<bi:crosstab>\n';
      
      crossTabXml += '  <bi:rowFields>\n';
      this.rowFields.forEach(field => {
        crossTabXml += `    <bi:field>${field}</bi:field>\n`;
      });
      crossTabXml += '  </bi:rowFields>\n';

      crossTabXml += '  <bi:columnFields>\n';
      this.columnFields.forEach(field => {
        crossTabXml += `    <bi:field>${field}</bi:field>\n`;
      });
      crossTabXml += '  </bi:columnFields>\n';

      crossTabXml += '  <bi:measures>\n';
      this.measureFields.forEach(measure => {
        crossTabXml += `    <bi:measure name="${measure.name}" aggregation="${measure.aggregation}"/>\n`;
      });
      crossTabXml += '  </bi:measures>\n';

      if (this.totalsConfig.showRowTotals) {
        crossTabXml += '  <bi:rowTotals>true</bi:rowTotals>\n';
      }
      if (this.totalsConfig.showColumnTotals) {
        crossTabXml += '  <bi:columnTotals>true</bi:columnTotals>\n';
      }
      if (this.totalsConfig.showGrandTotal) {
        crossTabXml += '  <bi:grandTotal>true</bi:grandTotal>\n';
      }

      crossTabXml += '</bi:crosstab>\n';

      await this.wordApi.insertText(crossTabXml);
      container.innerHTML = '';
    } catch (error) {
      console.error('Error inserting cross-tab:', error);
      alert('Error inserting cross-tab: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default CrossTabWizard;
export { CrossTabWizard };
