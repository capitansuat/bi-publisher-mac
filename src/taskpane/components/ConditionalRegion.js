/**
 * ConditionalRegion Component
 * Component for creating conditional regions with If, If-Else, and Choose/When/Otherwise
 * structures with condition builders supporting AND/OR logic for BI Publisher templates.
 */

class ConditionalRegion {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.conditionalMode = 'simple'; // 'simple', 'ifelse', 'choose'
    this.conditions = [];
    this.conditionLogic = 'and'; // 'and', 'or'
    this.xslExpression = '';
    this.dataFields = [];
  }

  /**
   * Render the ConditionalRegion component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'conditional-region-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Conditional Region';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Mode selector
    const modeSection = document.createElement('fieldset');
    modeSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const modeLegend = document.createElement('legend');
    modeLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    modeLegend.textContent = 'Conditional Type';
    modeSection.appendChild(modeLegend);

    const modes = [
      { value: 'simple', label: 'Simple If' },
      { value: 'ifelse', label: 'If-Else' },
      { value: 'choose', label: 'Choose/When/Otherwise' }
    ];

    modes.forEach(mode => {
      const radio = document.createElement('label');
      radio.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 6px; cursor: pointer; font-size: 12px;';
      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'conditional-mode';
      input.value = mode.value;
      input.checked = mode.value === this.conditionalMode;
      input.addEventListener('change', () => {
        this.conditionalMode = mode.value;
        this.render(container);
      });
      radio.appendChild(input);
      radio.appendChild(document.createTextNode(mode.label));
      modeSection.appendChild(radio);
    });

    wrapper.appendChild(modeSection);

    // Conditions builder
    const condSection = document.createElement('fieldset');
    condSection.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      padding: 12px;
      margin: 0;
      overflow-y: auto;
      max-height: 300px;
    `;
    
    const condLegend = document.createElement('legend');
    condLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    condLegend.textContent = 'Conditions';
    condSection.appendChild(condLegend);

    const condContent = document.createElement('div');
    condContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    this.loadDataFields();

    // Existing conditions
    if (this.conditions.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'color: #9ca3af; font-size: 12px; text-align: center; padding: 12px;';
      emptyMsg.textContent = 'No conditions added yet';
      condContent.appendChild(emptyMsg);
    } else {
      this.conditions.forEach((cond, index) => {
        const condItem = document.createElement('div');
        condItem.style.cssText = `
          padding: 10px;
          background: #f9fafb;
          border: 1px solid #e5e7eb;
          border-radius: 4px;
          display: grid;
          grid-template-columns: 1fr auto auto;
          gap: 8px;
          align-items: center;
          font-size: 11px;
        `;

        const display = document.createElement('span');
        display.textContent = `${cond.field} ${cond.operator} ${cond.value}`;
        display.style.cssText = 'font-weight: 500;';
        condItem.appendChild(display);

        const removeBtn = document.createElement('button');
        removeBtn.textContent = 'Remove';
        removeBtn.style.cssText = `
          padding: 4px 8px;
          background: #fee2e2;
          color: #991b1b;
          border: none;
          border-radius: 3px;
          font-size: 11px;
          cursor: pointer;
        `;
        removeBtn.addEventListener('click', () => {
          this.conditions.splice(index, 1);
          this.render(container);
        });
        condItem.appendChild(removeBtn);

        condContent.appendChild(condItem);
      });
    }

    // Add condition form
    const addCondForm = document.createElement('div');
    addCondForm.style.cssText = `
      padding: 10px;
      background: #eff6ff;
      border: 1px solid #bfdbfe;
      border-radius: 4px;
      display: grid;
      grid-template-columns: auto 1fr auto auto auto;
      gap: 6px;
      align-items: center;
      font-size: 11px;
    `;

    // Logic operator (if not first condition)
    if (this.conditions.length > 0) {
      const logicLabel = document.createElement('select');
      logicLabel.style.cssText = 'padding: 4px; border: 1px solid #3b82f6; border-radius: 2px; font-size: 10px;';
      logicLabel.innerHTML = `
        <option value="and" ${this.conditionLogic === 'and' ? 'selected' : ''}>AND</option>
        <option value="or" ${this.conditionLogic === 'or' ? 'selected' : ''}>OR</option>
      `;
      logicLabel.addEventListener('change', (e) => this.conditionLogic = e.target.value);
      addCondForm.appendChild(logicLabel);
    } else {
      addCondForm.appendChild(document.createElement('span'));
    }

    // Field selector
    const fieldSelect = document.createElement('select');
    fieldSelect.className = 'new-field-select';
    fieldSelect.style.cssText = 'padding: 4px; border: 1px solid #d1d5db; border-radius: 2px; font-size: 11px;';
    fieldSelect.innerHTML = '<option value="">Field</option>';
    this.dataFields.forEach(f => {
      fieldSelect.innerHTML += `<option value="${f.name}">${f.name}</option>`;
    });
    addCondForm.appendChild(fieldSelect);

    // Operator selector
    const operatorSelect = document.createElement('select');
    operatorSelect.className = 'new-operator-select';
    operatorSelect.style.cssText = 'padding: 4px; border: 1px solid #d1d5db; border-radius: 2px; font-size: 11px; width: 80px;';
    operatorSelect.innerHTML = `
      <option value="=">=</option>
      <option value="!=">!=</option>
      <option value="<">&lt;</option>
      <option value=">">&gt;</option>
      <option value="<=">&lt;=</option>
      <option value=">=">&gt;=</option>
      <option value="contains">contains</option>
      <option value="starts-with">starts-with</option>
      <option value="ends-with">ends-with</option>
      <option value="is-null">is-null</option>
    `;
    addCondForm.appendChild(operatorSelect);

    // Value input
    const valueInput = document.createElement('input');
    valueInput.type = 'text';
    valueInput.className = 'new-condition-value';
    valueInput.placeholder = 'Value';
    valueInput.style.cssText = 'padding: 4px; border: 1px solid #d1d5db; border-radius: 2px; font-size: 11px; min-width: 80px;';
    addCondForm.appendChild(valueInput);

    // Add button
    const addBtn = document.createElement('button');
    addBtn.textContent = 'Add';
    addBtn.style.cssText = `
      padding: 4px 8px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 3px;
      font-size: 11px;
      font-weight: 500;
      cursor: pointer;
    `;
    addBtn.addEventListener('click', () => {
      const field = fieldSelect.value;
      const operator = operatorSelect.value;
      const value = valueInput.value;

      if (field && operator && (value || operator === 'is-null')) {
        this.conditions.push({ field, operator, value });
        this.render(container);
      }
    });
    addCondForm.appendChild(addBtn);

    condContent.appendChild(addCondForm);
    condSection.appendChild(condContent);
    wrapper.appendChild(condSection);

    // XSL Expression preview
    const exprSection = document.createElement('fieldset');
    exprSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const exprLegend = document.createElement('legend');
    exprLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    exprLegend.textContent = 'XSL Expression Preview';
    exprSection.appendChild(exprLegend);

    const exprPreview = document.createElement('div');
    exprPreview.className = 'xsl-expression-preview';
    exprPreview.style.cssText = `
      background: #f3f4f6;
      border: 1px solid #d1d5db;
      border-radius: 3px;
      padding: 10px;
      font-family: 'Monaco', 'Courier New', monospace;
      font-size: 11px;
      color: #1f2937;
      white-space: pre-wrap;
      word-break: break-all;
      max-height: 150px;
      overflow-y: auto;
    `;
    exprPreview.textContent = this.generateXslExpression();
    exprSection.appendChild(exprPreview);

    wrapper.appendChild(exprSection);

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
    insertBtn.textContent = 'Insert Region';
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
    insertBtn.addEventListener('click', () => this.insertRegion());
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
   * Load available data fields
   */
  loadDataFields() {
    this.dataFields = [
      { name: 'Status', type: 'text' },
      { name: 'Amount', type: 'number' },
      { name: 'Date', type: 'date' },
      { name: 'Category', type: 'text' },
      { name: 'Region', type: 'text' }
    ];
  }

  /**
   * Generate XSL expression from conditions
   */
  generateXslExpression() {
    if (this.conditions.length === 0) return 'No conditions specified';

    let expr = '';

    if (this.conditionalMode === 'simple') {
      expr = '<?if:' + this.buildConditionExpression() + '?>\n';
      expr += '  <!-- Content to show when condition is true -->\n';
      expr += '<?end if?>';
    } else if (this.conditionalMode === 'ifelse') {
      expr = '<?if:' + this.buildConditionExpression() + '?>\n';
      expr += '  <!-- Content when true -->\n';
      expr += '<?else?>\n';
      expr += '  <!-- Content when false -->\n';
      expr += '<?end if?>';
    } else if (this.conditionalMode === 'choose') {
      expr = '<?choose?>\n';
      expr += '<?when:' + this.buildConditionExpression() + '?>\n';
      expr += '  <!-- First condition content -->\n';
      expr += '<?when:condition2?>\n';
      expr += '  <!-- Second condition content -->\n';
      expr += '<?otherwise?>\n';
      expr += '  <!-- Default content -->\n';
      expr += '<?end choose?>';
    }

    return expr;
  }

  /**
   * Build condition expression
   */
  buildConditionExpression() {
    if (this.conditions.length === 0) return '';

    const exprParts = this.conditions.map(cond => {
      const field = cond.field;
      const operator = cond.operator;
      const value = cond.value;

      if (operator === 'is-null') {
        return `${field} = ''`;
      } else if (['contains', 'starts-with', 'ends-with'].includes(operator)) {
        return `${operator}(${field}, '${value}')`;
      } else {
        return `${field} ${operator} '${value}'`;
      }
    });

    return exprParts.join(` ${this.conditionLogic.toUpperCase()} `);
  }

  /**
   * Wrap selected content
   */
  async wrapSelection() {
    try {
      const expr = this.generateXslExpression();
      await this.wordApi.insertText(expr);
    } catch (error) {
      console.error('Error wrapping selection:', error);
    }
  }

  /**
   * Insert conditional region
   */
  async insertRegion() {
    try {
      if (this.conditions.length === 0) {
        alert('Please add at least one condition');
        return;
      }

      const expr = this.generateXslExpression();
      await this.wordApi.insertText(expr);
    } catch (error) {
      console.error('Error inserting region:', error);
      alert('Error inserting region: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default ConditionalRegion;
export { ConditionalRegion };
