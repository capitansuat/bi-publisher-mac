/**
 * FormatHelper Component
 * Comprehensive formatting tools including number formats, date formats,
 * conditional formatting, watermarks, signatures, headers/footers, and page setup.
 */

class FormatHelper {
  constructor(services) {
    this.services = services;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.activeTab = 'numbers';
    this.numberFormat = '0,000.00';
    this.dateFormat = 'mm/dd/yyyy';
    this.currencyLocale = 'en-US';
    this.conditionalRules = [];
    this.watermarkText = '';
    this.pageSetup = {
      orientation: 'portrait',
      paperSize: 'letter',
      marginTop: 1,
      marginBottom: 1,
      marginLeft: 1,
      marginRight: 1
    };
  }

  /**
   * Render the FormatHelper component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'format-helper-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Format Helper';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Tab selector
    const tabBar = document.createElement('div');
    tabBar.style.cssText = `
      display: flex;
      gap: 4px;
      border-bottom: 2px solid #e5e7eb;
      overflow-x: auto;
    `;

    const tabs = [
      { id: 'numbers', label: 'Number Formats' },
      { id: 'dates', label: 'Date Formats' },
      { id: 'conditional', label: 'Conditional' },
      { id: 'watermark', label: 'Watermark' },
      { id: 'signature', label: 'Signature' },
      { id: 'headers', label: 'Headers/Footers' },
      { id: 'pagesetup', label: 'Page Setup' }
    ];

    tabs.forEach(tab => {
      const tabBtn = document.createElement('button');
      tabBtn.textContent = tab.label;
      tabBtn.style.cssText = `
        padding: 8px 12px;
        background: transparent;
        border: none;
        border-bottom: 3px solid ${this.activeTab === tab.id ? '#3b82f6' : 'transparent'};
        color: ${this.activeTab === tab.id ? '#3b82f6' : '#6b7280'};
        font-size: 12px;
        font-weight: ${this.activeTab === tab.id ? '600' : '500'};
        cursor: pointer;
        white-space: nowrap;
        transition: all 0.2s;
      `;
      tabBtn.onmouseover = () => tabBtn.style.color = '#3b82f6';
      tabBtn.onmouseout = () => {
        if (this.activeTab !== tab.id) tabBtn.style.color = '#6b7280';
      };
      tabBtn.addEventListener('click', () => {
        this.activeTab = tab.id;
        this.render(container);
      });
      tabBar.appendChild(tabBtn);
    });

    wrapper.appendChild(tabBar);

    // Tab content
    const contentArea = document.createElement('div');
    contentArea.style.cssText = `
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 12px;
      overflow-y: auto;
      min-height: 0;
    `;

    switch (this.activeTab) {
      case 'numbers':
        this.renderNumberFormats(contentArea);
        break;
      case 'dates':
        this.renderDateFormats(contentArea);
        break;
      case 'conditional':
        this.renderConditionalFormats(contentArea);
        break;
      case 'watermark':
        this.renderWatermark(contentArea);
        break;
      case 'signature':
        this.renderSignature(contentArea);
        break;
      case 'headers':
        this.renderHeadersFooters(contentArea);
        break;
      case 'pagesetup':
        this.renderPageSetup(contentArea);
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
   * Render number format options
   */
  renderNumberFormats(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Number Format Presets';
    section.appendChild(legend);

    const presets = [
      { format: '0', label: 'Whole Number (1234)', sample: '1234' },
      { format: '0.00', label: 'Decimal (1234.57)', sample: '1234.57' },
      { format: '#,##0', label: 'Thousands Separator (1,234)', sample: '1,234' },
      { format: '#,##0.00', label: 'Currency Format (1,234.57)', sample: '1,234.57' },
      { format: '0%', label: 'Percentage (123%)', sample: '123%' },
      { format: '0.00%', label: 'Percentage Decimal (12.34%)', sample: '12.34%' },
      { format: '0.00E+00', label: 'Scientific (1.23E+03)', sample: '1.23E+03' }
    ];

    presets.forEach(preset => {
      const label = document.createElement('label');
      label.style.cssText = `
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        margin-bottom: 6px;
        cursor: pointer;
        transition: background 0.15s;
      `;
      label.onmouseover = () => label.style.background = '#f9fafb';
      label.onmouseout = () => label.style.background = 'transparent';

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'number-preset';
      radio.value = preset.format;
      radio.checked = this.numberFormat === preset.format;
      radio.style.cssText = 'cursor: pointer;';
      radio.addEventListener('change', () => this.numberFormat = preset.format);
      label.appendChild(radio);

      const info = document.createElement('div');
      info.style.cssText = 'flex: 1;';
      const labelDiv = document.createElement('div');
      labelDiv.textContent = preset.label;
      labelDiv.style.cssText = 'font-size: 12px; font-weight: 500; color: #1f2937;';
      info.appendChild(labelDiv);
      const sampleDiv = document.createElement('div');
      sampleDiv.textContent = `Sample: ${preset.sample}`;
      sampleDiv.style.cssText = 'font-size: 11px; color: #6b7280;';
      info.appendChild(sampleDiv);
      label.appendChild(info);

      section.appendChild(label);
    });

    container.appendChild(section);

    // Custom format
    const customSection = document.createElement('fieldset');
    customSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const customLegend = document.createElement('legend');
    customLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    customLegend.textContent = 'Custom Format';
    customSection.appendChild(customLegend);

    const customInput = document.createElement('input');
    customInput.type = 'text';
    customInput.value = this.numberFormat;
    customInput.placeholder = 'Enter custom format code';
    customInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
    customInput.addEventListener('change', (e) => this.numberFormat = e.target.value);
    customSection.appendChild(customInput);

    container.appendChild(customSection);
  }

  /**
   * Render date format options
   */
  renderDateFormats(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Date Format Presets';
    section.appendChild(legend);

    const today = new Date();
    const presets = [
      { format: 'mm/dd/yyyy', label: 'US Format', sample: `${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}/${today.getFullYear()}` },
      { format: 'dd/mm/yyyy', label: 'European Format', sample: `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}` },
      { format: 'yyyy-mm-dd', label: 'ISO Format', sample: `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}` },
      { format: 'dddd, mmmm d, yyyy', label: 'Full Date', sample: `${today.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}` },
      { format: 'mmmm d, yyyy', label: 'Month Day, Year', sample: `${today.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}` }
    ];

    presets.forEach(preset => {
      const label = document.createElement('label');
      label.style.cssText = `
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        margin-bottom: 6px;
        cursor: pointer;
      `;

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'date-preset';
      radio.value = preset.format;
      radio.checked = this.dateFormat === preset.format;
      radio.addEventListener('change', () => this.dateFormat = preset.format);
      label.appendChild(radio);

      const info = document.createElement('div');
      info.style.cssText = 'flex: 1;';
      const labelDiv = document.createElement('div');
      labelDiv.textContent = preset.label;
      labelDiv.style.cssText = 'font-size: 12px; font-weight: 500; color: #1f2937;';
      info.appendChild(labelDiv);
      const sampleDiv = document.createElement('div');
      sampleDiv.textContent = preset.sample;
      sampleDiv.style.cssText = 'font-size: 11px; color: #6b7280;';
      info.appendChild(sampleDiv);
      label.appendChild(info);

      section.appendChild(label);
    });

    container.appendChild(section);
  }

  /**
   * Render conditional formatting
   */
  renderConditionalFormats(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Conditional Formatting Rules';
    section.appendChild(legend);

    const content = document.createElement('div');
    content.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const label = document.createElement('label');
    label.style.cssText = 'font-size: 12px; font-weight: 500; color: #1f2937;';
    label.textContent = 'Add a conditional format rule:';
    content.appendChild(label);

    const formDiv = document.createElement('div');
    formDiv.style.cssText = 'display: grid; grid-template-columns: 1fr auto; gap: 8px;';

    const ruleInput = document.createElement('input');
    ruleInput.type = 'text';
    ruleInput.placeholder = 'e.g., value > 1000';
    ruleInput.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
    formDiv.appendChild(ruleInput);

    const addBtn = document.createElement('button');
    addBtn.textContent = 'Add Rule';
    addBtn.style.cssText = `
      padding: 6px 12px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 3px;
      font-size: 11px;
      cursor: pointer;
    `;
    addBtn.addEventListener('click', () => {
      if (ruleInput.value) {
        this.conditionalRules.push({ rule: ruleInput.value, format: 'bold' });
        ruleInput.value = '';
        this.render(container);
      }
    });
    formDiv.appendChild(addBtn);

    content.appendChild(formDiv);

    if (this.conditionalRules.length > 0) {
      const rulesList = document.createElement('div');
      rulesList.style.cssText = 'display: flex; flex-direction: column; gap: 6px; margin-top: 8px;';

      this.conditionalRules.forEach((rule, index) => {
        const ruleDiv = document.createElement('div');
        ruleDiv.style.cssText = `
          padding: 8px;
          background: #f3f4f6;
          border-radius: 3px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          font-size: 11px;
        `;
        
        ruleDiv.innerHTML = `<span>${rule.rule} → ${rule.format}</span>`;
        
        const removeBtn = document.createElement('button');
        removeBtn.textContent = 'Remove';
        removeBtn.style.cssText = `
          padding: 2px 6px;
          background: transparent;
          border: none;
          color: #ef4444;
          cursor: pointer;
          font-size: 10px;
        `;
        removeBtn.addEventListener('click', () => {
          this.conditionalRules.splice(index, 1);
          this.render(container);
        });
        ruleDiv.appendChild(removeBtn);

        rulesList.appendChild(ruleDiv);
      });

      content.appendChild(rulesList);
    }

    section.appendChild(content);
    container.appendChild(section);
  }

  /**
   * Render watermark options
   */
  renderWatermark(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Watermark';
    section.appendChild(legend);

    const content = document.createElement('div');
    content.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const label = document.createElement('label');
    label.style.cssText = 'display: flex; flex-direction: column; gap: 4px; font-size: 12px;';
    label.appendChild(document.createTextNode('Watermark Text:'));
    const input = document.createElement('input');
    input.type = 'text';
    input.value = this.watermarkText;
    input.placeholder = 'e.g., CONFIDENTIAL';
    input.style.cssText = 'padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
    input.addEventListener('change', (e) => this.watermarkText = e.target.value);
    label.appendChild(input);
    content.appendChild(label);

    const presets = [
      { text: 'DRAFT' },
      { text: 'CONFIDENTIAL' },
      { text: 'DO NOT COPY' },
      { text: 'OFFICIAL USE ONLY' }
    ];

    const presetsDiv = document.createElement('div');
    presetsDiv.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 6px;';

    presets.forEach(preset => {
      const btn = document.createElement('button');
      btn.textContent = preset.text;
      btn.style.cssText = `
        padding: 6px 8px;
        background: #f3f4f6;
        border: 1px solid #d1d5db;
        border-radius: 3px;
        font-size: 11px;
        cursor: pointer;
      `;
      btn.addEventListener('click', () => {
        this.watermarkText = preset.text;
        this.render(container);
      });
      presetsDiv.appendChild(btn);
    });

    content.appendChild(presetsDiv);
    section.appendChild(content);
    container.appendChild(section);
  }

  /**
   * Render signature placeholder
   */
  renderSignature(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Digital Signature';
    section.appendChild(legend);

    const content = document.createElement('div');
    content.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const desc = document.createElement('p');
    desc.textContent = 'Add a placeholder for digital signatures.';
    desc.style.cssText = 'margin: 0; font-size: 12px; color: #6b7280;';
    content.appendChild(desc);

    const insertBtn = document.createElement('button');
    insertBtn.textContent = 'Insert Signature Block';
    insertBtn.style.cssText = `
      padding: 8px 16px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
    `;
    insertBtn.addEventListener('click', () => {
      console.log('Insert signature block');
    });
    content.appendChild(insertBtn);

    section.appendChild(content);
    container.appendChild(section);
  }

  /**
   * Render headers and footers
   */
  renderHeadersFooters(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Headers & Footers';
    section.appendChild(legend);

    const content = document.createElement('div');
    content.style.cssText = 'display: flex; flex-direction: column; gap: 10px;';

    ['Header', 'Footer'].forEach(type => {
      const div = document.createElement('div');
      div.style.cssText = 'display: flex; flex-direction: column; gap: 4px; padding-bottom: 8px; border-bottom: 1px solid #e5e7eb;';

      const label = document.createElement('label');
      label.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
      label.textContent = `${type} Content:`;
      div.appendChild(label);

      const textarea = document.createElement('textarea');
      textarea.placeholder = `Enter ${type.toLowerCase()} content...`;
      textarea.style.cssText = `
        padding: 8px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        font-size: 11px;
        font-family: monospace;
        height: 60px;
        resize: vertical;
      `;
      div.appendChild(textarea);

      content.appendChild(div);
    });

    section.appendChild(content);
    container.appendChild(section);
  }

  /**
   * Render page setup options
   */
  renderPageSetup(container) {
    const section = document.createElement('fieldset');
    section.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const legend = document.createElement('legend');
    legend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    legend.textContent = 'Page Setup';
    section.appendChild(legend);

    const content = document.createElement('div');
    content.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 12px;';

    // Orientation
    const orientLabel = document.createElement('label');
    orientLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    orientLabel.appendChild(document.createTextNode('Orientation:'));
    const orientSelect = document.createElement('select');
    orientSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    orientSelect.innerHTML = `
      <option value="portrait" selected>Portrait</option>
      <option value="landscape">Landscape</option>
    `;
    orientSelect.addEventListener('change', (e) => this.pageSetup.orientation = e.target.value);
    orientLabel.appendChild(orientSelect);
    content.appendChild(orientLabel);

    // Paper size
    const paperLabel = document.createElement('label');
    paperLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    paperLabel.appendChild(document.createTextNode('Paper Size:'));
    const paperSelect = document.createElement('select');
    paperSelect.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    paperSelect.innerHTML = `
      <option value="letter" selected>Letter (8.5" × 11")</option>
      <option value="a4">A4 (210mm × 297mm)</option>
      <option value="legal">Legal (8.5" × 14")</option>
    `;
    paperSelect.addEventListener('change', (e) => this.pageSetup.paperSize = e.target.value);
    paperLabel.appendChild(paperSelect);
    content.appendChild(paperLabel);

    // Margins
    ['Top', 'Bottom', 'Left', 'Right'].forEach((side, index) => {
      if (index === 0 || index === 2) {
        const label = document.createElement('label');
        label.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
        label.appendChild(document.createTextNode(`Margin ${side} (in):`));
        const input = document.createElement('input');
        input.type = 'number';
        input.min = '0.5';
        input.max = '2';
        input.step = '0.1';
        input.value = this.pageSetup[`margin${side}`] || 1;
        input.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
        input.addEventListener('change', (e) => {
          this.pageSetup[`margin${side}`] = parseFloat(e.target.value);
        });
        label.appendChild(input);
        content.appendChild(label);
      }
    });

    section.appendChild(content);
    container.appendChild(section);
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default FormatHelper;
export { FormatHelper };
