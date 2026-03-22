/**
 * BarcodeInserter Component
 * Component for inserting various barcode and QR code types with
 * configurable properties and live preview for BI Publisher templates.
 */

class BarcodeInserter {
  constructor(services) {
    this.services = services;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.barcodeType = 'code128';
    this.dataSource = 'static';
    this.staticValue = '123456789';
    this.dataField = null;
    this.width = 200;
    this.height = 80;
    this.showText = true;
    this.fontSize = 12;
    this.batchMode = false;
    this.dataFields = [];
  }

  /**
   * Render the BarcodeInserter component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'barcode-inserter-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Insert Barcode';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Barcode type selector
    const typeSection = document.createElement('fieldset');
    typeSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const typeLegend = document.createElement('legend');
    typeLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    typeLegend.textContent = 'Barcode Type';
    typeSection.appendChild(typeLegend);

    const typeContent = document.createElement('div');
    typeContent.style.cssText = 'display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px;';

    const barcodeTypes = [
      { value: 'code128', label: 'Code 128' },
      { value: 'code39', label: 'Code 39' },
      { value: 'ean13', label: 'EAN-13' },
      { value: 'upca', label: 'UPC-A' },
      { value: 'qr', label: 'QR Code' },
      { value: 'datamatrix', label: 'Data Matrix' },
      { value: 'pdf417', label: 'PDF417' },
      { value: 'itf', label: 'ITF' }
    ];

    barcodeTypes.forEach(type => {
      const label = document.createElement('label');
      label.style.cssText = `
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 8px;
        border: 1px solid ${this.barcodeType === type.value ? '#3b82f6' : '#d1d5db'};
        border-radius: 4px;
        cursor: pointer;
        transition: all 0.2s;
        background: ${this.barcodeType === type.value ? '#eff6ff' : 'white'};
      `;
      label.onmouseover = () => label.style.borderColor = '#3b82f6';
      label.onmouseout = () => {
        if (this.barcodeType !== type.value) label.style.borderColor = '#d1d5db';
      };

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'barcode-type';
      radio.value = type.value;
      radio.checked = this.barcodeType === type.value;
      radio.style.cssText = 'cursor: pointer;';
      radio.addEventListener('change', () => {
        this.barcodeType = type.value;
        this.render(container);
      });
      label.appendChild(radio);

      label.appendChild(document.createTextNode(type.label));
      typeContent.appendChild(label);
    });

    typeSection.appendChild(typeContent);
    wrapper.appendChild(typeSection);

    // Data source selector
    const dataSection = document.createElement('fieldset');
    dataSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const dataLegend = document.createElement('legend');
    dataLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    dataLegend.textContent = 'Data Source';
    dataSection.appendChild(dataLegend);

    const dataContent = document.createElement('div');
    dataContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    const staticRadio = document.createElement('label');
    staticRadio.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const staticInput = document.createElement('input');
    staticInput.type = 'radio';
    staticInput.name = 'data-source';
    staticInput.value = 'static';
    staticInput.checked = this.dataSource === 'static';
    staticInput.addEventListener('change', () => {
      this.dataSource = 'static';
      this.render(container);
    });
    staticRadio.appendChild(staticInput);
    staticRadio.appendChild(document.createTextNode('Static Value'));
    dataContent.appendChild(staticRadio);

    if (this.dataSource === 'static') {
      const valueInput = document.createElement('input');
      valueInput.type = 'text';
      valueInput.value = this.staticValue;
      valueInput.style.cssText = 'padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; margin-left: 24px;';
      valueInput.addEventListener('change', (e) => this.staticValue = e.target.value);
      dataContent.appendChild(valueInput);
    }

    const fieldRadio = document.createElement('label');
    fieldRadio.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const fieldInput = document.createElement('input');
    fieldInput.type = 'radio';
    fieldInput.name = 'data-source';
    fieldInput.value = 'field';
    fieldInput.checked = this.dataSource === 'field';
    fieldInput.addEventListener('change', () => {
      this.dataSource = 'field';
      this.render(container);
    });
    fieldRadio.appendChild(fieldInput);
    fieldRadio.appendChild(document.createTextNode('Data Field'));
    dataContent.appendChild(fieldRadio);

    if (this.dataSource === 'field') {
      this.loadDataFields();
      const fieldSelect = document.createElement('select');
      fieldSelect.style.cssText = 'padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; margin-left: 24px;';
      fieldSelect.innerHTML = '<option value="">-- Select field --</option>';
      this.dataFields.forEach(field => {
        fieldSelect.innerHTML += `<option value="${field.name}">${field.name}</option>`;
      });
      fieldSelect.addEventListener('change', (e) => this.dataField = e.target.value);
      dataContent.appendChild(fieldSelect);
    }

    dataSection.appendChild(dataContent);
    wrapper.appendChild(dataSection);

    // Properties section
    const propsSection = document.createElement('fieldset');
    propsSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const propsLegend = document.createElement('legend');
    propsLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    propsLegend.textContent = 'Properties';
    propsSection.appendChild(propsLegend);

    const propsContent = document.createElement('div');
    propsContent.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 12px;';

    const widthLabel = document.createElement('label');
    widthLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    widthLabel.appendChild(document.createTextNode('Width (px):'));
    const widthInput = document.createElement('input');
    widthInput.type = 'number';
    widthInput.value = this.width;
    widthInput.min = '50';
    widthInput.max = '600';
    widthInput.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    widthInput.addEventListener('change', (e) => this.width = parseInt(e.target.value));
    widthLabel.appendChild(widthInput);
    propsContent.appendChild(widthLabel);

    const heightLabel = document.createElement('label');
    heightLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    heightLabel.appendChild(document.createTextNode('Height (px):'));
    const heightInput = document.createElement('input');
    heightInput.type = 'number';
    heightInput.value = this.height;
    heightInput.min = '30';
    heightInput.max = '300';
    heightInput.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    heightInput.addEventListener('change', (e) => this.height = parseInt(e.target.value));
    heightLabel.appendChild(heightInput);
    propsContent.appendChild(heightLabel);

    propsSection.appendChild(propsContent);

    // Additional options
    const optionsDiv = document.createElement('div');
    optionsDiv.style.cssText = 'display: flex; flex-direction: column; gap: 8px; margin-top: 8px; padding-top: 8px; border-top: 1px solid #e5e7eb;';

    const showTextCheckbox = document.createElement('label');
    showTextCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const showTextInput = document.createElement('input');
    showTextInput.type = 'checkbox';
    showTextInput.checked = this.showText;
    showTextInput.addEventListener('change', (e) => {
      this.showText = e.target.checked;
      this.render(container);
    });
    showTextCheckbox.appendChild(showTextInput);
    showTextCheckbox.appendChild(document.createTextNode('Show text below barcode'));
    optionsDiv.appendChild(showTextCheckbox);

    if (this.showText && this.barcodeType !== 'qr' && this.barcodeType !== 'datamatrix') {
      const fontSizeLabel = document.createElement('label');
      fontSizeLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; font-size: 12px;';
      fontSizeLabel.appendChild(document.createTextNode('Font Size:'));
      const fontSizeSelect = document.createElement('select');
      fontSizeSelect.style.cssText = 'padding: 4px 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
      fontSizeSelect.innerHTML = `
        <option value="8" ${this.fontSize === 8 ? 'selected' : ''}>8pt</option>
        <option value="10" ${this.fontSize === 10 ? 'selected' : ''}>10pt</option>
        <option value="12" ${this.fontSize === 12 ? 'selected' : ''}>12pt</option>
        <option value="14" ${this.fontSize === 14 ? 'selected' : ''}>14pt</option>
      `;
      fontSizeSelect.addEventListener('change', (e) => this.fontSize = parseInt(e.target.value));
      fontSizeLabel.appendChild(fontSizeSelect);
      optionsDiv.appendChild(fontSizeLabel);
    }

    propsSection.appendChild(optionsDiv);
    wrapper.appendChild(propsSection);

    // Preview section
    const previewSection = document.createElement('fieldset');
    previewSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const previewLegend = document.createElement('legend');
    previewLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    previewLegend.textContent = 'Preview';
    previewSection.appendChild(previewLegend);

    const preview = document.createElement('div');
    preview.style.cssText = `
      width: 100%;
      height: 120px;
      background: white;
      border: 1px solid #e5e7eb;
      border-radius: 3px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #9ca3af;
      font-size: 12px;
      font-style: italic;
    `;
    preview.textContent = `${this.barcodeType.toUpperCase()} Barcode Preview (${this.width}×${this.height}px)`;
    previewSection.appendChild(preview);

    wrapper.appendChild(previewSection);

    // Batch mode
    const batchSection = document.createElement('fieldset');
    batchSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const batchLegend = document.createElement('legend');
    batchLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    batchLegend.textContent = 'Batch Mode';
    batchSection.appendChild(batchLegend);

    const batchCheckbox = document.createElement('label');
    batchCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
    const batchInput = document.createElement('input');
    batchInput.type = 'checkbox';
    batchInput.checked = this.batchMode;
    batchInput.addEventListener('change', (e) => this.batchMode = e.target.checked);
    batchCheckbox.appendChild(batchInput);
    batchCheckbox.appendChild(document.createTextNode('Create barcode for each record in repeating group'));
    batchSection.appendChild(batchCheckbox);

    wrapper.appendChild(batchSection);

    // Action buttons
    const buttonArea = document.createElement('div');
    buttonArea.style.cssText = `
      display: flex;
      gap: 8px;
      padding-top: 12px;
      border-top: 1px solid #d1d5db;
    `;

    const insertBtn = document.createElement('button');
    insertBtn.textContent = 'Insert Barcode';
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
    `;
    insertBtn.addEventListener('click', () => this.insertBarcode());
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
   * Load data fields
   */
  loadDataFields() {
    this.dataFields = [
      { name: 'ProductCode', type: 'text' },
      { name: 'SerialNumber', type: 'text' },
      { name: 'OrderNumber', type: 'text' }
    ];
  }

  /**
   * Insert barcode
   */
  async insertBarcode() {
    try {
      const dataValue = this.dataSource === 'static' ? this.staticValue : `<?${this.dataField}?>`;

      let barcodeXml = '<bi:barcode>\n';
      barcodeXml += `  <bi:type>${this.barcodeType}</bi:type>\n`;
      barcodeXml += `  <bi:data>${dataValue}</bi:data>\n`;
      barcodeXml += `  <bi:width>${this.width}</bi:width>\n`;
      barcodeXml += `  <bi:height>${this.height}</bi:height>\n`;

      if (this.showText && this.barcodeType !== 'qr' && this.barcodeType !== 'datamatrix') {
        barcodeXml += `  <bi:showText>true</bi:showText>\n`;
        barcodeXml += `  <bi:fontSize>${this.fontSize}</bi:fontSize>\n`;
      }

      if (this.batchMode) {
        barcodeXml += '  <bi:batchMode>true</bi:batchMode>\n';
      }

      barcodeXml += '</bi:barcode>\n';

      await this.wordApi.insertText(barcodeXml);
    } catch (error) {
      console.error('Error inserting barcode:', error);
      alert('Error inserting barcode: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default BarcodeInserter;
export { BarcodeInserter };
