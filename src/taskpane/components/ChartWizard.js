/**
 * ChartWizard Component
 * 3-step wizard for creating charts with multiple chart types,
 * data mapping, and appearance customization for BI Publisher templates.
 */

class ChartWizard {
  constructor(services) {
    this.services = services;
    this.xmlParser = services.xmlParser;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.currentStep = 1;
    this.chartType = 'bar';
    this.dataFields = [];
    this.categoryField = null;
    this.valueFields = [];
    this.seriesField = null;
    this.chartConfig = {
      title: 'Chart Title',
      legend: true,
      colors: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'],
      size: { width: 400, height: 300 },
      labels: true,
      show3d: false
    };
  }

  /**
   * Render the ChartWizard component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'chart-wizard-wrapper';
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
    header.style.cssText = 'display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;';
    
    const title = document.createElement('h3');
    title.textContent = 'Chart Wizard';
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
    nextBtn.textContent = this.currentStep === 3 ? 'Insert Chart' : 'Next';
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
        this.insertChart(container);
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
   * Step 1: Select chart type
   */
  renderStep1(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    const label = document.createElement('label');
    label.textContent = 'Select Chart Type:';
    label.style.cssText = 'font-size: 13px; font-weight: 600; color: #1f2937;';
    section.appendChild(label);

    const chartTypes = [
      { id: 'bar', name: 'Bar', icon: '📊' },
      { id: 'line', name: 'Line', icon: '📈' },
      { id: 'pie', name: 'Pie', icon: '🥧' },
      { id: 'area', name: 'Area', icon: '📉' },
      { id: 'scatter', name: 'Scatter', icon: '⚡' },
      { id: 'combo', name: 'Combo', icon: '🔀' },
      { id: 'gauge', name: 'Gauge', icon: '🎯' },
      { id: 'funnel', name: 'Funnel', icon: '⏳' },
      { id: 'radar', name: 'Radar', icon: '🔷' }
    ];

    const grid = document.createElement('div');
    grid.style.cssText = `
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
    `;

    chartTypes.forEach(type => {
      const card = document.createElement('div');
      card.className = 'chart-type-card';
      card.style.cssText = `
        border: 2px solid ${this.chartType === type.id ? '#3b82f6' : '#d1d5db'};
        border-radius: 8px;
        padding: 16px;
        text-align: center;
        cursor: pointer;
        transition: all 0.2s;
        background: ${this.chartType === type.id ? '#eff6ff' : 'white'};
      `;
      card.onmouseover = () => {
        card.style.borderColor = '#3b82f6';
        card.style.boxShadow = '0 2px 8px rgba(59, 130, 246, 0.1)';
      };
      card.onmouseout = () => {
        if (this.chartType !== type.id) {
          card.style.borderColor = '#d1d5db';
          card.style.boxShadow = 'none';
        }
      };
      card.addEventListener('click', () => {
        this.chartType = type.id;
        this.render(container);
      });

      const icon = document.createElement('div');
      icon.textContent = type.icon;
      icon.style.cssText = 'font-size: 32px; margin-bottom: 8px;';
      card.appendChild(icon);

      const name = document.createElement('div');
      name.textContent = type.name;
      name.style.cssText = 'font-size: 13px; font-weight: 500; color: #1f2937;';
      card.appendChild(name);

      grid.appendChild(card);
    });

    section.appendChild(grid);
    container.appendChild(section);
  }

  /**
   * Step 2: Data mapping
   */
  renderStep2(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    this.loadDataFields();

    // Category field
    const categorySection = document.createElement('fieldset');
    categorySection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const categoryLegend = document.createElement('legend');
    categoryLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    categoryLegend.textContent = 'Category (X-Axis)';
    categorySection.appendChild(categoryLegend);

    const categorySelect = document.createElement('select');
    categorySelect.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
    categorySelect.innerHTML = '<option>-- Select field --</option>';
    this.dataFields.forEach(field => {
      categorySelect.innerHTML += `<option value="${field.name}">${field.name}</option>`;
    });
    categorySelect.addEventListener('change', (e) => this.categoryField = e.target.value);
    categorySection.appendChild(categorySelect);

    section.appendChild(categorySection);

    // Value fields
    const valueSection = document.createElement('fieldset');
    valueSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const valueLegend = document.createElement('legend');
    valueLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    valueLegend.textContent = 'Values (Y-Axis)';
    valueSection.appendChild(valueLegend);

    const valueContent = document.createElement('div');
    valueContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    this.dataFields.forEach(field => {
      if (field.type === 'number') {
        const label = document.createElement('label');
        label.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = field.name;
        checkbox.addEventListener('change', (e) => {
          if (e.target.checked) {
            this.valueFields.push(field.name);
          } else {
            this.valueFields = this.valueFields.filter(f => f !== field.name);
          }
        });
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(field.name));
        valueContent.appendChild(label);
      }
    });

    valueSection.appendChild(valueContent);
    section.appendChild(valueSection);

    // Series field (for certain chart types)
    if (['combo', 'line', 'bar'].includes(this.chartType)) {
      const seriesSection = document.createElement('fieldset');
      seriesSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
      
      const seriesLegend = document.createElement('legend');
      seriesLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
      seriesLegend.textContent = 'Series (Optional)';
      seriesSection.appendChild(seriesLegend);

      const seriesSelect = document.createElement('select');
      seriesSelect.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
      seriesSelect.innerHTML = '<option>-- None --</option>';
      this.dataFields.forEach(field => {
        if (field.type === 'text') {
          seriesSelect.innerHTML += `<option value="${field.name}">${field.name}</option>`;
        }
      });
      seriesSelect.addEventListener('change', (e) => this.seriesField = e.target.value);
      seriesSection.appendChild(seriesSelect);

      section.appendChild(seriesSection);
    }

    container.appendChild(section);
  }

  /**
   * Step 3: Appearance and configuration
   */
  renderStep3(container) {
    const section = document.createElement('div');
    section.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

    // Title
    const titleSection = document.createElement('fieldset');
    titleSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const titleLegend = document.createElement('legend');
    titleLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    titleLegend.textContent = 'Chart Title';
    titleSection.appendChild(titleLegend);

    const titleInput = document.createElement('input');
    titleInput.type = 'text';
    titleInput.value = this.chartConfig.title;
    titleInput.style.cssText = 'width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;';
    titleInput.addEventListener('change', (e) => this.chartConfig.title = e.target.value);
    titleSection.appendChild(titleInput);

    section.appendChild(titleSection);

    // Size
    const sizeSection = document.createElement('fieldset');
    sizeSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const sizeLegend = document.createElement('legend');
    sizeLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    sizeLegend.textContent = 'Size';
    sizeSection.appendChild(sizeLegend);

    const sizeContent = document.createElement('div');
    sizeContent.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 8px; font-size: 12px;';

    const widthLabel = document.createElement('label');
    widthLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    widthLabel.appendChild(document.createTextNode('Width (px):'));
    const widthInput = document.createElement('input');
    widthInput.type = 'number';
    widthInput.value = this.chartConfig.size.width;
    widthInput.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    widthInput.addEventListener('change', (e) => this.chartConfig.size.width = parseInt(e.target.value));
    widthLabel.appendChild(widthInput);
    sizeContent.appendChild(widthLabel);

    const heightLabel = document.createElement('label');
    heightLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    heightLabel.appendChild(document.createTextNode('Height (px):'));
    const heightInput = document.createElement('input');
    heightInput.type = 'number';
    heightInput.value = this.chartConfig.size.height;
    heightInput.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px;';
    heightInput.addEventListener('change', (e) => this.chartConfig.size.height = parseInt(e.target.value));
    heightLabel.appendChild(heightInput);
    sizeContent.appendChild(heightLabel);

    sizeSection.appendChild(sizeContent);
    section.appendChild(sizeSection);

    // Legend and labels
    const optionsSection = document.createElement('fieldset');
    optionsSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const optionsLegend = document.createElement('legend');
    optionsLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    optionsLegend.textContent = 'Options';
    optionsSection.appendChild(optionsLegend);

    const legendCheckbox = document.createElement('label');
    legendCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 8px; cursor: pointer; font-size: 12px;';
    const legendInput = document.createElement('input');
    legendInput.type = 'checkbox';
    legendInput.checked = this.chartConfig.legend;
    legendInput.addEventListener('change', (e) => this.chartConfig.legend = e.target.checked);
    legendCheckbox.appendChild(legendInput);
    legendCheckbox.appendChild(document.createTextNode('Show legend'));
    optionsSection.appendChild(legendCheckbox);

    const labelsCheckbox = document.createElement('label');
    labelsCheckbox.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 8px; cursor: pointer; font-size: 12px;';
    const labelsInput = document.createElement('input');
    labelsInput.type = 'checkbox';
    labelsInput.checked = this.chartConfig.labels;
    labelsInput.addEventListener('change', (e) => this.chartConfig.labels = e.target.checked);
    labelsCheckbox.appendChild(labelsInput);
    labelsCheckbox.appendChild(document.createTextNode('Show data labels'));
    optionsSection.appendChild(labelsCheckbox);

    if (['bar', 'line', 'area', 'pie'].includes(this.chartType)) {
      const enable3d = document.createElement('label');
      enable3d.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';
      const enable3dInput = document.createElement('input');
      enable3dInput.type = 'checkbox';
      enable3dInput.checked = this.chartConfig.show3d;
      enable3dInput.addEventListener('change', (e) => this.chartConfig.show3d = e.target.checked);
      enable3d.appendChild(enable3dInput);
      enable3d.appendChild(document.createTextNode('3D Effect'));
      optionsSection.appendChild(enable3d);
    }

    section.appendChild(optionsSection);
    container.appendChild(section);
  }

  /**
   * Load data fields
   */
  loadDataFields() {
    this.dataFields = [
      { name: 'Category', type: 'text' },
      { name: 'Value1', type: 'number' },
      { name: 'Value2', type: 'number' },
      { name: 'Series', type: 'text' }
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
   * Insert chart into document
   */
  async insertChart(container) {
    try {
      // Generate BI Publisher chart XML
      let chartXml = '<bi:chart>\n';
      chartXml += `  <bi:type>${this.chartType}</bi:type>\n`;
      chartXml += `  <bi:title>${this.chartConfig.title}</bi:title>\n`;
      chartXml += `  <bi:legend>${this.chartConfig.legend ? 'true' : 'false'}</bi:legend>\n`;
      chartXml += `  <bi:labels>${this.chartConfig.labels ? 'true' : 'false'}</bi:labels>\n`;
      
      if (this.categoryField) {
        chartXml += `  <bi:categoryField>${this.categoryField}</bi:categoryField>\n`;
      }
      
      this.valueFields.forEach(field => {
        chartXml += `  <bi:valueField>${field}</bi:valueField>\n`;
      });

      if (this.seriesField) {
        chartXml += `  <bi:seriesField>${this.seriesField}</bi:seriesField>\n`;
      }

      chartXml += `  <bi:width>${this.chartConfig.size.width}</bi:width>\n`;
      chartXml += `  <bi:height>${this.chartConfig.size.height}</bi:height>\n`;
      chartXml += '</bi:chart>\n';

      await this.wordApi.insertText(chartXml);
      container.innerHTML = '';
    } catch (error) {
      console.error('Error inserting chart:', error);
      alert('Error inserting chart: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default ChartWizard;
export { ChartWizard };
