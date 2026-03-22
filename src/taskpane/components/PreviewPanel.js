/**
 * PreviewPanel Component
 * Component for previewing report output in multiple formats (PDF, HTML, Excel, RTF, CSV, PowerPoint)
 * with parameter input, sample data loading, and preview controls for BI Publisher templates.
 */

class PreviewPanel {
  constructor(services) {
    this.services = services;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.outputFormat = 'pdf';
    this.parameters = {};
    this.previewData = null;
    this.currentPage = 1;
    this.totalPages = 1;
    this.zoomLevel = 100;
    this.isPreviewLoading = false;
  }

  /**
   * Render the PreviewPanel component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'preview-panel-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Report Preview';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Format selector
    const formatSection = document.createElement('fieldset');
    formatSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const formatLegend = document.createElement('legend');
    formatLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    formatLegend.textContent = 'Output Format';
    formatSection.appendChild(formatLegend);

    const formatContent = document.createElement('div');
    formatContent.style.cssText = 'display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px;';

    const formats = [
      { value: 'pdf', label: 'PDF', icon: '📄' },
      { value: 'html', label: 'HTML', icon: '🌐' },
      { value: 'xlsx', label: 'Excel', icon: '📊' },
      { value: 'rtf', label: 'RTF', icon: '📝' },
      { value: 'csv', label: 'CSV', icon: '📋' },
      { value: 'pptx', label: 'PowerPoint', icon: '🎯' }
    ];

    formats.forEach(format => {
      const label = document.createElement('label');
      label.style.cssText = `
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 6px;
        padding: 8px;
        border: 2px solid ${this.outputFormat === format.value ? '#3b82f6' : '#d1d5db'};
        border-radius: 4px;
        cursor: pointer;
        transition: all 0.2s;
        background: ${this.outputFormat === format.value ? '#eff6ff' : 'white'};
      `;
      label.onmouseover = () => label.style.borderColor = '#3b82f6';
      label.onmouseout = () => {
        if (this.outputFormat !== format.value) label.style.borderColor = '#d1d5db';
      };

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'output-format';
      radio.value = format.value;
      radio.checked = this.outputFormat === format.value;
      radio.style.cssText = 'cursor: pointer;';
      radio.addEventListener('change', () => {
        this.outputFormat = format.value;
        this.render(container);
      });
      label.appendChild(radio);

      const icon = document.createElement('span');
      icon.textContent = format.icon;
      icon.style.cssText = 'font-size: 20px;';
      label.appendChild(icon);

      const name = document.createElement('span');
      name.textContent = format.label;
      name.style.cssText = 'font-size: 11px; font-weight: 500;';
      label.appendChild(name);

      formatContent.appendChild(label);
    });

    formatSection.appendChild(formatContent);
    wrapper.appendChild(formatSection);

    // Parameters section
    const paramSection = document.createElement('fieldset');
    paramSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0; max-height: 150px; overflow-y: auto;';
    
    const paramLegend = document.createElement('legend');
    paramLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    paramLegend.textContent = 'Parameters';
    paramSection.appendChild(paramLegend);

    const paramContent = document.createElement('div');
    paramContent.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';

    // Sample parameters
    const sampleParams = [
      { name: 'StartDate', type: 'date', value: '2024-01-01' },
      { name: 'EndDate', type: 'date', value: '2024-12-31' },
      { name: 'Region', type: 'text', value: 'North America' }
    ];

    sampleParams.forEach(param => {
      const row = document.createElement('div');
      row.style.cssText = 'display: grid; grid-template-columns: 100px 1fr; gap: 8px; align-items: center;';

      const label = document.createElement('label');
      label.textContent = param.name + ':';
      label.style.cssText = 'font-size: 11px; color: #6b7280; font-weight: 500;';
      row.appendChild(label);

      const input = document.createElement('input');
      input.type = param.type === 'date' ? 'date' : 'text';
      input.value = param.value;
      input.style.cssText = 'padding: 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
      input.addEventListener('change', (e) => {
        this.parameters[param.name] = e.target.value;
      });
      row.appendChild(input);

      paramContent.appendChild(row);
    });

    paramSection.appendChild(paramContent);
    wrapper.appendChild(paramSection);

    // Preview controls
    const controlSection = document.createElement('div');
    controlSection.style.cssText = 'display: flex; gap: 8px; align-items: center;';

    const previewBtn = document.createElement('button');
    previewBtn.textContent = 'Preview Report';
    previewBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #3b82f6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.2s;
    `;
    previewBtn.onmouseover = () => previewBtn.style.background = '#2563eb';
    previewBtn.onmouseout = () => previewBtn.style.background = '#3b82f6';
    previewBtn.addEventListener('click', () => this.generatePreview());
    controlSection.appendChild(previewBtn);

    const loadSampleBtn = document.createElement('button');
    loadSampleBtn.textContent = 'Load Sample Data';
    loadSampleBtn.style.cssText = `
      padding: 8px 12px;
      background: #e5e7eb;
      color: #374151;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
    `;
    loadSampleBtn.addEventListener('click', () => this.loadSampleData());
    controlSection.appendChild(loadSampleBtn);

    wrapper.appendChild(controlSection);

    // Preview area with zoom controls
    const previewContainer = document.createElement('div');
    previewContainer.style.cssText = `
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 8px;
      overflow: hidden;
      min-height: 0;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      background: #f9fafb;
    `;

    // Zoom and page controls
    const controlBar = document.createElement('div');
    controlBar.style.cssText = `
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 8px 12px;
      border-bottom: 1px solid #d1d5db;
      background: white;
    `;

    const zoomLabel = document.createElement('label');
    zoomLabel.style.cssText = 'font-size: 11px; color: #6b7280; white-space: nowrap;';
    zoomLabel.textContent = 'Zoom:';
    controlBar.appendChild(zoomLabel);

    const zoomSelect = document.createElement('select');
    zoomSelect.style.cssText = 'padding: 4px 6px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
    zoomSelect.innerHTML = `
      <option value="50">50%</option>
      <option value="75">75%</option>
      <option value="100" selected>100%</option>
      <option value="125">125%</option>
      <option value="150">150%</option>
      <option value="200">200%</option>
    `;
    zoomSelect.addEventListener('change', (e) => {
      this.zoomLevel = parseInt(e.target.value);
      this.updatePreviewZoom();
    });
    controlBar.appendChild(zoomSelect);

    controlBar.appendChild(document.createElement('div')); // Spacer

    const pageLabel = document.createElement('label');
    pageLabel.style.cssText = 'font-size: 11px; color: #6b7280;';
    pageLabel.textContent = 'Page:';
    controlBar.appendChild(pageLabel);

    const pageInput = document.createElement('input');
    pageInput.type = 'number';
    pageInput.min = '1';
    pageInput.value = this.currentPage;
    pageInput.style.cssText = 'width: 40px; padding: 4px; border: 1px solid #d1d5db; border-radius: 3px; font-size: 11px;';
    pageInput.addEventListener('change', (e) => {
      const page = parseInt(e.target.value);
      if (page >= 1 && page <= this.totalPages) {
        this.currentPage = page;
        this.updatePreviewPage();
      }
    });
    controlBar.appendChild(pageInput);

    const pageInfo = document.createElement('span');
    pageInfo.style.cssText = 'font-size: 11px; color: #6b7280;';
    pageInfo.textContent = ` / ${this.totalPages}`;
    controlBar.appendChild(pageInfo);

    previewContainer.appendChild(controlBar);

    // Preview content area
    const previewContent = document.createElement('div');
    previewContent.className = 'preview-content';
    previewContent.style.cssText = `
      flex: 1;
      overflow: auto;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #f3f4f6;
      padding: 16px;
    `;

    const emptyMsg = document.createElement('div');
    emptyMsg.style.cssText = `
      text-align: center;
      color: #9ca3af;
      font-size: 13px;
    `;
    emptyMsg.innerHTML = '<p>No preview available</p><p style="font-size: 11px; margin-top: 8px;">Click "Preview Report" to generate preview</p>';
    previewContent.appendChild(emptyMsg);

    previewContainer.appendChild(previewContent);
    wrapper.appendChild(previewContainer);

    // Bottom action buttons
    const actionArea = document.createElement('div');
    actionArea.style.cssText = `
      display: flex;
      gap: 8px;
      padding-top: 12px;
      border-top: 1px solid #d1d5db;
    `;

    const printBtn = document.createElement('button');
    printBtn.textContent = 'Print';
    printBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #8b5cf6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    `;
    printBtn.addEventListener('click', () => this.printPreview());
    actionArea.appendChild(printBtn);

    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save As';
    saveBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    `;
    saveBtn.addEventListener('click', () => this.savePreview());
    actionArea.appendChild(saveBtn);

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
    actionArea.appendChild(closeBtn);

    wrapper.appendChild(actionArea);
    container.appendChild(wrapper);
  }

  /**
   * Generate preview
   */
  async generatePreview() {
    try {
      const msg = document.querySelector('.preview-content');
      if (msg) {
        msg.innerHTML = '<div style="text-align: center; color: #6b7280;"><p>Generating preview...</p></div>';
      }

      // Simulate preview generation
      await new Promise(resolve => setTimeout(resolve, 1500));

      this.totalPages = Math.floor(Math.random() * 5) + 1;
      this.currentPage = 1;

      const previewContent = document.querySelector('.preview-content');
      if (previewContent) {
        previewContent.innerHTML = `
          <div style="
            width: 100%;
            height: 100%;
            background: white;
            border: 1px solid #d1d5db;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            overflow: auto;
            transform: scale(${this.zoomLevel / 100});
            transform-origin: top left;
          ">
            <h2 style="margin: 0 0 12px 0; font-size: 18px; color: #1f2937;">Report Preview</h2>
            <p style="color: #6b7280; font-size: 12px; line-height: 1.6;">
              This is a sample preview of your report in ${this.outputFormat.toUpperCase()} format.
            </p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 16px;">
              <tr style="background: #f3f4f6;">
                <th style="border: 1px solid #d1d5db; padding: 8px; text-align: left; font-size: 12px;">Field 1</th>
                <th style="border: 1px solid #d1d5db; padding: 8px; text-align: left; font-size: 12px;">Field 2</th>
                <th style="border: 1px solid #d1d5db; padding: 8px; text-align: left; font-size: 12px;">Field 3</th>
              </tr>
              <tr>
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 1</td>
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 2</td>
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 3</td>
              </tr>
              <tr style="background: #f9fafb;">
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 4</td>
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 5</td>
                <td style="border: 1px solid #d1d5db; padding: 8px; font-size: 12px;">Data 6</td>
              </tr>
            </table>
          </div>
        `;
      }
    } catch (error) {
      console.error('Error generating preview:', error);
    }
  }

  /**
   * Load sample data
   */
  loadSampleData() {
    console.log('Loading sample data...');
    // Sample data would be loaded here
  }

  /**
   * Update preview zoom level
   */
  updatePreviewZoom() {
    const preview = document.querySelector('.preview-content > div');
    if (preview) {
      preview.style.transform = `scale(${this.zoomLevel / 100})`;
    }
  }

  /**
   * Update preview page
   */
  updatePreviewPage() {
    const pageInput = document.querySelector('input[type="number"]');
    if (pageInput) {
      pageInput.value = this.currentPage;
    }
  }

  /**
   * Print preview
   */
  printPreview() {
    window.print();
  }

  /**
   * Save preview
   */
  savePreview() {
    const filename = `report.${this.outputFormat}`;
    console.log(`Saving as ${filename}`);
    alert(`Saving as ${filename}`);
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default PreviewPanel;
export { PreviewPanel };
