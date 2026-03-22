/**
 * TranslationManager Component
 * Component for managing template translations with XLIFF import/export,
 * translation memory, and multi-language support for BI Publisher templates.
 */

class TranslationManager {
  constructor(services) {
    this.services = services;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.sourceLanguage = 'en';
    this.targetLanguage = 'es';
    this.translationPairs = [];
    this.translationMemory = {};
    this.extractedStrings = [];
  }

  /**
   * Render the TranslationManager component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'translation-manager-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Translation Manager';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Language selector
    const langSection = document.createElement('fieldset');
    langSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const langLegend = document.createElement('legend');
    langLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    langLegend.textContent = 'Languages';
    langSection.appendChild(langLegend);

    const langContent = document.createElement('div');
    langContent.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 12px;';

    const sourceLangLabel = document.createElement('label');
    sourceLangLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    sourceLangLabel.appendChild(document.createTextNode('Source Language:'));
    const sourceLangSelect = document.createElement('select');
    sourceLangSelect.style.cssText = 'padding: 8px; border: 1px solid #d1d5db; border-radius: 4px;';
    sourceLangSelect.innerHTML = `
      <option value="en" selected>English</option>
      <option value="es">Spanish</option>
      <option value="fr">French</option>
      <option value="de">German</option>
      <option value="it">Italian</option>
      <option value="pt">Portuguese</option>
      <option value="ja">Japanese</option>
      <option value="zh">Chinese</option>
    `;
    sourceLangSelect.addEventListener('change', (e) => {
      this.sourceLanguage = e.target.value;
    });
    sourceLangLabel.appendChild(sourceLangSelect);
    langContent.appendChild(sourceLangLabel);

    const targetLangLabel = document.createElement('label');
    targetLangLabel.style.cssText = 'display: flex; flex-direction: column; gap: 4px;';
    targetLangLabel.appendChild(document.createTextNode('Target Language:'));
    const targetLangSelect = document.createElement('select');
    targetLangSelect.style.cssText = 'padding: 8px; border: 1px solid #d1d5db; border-radius: 4px;';
    targetLangSelect.innerHTML = `
      <option value="es" selected>Spanish</option>
      <option value="en">English</option>
      <option value="fr">French</option>
      <option value="de">German</option>
      <option value="it">Italian</option>
      <option value="pt">Portuguese</option>
      <option value="ja">Japanese</option>
      <option value="zh">Chinese</option>
    `;
    targetLangSelect.addEventListener('change', (e) => {
      this.targetLanguage = e.target.value;
    });
    targetLangLabel.appendChild(targetLangSelect);
    langContent.appendChild(targetLangLabel);

    langSection.appendChild(langContent);
    wrapper.appendChild(langSection);

    // Extract and view translations
    const actionSection = document.createElement('div');
    actionSection.style.cssText = 'display: flex; gap: 8px;';

    const extractBtn = document.createElement('button');
    extractBtn.textContent = 'Extract Text';
    extractBtn.style.cssText = `
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
    extractBtn.addEventListener('click', () => this.extractTranslatableText());
    actionSection.appendChild(extractBtn);

    const importXliffBtn = document.createElement('button');
    importXliffBtn.textContent = 'Import XLIFF';
    importXliffBtn.style.cssText = `
      padding: 8px 16px;
      background: #8b5cf6;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    `;
    importXliffBtn.addEventListener('click', () => this.importXliff());
    actionSection.appendChild(importXliffBtn);

    const exportXliffBtn = document.createElement('button');
    exportXliffBtn.textContent = 'Export XLIFF';
    exportXliffBtn.style.cssText = `
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
      display: ${this.translationPairs.length === 0 ? 'none' : 'block'};
    `;
    exportXliffBtn.addEventListener('click', () => this.exportXliff());
    actionSection.appendChild(exportXliffBtn);

    wrapper.appendChild(actionSection);

    // Translations table
    const tableSection = document.createElement('fieldset');
    tableSection.style.cssText = `
      border: 1px solid #d1d5db;
      border-radius: 4px;
      padding: 0;
      margin: 0;
      overflow: hidden;
      flex: 1;
      display: flex;
      flex-direction: column;
      min-height: 0;
    `;

    const tableLegend = document.createElement('legend');
    tableLegend.style.cssText = 'padding: 0 12px; font-size: 12px; font-weight: 600; color: #1f2937;';
    tableLegend.textContent = `Translation Pairs (${this.translationPairs.length})`;
    tableSection.appendChild(tableLegend);

    const tableContainer = document.createElement('div');
    tableContainer.style.cssText = `
      flex: 1;
      overflow-y: auto;
      overflow-x: hidden;
      border-top: 1px solid #d1d5db;
      padding: 12px;
      display: flex;
      flex-direction: column;
      gap: 8px;
    `;

    if (this.translationPairs.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'color: #9ca3af; font-size: 12px; text-align: center; padding: 24px;';
      emptyMsg.textContent = 'No translations yet. Extract text from your document first.';
      tableContainer.appendChild(emptyMsg);
    } else {
      // Header
      const header = document.createElement('div');
      header.style.cssText = `
        display: grid;
        grid-template-columns: 1fr 1fr auto;
        gap: 12px;
        padding: 8px;
        background: #f3f4f6;
        border-radius: 4px;
        font-weight: 600;
        font-size: 11px;
        color: #6b7280;
        position: sticky;
        top: 0;
      `;
      header.innerHTML = `<div>Source (${this.sourceLanguage.toUpperCase()})</div><div>Target (${this.targetLanguage.toUpperCase()})</div><div></div>`;
      tableContainer.appendChild(header);

      // Rows
      this.translationPairs.forEach((pair, index) => {
        const row = document.createElement('div');
        row.style.cssText = `
          display: grid;
          grid-template-columns: 1fr 1fr auto;
          gap: 12px;
          padding: 10px;
          border: 1px solid #d1d5db;
          border-radius: 3px;
          background: white;
          font-size: 11px;
        `;

        const sourceDiv = document.createElement('div');
        sourceDiv.textContent = pair.source;
        sourceDiv.style.cssText = 'color: #1f2937; word-break: break-word;';
        row.appendChild(sourceDiv);

        const targetInput = document.createElement('input');
        targetInput.type = 'text';
        targetInput.value = pair.target || '';
        targetInput.placeholder = 'Enter translation';
        targetInput.style.cssText = `
          padding: 6px;
          border: 1px solid #d1d5db;
          border-radius: 3px;
          font-size: 11px;
          font-family: inherit;
        `;
        targetInput.addEventListener('change', (e) => {
          pair.target = e.target.value;
        });
        row.appendChild(targetInput);

        const removeBtn = document.createElement('button');
        removeBtn.textContent = '✕';
        removeBtn.style.cssText = `
          padding: 4px 8px;
          background: transparent;
          border: none;
          color: #ef4444;
          cursor: pointer;
          font-weight: bold;
        `;
        removeBtn.addEventListener('click', () => {
          this.translationPairs.splice(index, 1);
          this.render(container);
        });
        row.appendChild(removeBtn);

        tableContainer.appendChild(row);
      });
    }

    tableSection.appendChild(tableContainer);
    wrapper.appendChild(tableSection);

    // Progress bar
    if (this.translationPairs.length > 0) {
      const translated = this.translationPairs.filter(p => p.target && p.target.trim()).length;
      const progress = Math.round((translated / this.translationPairs.length) * 100);

      const progressSection = document.createElement('div');
      progressSection.style.cssText = 'display: flex; align-items: center; gap: 8px; padding: 8px; background: #f9fafb; border: 1px solid #d1d5db; border-radius: 4px;';

      const progressBar = document.createElement('div');
      progressBar.style.cssText = `
        flex: 1;
        height: 6px;
        background: #e5e7eb;
        border-radius: 3px;
        overflow: hidden;
      `;
      const progressFill = document.createElement('div');
      progressFill.style.cssText = `
        height: 100%;
        background: #10b981;
        width: ${progress}%;
        transition: width 0.3s;
      `;
      progressBar.appendChild(progressFill);
      progressSection.appendChild(progressBar);

      const progressText = document.createElement('span');
      progressText.style.cssText = 'font-size: 11px; color: #6b7280; white-space: nowrap;';
      progressText.textContent = `${progress}% (${translated}/${this.translationPairs.length})`;
      progressSection.appendChild(progressText);

      wrapper.appendChild(progressSection);
    }

    // Apply translations button
    const applySection = document.createElement('div');
    applySection.style.cssText = 'display: flex; gap: 8px; padding-top: 12px; border-top: 1px solid #d1d5db;';

    const applyBtn = document.createElement('button');
    applyBtn.textContent = 'Apply Translations';
    applyBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      display: ${this.translationPairs.length === 0 ? 'none' : 'block'};
    `;
    applyBtn.addEventListener('click', () => this.applyTranslations());
    applySection.appendChild(applyBtn);

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
    applySection.appendChild(closeBtn);

    wrapper.appendChild(applySection);
    container.appendChild(wrapper);
  }

  /**
   * Extract translatable text from template
   */
  extractTranslatableText() {
    console.log('Extracting translatable text...');
    
    this.translationPairs = [
      { source: 'Report Title', target: '' },
      { source: 'Date Range', target: '' },
      { source: 'Total Sales', target: '' },
      { source: 'Region', target: '' },
      { source: 'Product Category', target: '' },
      { source: 'Year to Date', target: '' },
      { source: 'Previous Year', target: '' },
      { source: 'Variance', target: '' },
      { source: 'Percentage Change', target: '' },
      { source: 'Summary Statistics', target: '' }
    ];

    const container = document.querySelector('.translation-manager-wrapper').parentElement;
    if (container) this.render(container);
  }

  /**
   * Import XLIFF file
   */
  importXliff() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlf,.xliff';
    input.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
          const xliffContent = event.target.result;
          this.parseXliff(xliffContent);
          const container = document.querySelector('.translation-manager-wrapper').parentElement;
          if (container) this.render(container);
        };
        reader.readAsText(file);
      }
    });
    input.click();
  }

  /**
   * Parse XLIFF content
   */
  parseXliff(xliffContent) {
    try {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xliffContent, 'text/xml');
      const transUnits = xmlDoc.querySelectorAll('trans-unit');
      
      this.translationPairs = [];
      transUnits.forEach(unit => {
        const source = unit.querySelector('source')?.textContent || '';
        const target = unit.querySelector('target')?.textContent || '';
        if (source) {
          this.translationPairs.push({ source, target });
        }
      });
    } catch (error) {
      console.error('Error parsing XLIFF:', error);
      alert('Error parsing XLIFF file');
    }
  }

  /**
   * Export to XLIFF format
   */
  exportXliff() {
    let xliffContent = `<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file source-language="${this.sourceLanguage}" target-language="${this.targetLanguage}">
    <body>`;

    this.translationPairs.forEach((pair, index) => {
      xliffContent += `
      <trans-unit id="str_${index}">
        <source>${this.escapeXml(pair.source)}</source>
        <target>${this.escapeXml(pair.target || '')}</target>
      </trans-unit>`;
    });

    xliffContent += `
    </body>
  </file>
</xliff>`;

    const blob = new Blob([xliffContent], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `translations_${this.sourceLanguage}_${this.targetLanguage}.xlf`;
    link.click();
  }

  /**
   * Escape XML special characters
   */
  escapeXml(str) {
    return str.replace(/[<>&'"]/g, (c) => {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
        default: return c;
      }
    });
  }

  /**
   * Apply translations to template
   */
  async applyTranslations() {
    try {
      let translatedContent = '<!-- Localized Template -->\n';
      
      this.translationPairs.forEach(pair => {
        if (pair.target) {
          translatedContent += `<!-- ${pair.source} -->\n`;
          translatedContent += `<?${pair.target}?>\n`;
        }
      });

      await this.wordApi.insertText(translatedContent);
      alert('Translations applied successfully');
    } catch (error) {
      console.error('Error applying translations:', error);
      alert('Error applying translations: ' + error.message);
    }
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default TranslationManager;
export { TranslationManager };
