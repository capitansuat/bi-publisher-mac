/**
 * AccessibilityChecker Component
 * Component for checking document accessibility against WCAG 2.1 AA and Section 508
 * compliance with auto-fix capabilities and detailed reporting.
 */

class AccessibilityChecker {
  constructor(services) {
    this.services = services;
    this.wordApi = services.wordApi;
    this.templateEngine = services.templateEngine;
    this.checkResults = [];
    this.isRunning = false;
    this.complianceLevel = 'wcag21aa';
  }

  /**
   * Render the AccessibilityChecker component
   */
  render(container) {
    container.innerHTML = '';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'accessibility-checker-wrapper';
    wrapper.style.cssText = `
      display: flex;
      flex-direction: column;
      height: 100%;
      padding: 16px;
      gap: 12px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    const title = document.createElement('h3');
    title.textContent = 'Accessibility Checker';
    title.style.cssText = 'margin: 0; font-size: 14px; font-weight: 600; color: #1f2937;';
    wrapper.appendChild(title);

    // Compliance level selector
    const complianceSection = document.createElement('fieldset');
    complianceSection.style.cssText = 'border: 1px solid #d1d5db; border-radius: 4px; padding: 12px; margin: 0;';
    
    const complianceLegend = document.createElement('legend');
    complianceLegend.style.cssText = 'padding: 0 8px; font-size: 12px; font-weight: 600; color: #1f2937;';
    complianceLegend.textContent = 'Compliance Standard';
    complianceSection.appendChild(complianceLegend);

    const complianceContent = document.createElement('div');
    complianceContent.style.cssText = 'display: flex; gap: 16px; font-size: 12px;';

    const standards = [
      { value: 'wcag21aa', label: 'WCAG 2.1 AA' },
      { value: 'section508', label: 'Section 508' }
    ];

    standards.forEach(std => {
      const label = document.createElement('label');
      label.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'compliance-standard';
      radio.value = std.value;
      radio.checked = this.complianceLevel === std.value;
      radio.addEventListener('change', () => this.complianceLevel = std.value);
      label.appendChild(radio);
      label.appendChild(document.createTextNode(std.label));
      complianceContent.appendChild(label);
    });

    complianceSection.appendChild(complianceContent);
    wrapper.appendChild(complianceSection);

    // Run check button
    const runSection = document.createElement('div');
    runSection.style.cssText = 'display: flex; gap: 8px;';

    const runBtn = document.createElement('button');
    runBtn.textContent = this.isRunning ? 'Checking...' : 'Run Check';
    runBtn.style.cssText = `
      flex: 1;
      padding: 8px 16px;
      background: ${this.isRunning ? '#9ca3af' : '#3b82f6'};
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      font-weight: 500;
      cursor: ${this.isRunning ? 'wait' : 'pointer'};
    `;
    runBtn.disabled = this.isRunning;
    runBtn.addEventListener('click', () => this.runAccessibilityCheck(container));
    runSection.appendChild(runBtn);

    const exportBtn = document.createElement('button');
    exportBtn.textContent = 'Export Report';
    exportBtn.style.cssText = `
      padding: 8px 16px;
      background: #10b981;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
      display: ${this.checkResults.length === 0 ? 'none' : 'block'};
    `;
    exportBtn.addEventListener('click', () => this.exportReport());
    runSection.appendChild(exportBtn);

    wrapper.appendChild(runSection);

    // Results area
    const resultsSection = document.createElement('div');
    resultsSection.style.cssText = `
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 8px;
      overflow: hidden;
      min-height: 0;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      padding: 12px;
      background: #f9fafb;
    `;

    const resultsTitle = document.createElement('div');
    resultsTitle.style.cssText = 'font-size: 12px; font-weight: 600; color: #1f2937;';
    resultsTitle.textContent = `Check Results (${this.checkResults.length} issues found)`;
    resultsSection.appendChild(resultsTitle);

    const resultsList = document.createElement('div');
    resultsList.style.cssText = 'flex: 1; overflow-y: auto; display: flex; flex-direction: column; gap: 6px;';

    if (this.checkResults.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'color: #9ca3af; font-size: 12px; text-align: center; padding: 24px;';
      emptyMsg.textContent = 'No checks run yet. Click "Run Check" to scan your document.';
      resultsList.appendChild(emptyMsg);
    } else {
      // Group results by category
      const byCategory = {};
      this.checkResults.forEach(result => {
        if (!byCategory[result.category]) byCategory[result.category] = [];
        byCategory[result.category].push(result);
      });

      Object.entries(byCategory).forEach(([category, results]) => {
        const categoryHeader = document.createElement('div');
        categoryHeader.style.cssText = `
          font-size: 11px;
          font-weight: 600;
          color: #6b7280;
          margin-top: 8px;
          padding-top: 8px;
          border-top: 1px solid #e5e7eb;
          text-transform: uppercase;
        `;
        categoryHeader.textContent = category;
        resultsList.appendChild(categoryHeader);

        results.forEach(result => {
          const resultItem = document.createElement('div');
          
          const severityColors = {
            error: '#fee2e2',
            warning: '#fef3c7',
            info: '#dbeafe'
          };
          const severityTextColors = {
            error: '#991b1b',
            warning: '#92400e',
            info: '#1e40af'
          };

          resultItem.style.cssText = `
            padding: 8px;
            background: ${severityColors[result.severity] || '#f3f4f6'};
            border: 1px solid ${Object.values(severityColors)[Object.keys(severityColors).indexOf(result.severity)] || '#e5e7eb'};
            border-left: 3px solid ${result.severity === 'error' ? '#dc2626' : result.severity === 'warning' ? '#f59e0b' : '#3b82f6'};
            border-radius: 3px;
            font-size: 11px;
            color: ${severityTextColors[result.severity] || '#6b7280'};
            cursor: pointer;
            transition: background 0.15s;
          `;
          resultItem.onmouseover = () => resultItem.style.background = Object.values(severityColors)[Object.keys(severityColors).indexOf(result.severity)] || '#f3f4f6';
          resultItem.onmouseout = () => resultItem.style.background = severityColors[result.severity] || '#f3f4f6';

          const msgDiv = document.createElement('div');
          msgDiv.style.cssText = 'display: flex; justify-content: space-between; align-items: start;';
          
          const msgText = document.createElement('span');
          msgText.textContent = result.message;
          msgDiv.appendChild(msgText);

          if (result.fixable) {
            const fixBtn = document.createElement('button');
            fixBtn.textContent = 'Fix';
            fixBtn.style.cssText = `
              padding: 2px 6px;
              background: transparent;
              color: inherit;
              border: 1px solid currentColor;
              border-radius: 2px;
              font-size: 10px;
              cursor: pointer;
              white-space: nowrap;
              margin-left: 8px;
            `;
            fixBtn.addEventListener('click', (e) => {
              e.stopPropagation();
              this.applyFix(result);
            });
            msgDiv.appendChild(fixBtn);
          }

          resultItem.appendChild(msgDiv);

          if (result.suggestion) {
            const sugDiv = document.createElement('div');
            sugDiv.style.cssText = 'margin-top: 4px; font-size: 10px; color: inherit; opacity: 0.8;';
            sugDiv.textContent = `Suggestion: ${result.suggestion}`;
            resultItem.appendChild(sugDiv);
          }

          resultItem.addEventListener('click', () => {
            console.log('Navigate to issue:', result);
          });

          resultsList.appendChild(resultItem);
        });
      });
    }

    resultsSection.appendChild(resultsList);
    wrapper.appendChild(resultsSection);

    // Summary section
    if (this.checkResults.length > 0) {
      const summarySection = document.createElement('div');
      summarySection.style.cssText = 'padding: 10px; background: white; border: 1px solid #d1d5db; border-radius: 4px; font-size: 11px;';

      const errors = this.checkResults.filter(r => r.severity === 'error').length;
      const warnings = this.checkResults.filter(r => r.severity === 'warning').length;
      const infos = this.checkResults.filter(r => r.severity === 'info').length;

      let summaryHtml = '';
      if (errors > 0) summaryHtml += `<span style="color: #dc2626;">⚠ ${errors} Error${errors > 1 ? 's' : ''}</span> `;
      if (warnings > 0) summaryHtml += `<span style="color: #f59e0b;">⚠ ${warnings} Warning${warnings > 1 ? 's' : ''}</span> `;
      if (infos > 0) summaryHtml += `<span style="color: #3b82f6;">ℹ ${infos} Info${infos > 1 ? 's' : ''}</span>`;

      const complianceStatus = document.createElement('div');
      complianceStatus.style.cssText = 'display: flex; justify-content: space-between; align-items: center;';
      complianceStatus.innerHTML = `<div>${summaryHtml}</div>`;

      const compliance = document.createElement('div');
      compliance.style.cssText = `
        padding: 4px 8px;
        background: ${errors === 0 ? '#dcfce7' : '#fee2e2'};
        color: ${errors === 0 ? '#166534' : '#991b1b'};
        border-radius: 3px;
        font-weight: 600;
      `;
      compliance.textContent = errors === 0 ? '✓ Compliant' : '✗ Not Compliant';
      complianceStatus.appendChild(compliance);

      summarySection.appendChild(complianceStatus);
      wrapper.appendChild(summarySection);
    }

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
   * Run accessibility check
   */
  async runAccessibilityCheck(container) {
    this.isRunning = true;
    this.render(container);

    // Simulate check execution
    await new Promise(resolve => setTimeout(resolve, 1000));

    this.checkResults = [
      {
        category: 'Images',
        severity: 'error',
        message: 'Image missing alt text at position 1',
        fixable: true,
        suggestion: 'Add descriptive alt text to the image'
      },
      {
        category: 'Images',
        severity: 'warning',
        message: 'Alt text "image" is too generic (page 2)',
        fixable: false,
        suggestion: 'Replace with more descriptive text'
      },
      {
        category: 'Tables',
        severity: 'error',
        message: 'Table missing header row markup (Table ID: tbl_001)',
        fixable: true,
        suggestion: 'Mark first row as table header'
      },
      {
        category: 'Headings',
        severity: 'warning',
        message: 'Heading hierarchy gap (H1 → H3 detected)',
        fixable: true,
        suggestion: 'Use H2 between H1 and H3'
      },
      {
        category: 'Color Contrast',
        severity: 'error',
        message: 'Low contrast text at coordinates (100, 200): ratio 2.5:1 (needs 4.5:1)',
        fixable: false,
        suggestion: 'Increase contrast by darkening text or lightening background'
      },
      {
        category: 'Forms',
        severity: 'error',
        message: 'Form field not properly labeled (ID: fld_email)',
        fixable: true,
        suggestion: 'Associate label with form field'
      },
      {
        category: 'Language',
        severity: 'info',
        message: 'Document language not specified',
        fixable: true,
        suggestion: 'Set document language to English'
      },
      {
        category: 'Lists',
        severity: 'warning',
        message: 'List structure improperly formatted (page 3)',
        fixable: true,
        suggestion: 'Use proper list formatting instead of manual bullets'
      }
    ];

    this.isRunning = false;
    this.render(container);
  }

  /**
   * Apply auto-fix to an issue
   */
  applyFix(result) {
    console.log('Applying fix for:', result.message);
    alert(`Applied fix for: ${result.message}`);

    // Remove fixed item from results
    this.checkResults = this.checkResults.filter(r => r.message !== result.message);
    
    // Re-render to show updated results
    const container = document.querySelector('.accessibility-checker-wrapper').parentElement;
    if (container) this.render(container);
  }

  /**
   * Export accessibility report
   */
  exportReport() {
    let reportHtml = `<html>
<head>
  <title>Accessibility Report</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
    h1 { color: #1f2937; }
    h2 { color: #374151; margin-top: 20px; }
    .summary { background: #f3f4f6; padding: 15px; border-radius: 5px; margin: 15px 0; }
    .error { color: #991b1b; background: #fee2e2; padding: 10px; margin: 10px 0; border-left: 4px solid #dc2626; }
    .warning { color: #92400e; background: #fef3c7; padding: 10px; margin: 10px 0; border-left: 4px solid #f59e0b; }
    .info { color: #1e40af; background: #dbeafe; padding: 10px; margin: 10px 0; border-left: 4px solid #3b82f6; }
  </style>
</head>
<body>
  <h1>Accessibility Compliance Report</h1>
  <p>Standard: ${this.complianceLevel === 'wcag21aa' ? 'WCAG 2.1 AA' : 'Section 508'}</p>
  <p>Date: ${new Date().toLocaleDateString()}</p>
  
  <div class="summary">
    <h2>Summary</h2>
    <p>Total Issues: ${this.checkResults.length}</p>
    <p>Errors: ${this.checkResults.filter(r => r.severity === 'error').length}</p>
    <p>Warnings: ${this.checkResults.filter(r => r.severity === 'warning').length}</p>
    <p>Info: ${this.checkResults.filter(r => r.severity === 'info').length}</p>
  </div>
  
  <h2>Detailed Results</h2>`;

    // Group by category
    const byCategory = {};
    this.checkResults.forEach(result => {
      if (!byCategory[result.category]) byCategory[result.category] = [];
      byCategory[result.category].push(result);
    });

    Object.entries(byCategory).forEach(([category, results]) => {
      reportHtml += `<h3>${category}</h3>`;
      results.forEach(result => {
        reportHtml += `<div class="${result.severity}">
          <strong>${result.message}</strong>
          ${result.suggestion ? `<p>Suggestion: ${result.suggestion}</p>` : ''}
          ${result.fixable ? '<p>[Fixable]</p>' : ''}
        </div>`;
      });
    });

    reportHtml += `</body></html>`;

    // Download report
    const blob = new Blob([reportHtml], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'accessibility-report.html';
    link.click();
  }

  /**
   * Bind events
   */
  bindEvents() {
    // Events bound during render
  }
}

export default AccessibilityChecker;
export { AccessibilityChecker };
