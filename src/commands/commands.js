/**
 * BI Publisher Template Builder - Ribbon Command Handlers
 *
 * Maps every ribbon button/menu item (manifest.xml FunctionName values)
 * to a handler that stores the requested panel in document settings,
 * so the taskpane can switch to it on open.
 */

/* global Office */

Office.onReady(() => {
  console.log('[Commands] BI Publisher commands module ready');
});

// ---------------------------------------------------------------------------
// Helper: persist which panel should be shown, then signal completion
// ---------------------------------------------------------------------------
function requestPanel(panelId, event) {
  try {
    Office.context.document.settings.set('bip_requested_panel', panelId);
    Office.context.document.settings.saveAsync(() => {
      console.log(`[Commands] Requested panel: ${panelId}`);
      if (event) event.completed();
    });
  } catch (err) {
    console.error('[Commands] requestPanel error', err);
    if (event) event.completed();
  }
}

// ---------------------------------------------------------------------------
// Online group
// ---------------------------------------------------------------------------
function cmdOpen(event) { requestPanel('open-report', event); }
function cmdUpload(event) { requestPanel('upload-template', event); }

// ---------------------------------------------------------------------------
// Load Data group
// ---------------------------------------------------------------------------
function cmdSampleXML(event) { requestPanel('sample-xml', event); }
function cmdXMLSchema(event) { requestPanel('xml-schema', event); }

// ---------------------------------------------------------------------------
// Insert group
// ---------------------------------------------------------------------------
function cmdTableWizard(event) { requestPanel('table-wizard', event); }
function cmdPivotTable(event) { requestPanel('pivot-table', event); }
function cmdChart(event) { requestPanel('chart', event); }
function cmdField(event) { requestPanel('insert-field', event); }
function cmdTableForm(event) { requestPanel('table-form', event); }
function cmdRepeatingGroup(event) { requestPanel('repeating-group', event); }
function cmdCondFormat(event) { requestPanel('conditional-format', event); }
function cmdCondRegion(event) { requestPanel('conditional-region', event); }
function cmdAllFields(event) { requestPanel('all-fields', event); }

// ---------------------------------------------------------------------------
// Preview group
// ---------------------------------------------------------------------------
function cmdPreviewPDF(event) { requestPanel('preview:pdf', event); }
function cmdPreviewHTML(event) { requestPanel('preview:html', event); }
function cmdPreviewExcel(event) { requestPanel('preview:excel', event); }
function cmdPreviewExcel2000(event) { requestPanel('preview:excel2000', event); }
function cmdPreviewRTF(event) { requestPanel('preview:rtf', event); }
function cmdPreviewPPTX(event) { requestPanel('preview:pptx', event); }
function cmdPreviewDOCX(event) { requestPanel('preview:docx', event); }

// ---------------------------------------------------------------------------
// Tools group
// ---------------------------------------------------------------------------
function cmdFieldBrowser(event) { requestPanel('field-browser', event); }
function cmdValidate(event) { requestPanel('validate-template', event); }
function cmdAccessibility(event) { requestPanel('accessibility', event); }
function cmdExtractText(event) { requestPanel('translation:extract', event); }
function cmdPreviewTranslation(event) { requestPanel('translation:preview', event); }
function cmdLocalize(event) { requestPanel('translation:localize', event); }
function cmdExportXslFo(event) { requestPanel('export:xslfo', event); }
function cmdExportXml(event) { requestPanel('export:xml', event); }
function cmdExportPdf(event) { requestPanel('export:pdf', event); }

// ---------------------------------------------------------------------------
// Options group
// ---------------------------------------------------------------------------
function cmdOptions(event) { requestPanel('options', event); }
function cmdHelp(event) { requestPanel('help', event); }
function cmdAbout(event) { requestPanel('about', event); }

// ---------------------------------------------------------------------------
// Register all functions with Office (names must match manifest FunctionName)
// ---------------------------------------------------------------------------
Office.actions.associate("cmdOpen", cmdOpen);
Office.actions.associate("cmdUpload", cmdUpload);
Office.actions.associate("cmdSampleXML", cmdSampleXML);
Office.actions.associate("cmdXMLSchema", cmdXMLSchema);
Office.actions.associate("cmdTableWizard", cmdTableWizard);
Office.actions.associate("cmdPivotTable", cmdPivotTable);
Office.actions.associate("cmdChart", cmdChart);
Office.actions.associate("cmdField", cmdField);
Office.actions.associate("cmdTableForm", cmdTableForm);
Office.actions.associate("cmdRepeatingGroup", cmdRepeatingGroup);
Office.actions.associate("cmdCondFormat", cmdCondFormat);
Office.actions.associate("cmdCondRegion", cmdCondRegion);
Office.actions.associate("cmdAllFields", cmdAllFields);
Office.actions.associate("cmdPreviewPDF", cmdPreviewPDF);
Office.actions.associate("cmdPreviewHTML", cmdPreviewHTML);
Office.actions.associate("cmdPreviewExcel", cmdPreviewExcel);
Office.actions.associate("cmdPreviewExcel2000", cmdPreviewExcel2000);
Office.actions.associate("cmdPreviewRTF", cmdPreviewRTF);
Office.actions.associate("cmdPreviewPPTX", cmdPreviewPPTX);
Office.actions.associate("cmdPreviewDOCX", cmdPreviewDOCX);
Office.actions.associate("cmdFieldBrowser", cmdFieldBrowser);
Office.actions.associate("cmdValidate", cmdValidate);
Office.actions.associate("cmdAccessibility", cmdAccessibility);
Office.actions.associate("cmdExtractText", cmdExtractText);
Office.actions.associate("cmdPreviewTranslation", cmdPreviewTranslation);
Office.actions.associate("cmdLocalize", cmdLocalize);
Office.actions.associate("cmdExportXslFo", cmdExportXslFo);
Office.actions.associate("cmdExportXml", cmdExportXml);
Office.actions.associate("cmdExportPdf", cmdExportPdf);
Office.actions.associate("cmdOptions", cmdOptions);
Office.actions.associate("cmdHelp", cmdHelp);
Office.actions.associate("cmdAbout", cmdAbout);
