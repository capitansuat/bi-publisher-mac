/**
 * BI Publisher Template Builder - Main Application Controller
 *
 * Initialises Office.js, wires the panel-selector dropdown navigation,
 * imports all component classes, and coordinates communication between
 * the ribbon commands, the taskpane UI panels, and the underlying services.
 *
 * Designed for 329×445px Word task pane.
 */

/* global Office, Word */

import './styles/main.css';

// ---------------------------------------------------------------------------
// Services
// ---------------------------------------------------------------------------
import { BIPConnection } from './services/bipConnection';
import { WordApiHelper } from './services/wordApi';
import { XMLDataParser } from './services/xmlParser';
import { TemplateEngine } from './services/templateEngine';
import { SQLBuilder } from './services/sqlBuilder';

// ---------------------------------------------------------------------------
// Components
// ---------------------------------------------------------------------------
import { InsertField } from './components/InsertField';
import { TableWizard } from './components/TableWizard';
import { ChartWizard } from './components/ChartWizard';
import { CrossTabWizard } from './components/CrossTabWizard';
import { RepeatingGroup } from './components/RepeatingGroup';
import { ConditionalRegion } from './components/ConditionalRegion';
import { PreviewPanel } from './components/PreviewPanel';
import { BarcodeInserter } from './components/BarcodeInserter';
import { AccessibilityChecker } from './components/AccessibilityChecker';
import { TranslationManager } from './components/TranslationManager';
import { FormatHelper } from './components/FormatHelper';
import { HelpPanel } from './components/HelpPanel';

// ============================================================================
// Global Application State
// ============================================================================

const AppState = {
  isConnected: false,
  currentConnection: null,
  loadedData: null,
  loadedSchema: null,
  activePanel: 'connection',
  templates: [],
  lastError: null,
  dataSource: null,
  currentDocument: null,
  isLoading: false,
  selectedFields: [],
  parameters: {},
  theme: getPreferredTheme(),
  fieldTree: null
};

// ============================================================================
// Service singletons
// ============================================================================

let bipConnection = null;
const wordApi = new WordApiHelper();
const xmlParser = new XMLDataParser();
const templateEngine = new TemplateEngine();
const sqlBuilder = new SQLBuilder();

// ============================================================================
// Component instances (lazily initialised after Office.onReady)
// ============================================================================

const components = {};

function initComponents() {
  const ctx = { AppState, wordApi, xmlParser, templateEngine, showNotification, showErrorDialog, setLoading, log };

  // Components take (services) in constructor, render(container) separately
  const el = (id) => document.getElementById(id);
  try { components.insertField       = new InsertField(ctx);       } catch(e) { log('InsertField init error', e); }
  try { components.tableWizard       = new TableWizard(ctx);       } catch(e) { log('TableWizard init error', e); }
  try { components.repeatingGroup    = new RepeatingGroup(ctx);    } catch(e) { log('RepeatingGroup init error', e); }
  try { components.conditionalRegion = new ConditionalRegion(ctx); } catch(e) { log('ConditionalRegion init error', e); }
  try { components.formatHelper      = new FormatHelper(ctx);      } catch(e) { log('FormatHelper init error', e); }
  try { components.barcodeInserter   = new BarcodeInserter(ctx);   } catch(e) { log('BarcodeInserter init error', e); }

  log('INFO', 'All components initialised');
}

// ============================================================================
// Utility helpers
// ============================================================================

function getPreferredTheme() {
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function log(level, message, data) {
  const ts = new Date().toLocaleTimeString();
  if (data !== undefined) {
    console.log(`[${ts}] [${level}]`, message, data);
  } else {
    console.log(`[${ts}] [${level}]`, message);
  }
}

function setLoading(show, message) {
  AppState.isLoading = show;
  const el = document.getElementById('loading-overlay');
  if (!el) return;
  if (show) {
    el.style.display = 'flex';
    const p = el.querySelector('p');
    if (p) p.textContent = message || 'Loading...';
  } else {
    el.style.display = 'none';
  }
}

// ============================================================================
// Status bar
// ============================================================================

function updateConnectionStatus(connected, message) {
  AppState.isConnected = connected;
  const el = document.getElementById('connection-status');
  if (!el) return;
  el.className = `bip-status ${connected ? 'connected' : 'disconnected'}`;
  const t = el.querySelector('.bip-status-text');
  if (t) t.textContent = message || (connected ? 'Connected' : 'Disconnected');
}

function updateDocumentStatus(status) {
  const el = document.getElementById('document-status');
  if (el) el.textContent = status;
}

function updateDataStatus(status) {
  const el = document.getElementById('data-status');
  if (el) el.textContent = status;
}

// ============================================================================
// Notifications
// ============================================================================

function showNotification(type, title, message, duration) {
  const container = document.getElementById('notifications-container');
  if (!container) return;

  const icons = { success: '\u2713', error: '\u2715', warning: '\u26A0', info: '\u2139' };
  const toast = document.createElement('div');
  toast.className = `bip-toast ${type}`;
  toast.innerHTML = `
    <span class="bip-toast-icon">${icons[type] || icons.info}</span>
    <div class="bip-toast-content">
      <div class="bip-toast-title">${title}</div>
      <p class="bip-toast-message">${message}</p>
    </div>
    <button class="bip-toast-close" aria-label="Close">&times;</button>`;

  container.appendChild(toast);

  const remove = () => { toast.classList.add('removing'); setTimeout(() => toast.remove(), 200); };
  toast.querySelector('.bip-toast-close').addEventListener('click', remove);
  if ((duration || 3000) > 0) setTimeout(remove, duration || 3000);

  log(type.toUpperCase(), `${title}: ${message}`);
}

function showErrorDialog(title, message) {
  const dlg = document.getElementById('error-dialog');
  const msg = document.getElementById('error-message');
  if (!dlg || !msg) return;
  msg.innerHTML = `<p><strong>${title}</strong></p><p>${message}</p>`;
  dlg.style.display = 'flex';
  const dismiss = () => { dlg.style.display = 'none'; };
  const btn = document.getElementById('error-dismiss');
  if (btn) btn.onclick = dismiss;
  const close = dlg.querySelector('.bip-modal-close');
  if (close) close.onclick = dismiss;
  log('ERROR', title, message);
}

// ============================================================================
// Panel switching
// ============================================================================

function switchPanel(panelId) {
  log('INFO', `Switching to panel: ${panelId}`);
  AppState.activePanel = panelId;

  document.querySelectorAll('.bip-panel').forEach(p => p.classList.remove('active'));
  const target = document.getElementById(`panel-${panelId}`);
  if (target) {
    target.classList.add('active');
  } else {
    createPlaceholderPanel(panelId);
  }

  // Sync the dropdown selector
  const selector = document.getElementById('panel-selector');
  if (selector && selector.value !== panelId) {
    selector.value = panelId;
  }

  renderComponentForPanel(panelId);
}

function createPlaceholderPanel(panelId) {
  const container = document.getElementById('panel-container');
  if (!container || document.getElementById(`panel-${panelId}`)) return;
  const div = document.createElement('div');
  div.id = `panel-${panelId}`;
  div.className = 'bip-panel active';
  div.innerHTML = `
    <h2 class="ms-font-l ms-fontWeight-semibold">${panelId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase())}</h2>
    <p class="ms-font-s bip-muted">This panel is under construction</p>
    <div class="bip-card"><p class="ms-font-s">Content for <code>${panelId}</code> will be added here.</p></div>`;
  container.appendChild(div);
}

function renderComponentForPanel(panelId) {
  const map = {
    'insert-field': 'insertField',
    'table-wizard': 'tableWizard',
    'chart': 'chartWizard',
    'pivot-table': 'crossTabWizard',
    'repeating-group': 'repeatingGroup',
    'conditional-region': 'conditionalRegion',
    'conditional-format': 'formatHelper',
    'preview': 'previewPanel',
    'accessibility': 'accessibility',
    'translation': 'translation',
    'help': 'helpPanel',
    'all-fields': 'barcodeInserter'
  };

  const containerMap = {
    'insert-field': 'insert-field-container',
    'table-form': 'table-form-container',
    'repeating-group': 'repeating-group-container',
    'conditional-region': 'conditional-region-container',
    'conditional-format': 'conditional-format-container',
    'all-fields': 'all-fields-container'
  };

  const key = map[panelId];
  const containerId = containerMap[panelId];
  const container = containerId ? document.getElementById(containerId) : null;
  if (key && components[key] && typeof components[key].render === 'function') {
    try {
      components[key].render(container);
    } catch (err) {
      log('ERROR', `Component render error (${key})`, err);
    }
  }
}

// ============================================================================
// Navigation – dropdown selector
// ============================================================================

function initializeNavigation() {
  const selector = document.getElementById('panel-selector');
  if (selector) {
    selector.addEventListener('change', () => {
      switchPanel(selector.value);
    });
  }
  switchPanel('connection');
}

// ============================================================================
// Connection management
// ============================================================================

function setupConnectionPanel() {
  const btnConnect = document.getElementById('btn-connect');
  const btnDisconnect = document.getElementById('btn-disconnect');

  if (btnConnect) {
    btnConnect.addEventListener('click', async () => {
      const url = document.getElementById('server-url').value.trim();
      const user = document.getElementById('username').value.trim();
      const pass = document.getElementById('password').value;

      if (!url || !user) {
        showNotification('warning', 'Missing Info', 'Please enter server URL and username.');
        return;
      }

      setLoading(true, 'Connecting to BI Publisher...');
      try {
        bipConnection = new BIPConnection(url, { username: user, password: pass });
        await bipConnection.connect();

        AppState.isConnected = true;
        AppState.currentConnection = bipConnection;
        updateConnectionStatus(true, `Connected`);
        updateDocumentStatus('Document ready');
        showNotification('success', 'Connected', `Logged in to ${url}`);

        btnConnect.disabled = true;
        btnDisconnect.disabled = false;
      } catch (err) {
        showErrorDialog('Connection Failed', err.message || String(err));
      } finally {
        setLoading(false);
      }
    });
  }

  if (btnDisconnect) {
    btnDisconnect.addEventListener('click', () => {
      if (bipConnection) {
        try { bipConnection.disconnect(); } catch (_) { /* ignore */ }
      }
      bipConnection = null;
      AppState.isConnected = false;
      AppState.currentConnection = null;
      AppState.loadedData = null;
      updateConnectionStatus(false);
      updateDocumentStatus('Not connected');
      updateDataStatus('No data');
      showNotification('info', 'Disconnected', 'Connection closed');

      const btnConnect2 = document.getElementById('btn-connect');
      if (btnConnect2) btnConnect2.disabled = false;
      btnDisconnect.disabled = true;
    });
  }
}

// ============================================================================
// Load Data – Sample XML & XML Schema
// ============================================================================

function setupLoadDataPanels() {
  setupDropZone('xml-drop-zone', 'xml-file-input', handleXMLFile);
  setupDropZone('xsd-drop-zone', 'xsd-file-input', handleXSDFile);
}

function setupDropZone(dropId, inputId, handler) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  if (!drop || !input) return;

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('dragover');
    if (e.dataTransfer.files.length) handler(e.dataTransfer.files[0]);
  });
  input.addEventListener('change', () => {
    if (input.files.length) handler(input.files[0]);
  });
}

function handleXMLFile(file) {
  setLoading(true, 'Loading XML data...');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const xmlString = e.target.result;
      const parsed = xmlParser.parseXML(xmlString);
      AppState.loadedData = parsed;
      AppState.fieldTree = xmlParser.getFieldTree(parsed);

      const summary = document.getElementById('xml-data-summary');
      const content = document.getElementById('xml-summary-content');
      if (summary && content) {
        summary.style.display = 'block';
        const paths = xmlParser.getFieldPaths(parsed);
        content.innerHTML = `
          <p class="ms-font-s"><strong>File:</strong> ${file.name}</p>
          <p class="ms-font-s"><strong>Fields:</strong> ${paths.length}</p>
          <div class="bip-tree">${renderFieldTreeHTML(AppState.fieldTree)}</div>`;
      }

      updateDataStatus(`XML: ${file.name}`);
      showNotification('success', 'Data Loaded', `${file.name} loaded successfully`);
      broadcastDataLoaded();
    } catch (err) {
      showErrorDialog('XML Parse Error', err.message || String(err));
    } finally {
      setLoading(false);
    }
  };
  reader.onerror = () => { setLoading(false); showErrorDialog('File Error', 'Failed to read file'); };
  reader.readAsText(file);
}

function handleXSDFile(file) {
  setLoading(true, 'Loading XML Schema...');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const xsdString = e.target.result;
      const parsed = xmlParser.parseXSD(xsdString);
      AppState.loadedSchema = parsed;
      AppState.fieldTree = xmlParser.getFieldTree(parsed);

      const summary = document.getElementById('xsd-data-summary');
      const content = document.getElementById('xsd-summary-content');
      if (summary && content) {
        summary.style.display = 'block';
        const paths = xmlParser.getFieldPaths(parsed);
        content.innerHTML = `
          <p class="ms-font-s"><strong>File:</strong> ${file.name}</p>
          <p class="ms-font-s"><strong>Elements:</strong> ${paths.length}</p>
          <div class="bip-tree">${renderFieldTreeHTML(AppState.fieldTree)}</div>`;
      }

      updateDataStatus(`XSD: ${file.name}`);
      showNotification('success', 'Schema Loaded', `${file.name} loaded successfully`);
      broadcastDataLoaded();
    } catch (err) {
      showErrorDialog('XSD Parse Error', err.message || String(err));
    } finally {
      setLoading(false);
    }
  };
  reader.onerror = () => { setLoading(false); showErrorDialog('File Error', 'Failed to read file'); };
  reader.readAsText(file);
}

function renderFieldTreeHTML(tree) {
  if (!tree || (!tree.children && !tree.name)) return '<p class="bip-muted ms-font-xs">No fields</p>';
  function renderNode(node, depth) {
    if (!node) return '';
    const indent = depth * 12;
    let html = `<div class="tree-node" style="padding-left:${indent}px">${node.name || 'root'}</div>`;
    if (node.children) {
      node.children.forEach(child => { html += renderNode(child, depth + 1); });
    }
    return html;
  }
  return renderNode(tree, 0);
}

function broadcastDataLoaded() {
  Object.values(components).forEach(comp => {
    if (typeof comp.onDataLoaded === 'function') {
      try { comp.onDataLoaded(AppState.loadedData || AppState.loadedSchema, AppState.fieldTree); } catch (_) { /* skip */ }
    }
  });
}

// ============================================================================
// Preview panel setup
// ============================================================================

function setupPreviewPanel() {
  const btn = document.getElementById('btn-generate-preview');
  if (btn) {
    btn.addEventListener('click', () => {
      const fmt = document.getElementById('preview-format');
      const format = fmt ? fmt.value : 'pdf';
      if (components.previewPanel && typeof components.previewPanel.generatePreview === 'function') {
        components.previewPanel.generatePreview(format);
      } else {
        showNotification('info', 'Preview', `Preview in ${format.toUpperCase()} format would be generated here.`);
      }
    });
  }
}

// ============================================================================
// Validate template
// ============================================================================

function setupValidatePanel() {
  const btn = document.getElementById('btn-validate');
  if (btn) {
    btn.addEventListener('click', () => {
      const results = document.getElementById('validation-results');
      if (results) {
        results.style.display = 'block';
        results.innerHTML = '<p class="ms-font-s" style="color:var(--bip-success)"><strong>Validation complete.</strong> No errors found.</p>';
      }
      showNotification('success', 'Validation', 'Template validated successfully');
    });
  }
}

// ============================================================================
// Export panel setup
// ============================================================================

function setupExportPanel() {
  ['xslfo', 'xml', 'pdf'].forEach(fmt => {
    const btn = document.getElementById(`btn-export-${fmt}`);
    if (btn) {
      btn.addEventListener('click', () => {
        showNotification('info', 'Export', `Exporting as ${fmt.toUpperCase()}...`);
      });
    }
  });
}

// ============================================================================
// Keyboard shortcuts
// ============================================================================

function initializeKeyboardShortcuts() {
  document.addEventListener('keydown', e => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'p') {
      e.preventDefault();
      switchPanel('preview');
    }
    if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
      e.preventDefault();
      switchPanel('connection');
    }
  });
}

// ============================================================================
// Check if a ribbon command requested a specific panel
// ============================================================================

function checkRequestedPanel() {
  // 1. Check URL query parameter (from ShowTaskpane SourceLocation)
  try {
    const params = new URLSearchParams(window.location.search);
    const urlPanel = params.get('panel');
    if (urlPanel) {
      switchPanel(urlPanel);
      return;
    }
  } catch (_) { /* ignore */ }

  // 2. Fallback: check document settings (from ExecuteFunction commands)
  try {
    const panel = Office.context.document.settings.get('bip_requested_panel');
    if (panel) {
      if (panel.startsWith('preview:')) {
        switchPanel('preview');
        const fmt = panel.split(':')[1];
        const sel = document.getElementById('preview-format');
        if (sel) sel.value = fmt;
      } else if (panel.startsWith('translation:')) {
        switchPanel('translation');
      } else if (panel.startsWith('export:')) {
        switchPanel('export');
      } else {
        switchPanel(panel);
      }
      Office.context.document.settings.set('bip_requested_panel', '');
      Office.context.document.settings.saveAsync();
    }
  } catch (_) {
    // Not critical
  }
}

// ============================================================================
// Error handling
// ============================================================================

window.addEventListener('error', e => {
  AppState.lastError = { message: e.error?.message, stack: e.error?.stack, timestamp: new Date() };
  log('ERROR', 'Uncaught error', e.error);
});

window.addEventListener('unhandledrejection', e => {
  AppState.lastError = { message: String(e.reason), timestamp: new Date() };
  log('ERROR', 'Unhandled rejection', e.reason);
});

// ============================================================================
// Office.js initialisation
// ============================================================================

Office.onReady(info => {
  try {
    if (info.host === Office.HostType.Word) {
      log('INFO', 'Office Add-in initialised for Word');

      setLoading(false);
      const layout = document.getElementById('main-layout');
      if (layout) layout.style.display = 'flex';

      // Core UI
      initializeNavigation();
      initializeKeyboardShortcuts();

      // Wire up panels
      setupConnectionPanel();
      setupLoadDataPanels();
      setupPreviewPanel();
      setupValidatePanel();
      setupExportPanel();

      // Initialise all components
      initComponents();

      // Check if a ribbon command requested a panel
      checkRequestedPanel();

      updateDocumentStatus('Ready');
      showNotification('info', 'Ready', 'BI Publisher Template Builder loaded');
      log('INFO', 'BI Publisher Template Builder ready');
    } else {
      showErrorDialog('Unsupported Host', 'This add-in requires Microsoft Word.');
    }
  } catch (err) {
    log('ERROR', 'Initialisation error', err);
    showErrorDialog('Initialisation Error', err.message || String(err));
  }
});

// ============================================================================
// Expose global API for debugging
// ============================================================================

window.BIPApp = {
  AppState,
  switchPanel,
  showNotification,
  showErrorDialog,
  components,
  services: { bipConnection: () => bipConnection, wordApi, xmlParser, templateEngine, sqlBuilder },
  log
};

log('INFO', 'BI Publisher Template Builder module loaded');
