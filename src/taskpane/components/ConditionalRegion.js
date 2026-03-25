/**
 * ConditionalRegion Component
 * If, If-Else, Choose/When/Otherwise for BI Publisher templates.
 * Inserts content controls into Word via Word.run().
 */

class ConditionalRegion {
  constructor(services) {
    this.services = services;
    this.mode = 'simple'; // 'simple' | 'ifelse' | 'choose'
    this.dataFields = [];
    // For simple/ifelse: single condition set
    this.conditions = [];
    // For choose: multiple when blocks, each with own conditions
    this.whenBlocks = [{ conditions: [], logic: 'and' }];
    this.conditionLogic = 'and';
    this._container = null;
  }

  render(container) {
    this._container = container;
    container.innerHTML = '';
    this.loadDataFields();

    const w = document.createElement('div');
    w.style.cssText = 'display:flex;flex-direction:column;gap:12px;padding:12px;font-family:-apple-system,BlinkMacSystemFont,sans-serif;font-size:13px;';

    // Mode selection
    w.appendChild(this._buildModeSelector());

    // Condition builder(s)
    if (this.mode === 'choose') {
      w.appendChild(this._buildChooseUI());
    } else {
      w.appendChild(this._buildConditionBuilder(this.conditions, this.conditionLogic, (conds, logic) => {
        this.conditions = conds;
        this.conditionLogic = logic;
        this._updatePreview();
      }));
    }

    // XSL Preview
    const previewBox = document.createElement('div');
    previewBox.id = 'cond-preview';
    previewBox.style.cssText = 'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:6px;font-family:monospace;font-size:11px;white-space:pre-wrap;max-height:200px;overflow:auto;';
    previewBox.textContent = this._generateExpression();
    w.appendChild(previewBox);

    // Action buttons
    w.appendChild(this._buildButtons());

    container.appendChild(w);
  }

  _buildModeSelector() {
    const fieldset = document.createElement('fieldset');
    fieldset.style.cssText = 'border:1px solid #555;border-radius:6px;padding:10px;margin:0;';
    const legend = document.createElement('legend');
    legend.textContent = 'Type';
    legend.style.cssText = 'padding:0 6px;font-weight:600;font-size:12px;';
    fieldset.appendChild(legend);

    const modes = [
      { value: 'simple', label: 'Simple If', desc: 'Show content when condition is true' },
      { value: 'ifelse', label: 'If / Else', desc: 'Show different content for true/false' },
      { value: 'choose', label: 'Choose / When / Otherwise', desc: 'Multiple conditions with fallback' }
    ];

    modes.forEach(m => {
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer;';
      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'condMode';
      radio.value = m.value;
      radio.checked = this.mode === m.value;
      radio.addEventListener('change', () => {
        this.mode = m.value;
        // Reset when blocks for choose mode
        if (m.value === 'choose' && this.whenBlocks.length === 0) {
          this.whenBlocks = [{ conditions: [], logic: 'and' }];
        }
        this.render(this._container);
      });
      label.appendChild(radio);
      const span = document.createElement('span');
      span.innerHTML = `<strong>${m.label}</strong><br><span style="font-size:11px;color:#999;">${m.desc}</span>`;
      label.appendChild(span);
      fieldset.appendChild(label);
    });

    return fieldset;
  }

  _buildConditionBuilder(conditions, logic, onChange) {
    const section = document.createElement('div');
    section.style.cssText = 'border:1px solid #555;border-radius:6px;padding:10px;';

    // Existing conditions
    conditions.forEach((cond, idx) => {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:4px;padding:4px 0;background:#2a2a2a;border-radius:4px;padding:6px;margin-bottom:4px;';
      row.innerHTML = `<span style="flex:1;font-size:12px;">${cond.field} ${cond.operator} ${cond.value || ''}</span>`;
      const removeBtn = document.createElement('button');
      removeBtn.textContent = 'X';
      removeBtn.style.cssText = 'background:#ef4444;color:white;border:none;border-radius:3px;padding:2px 8px;cursor:pointer;font-size:11px;';
      removeBtn.addEventListener('click', () => {
        conditions.splice(idx, 1);
        onChange(conditions, logic);
        this.render(this._container);
      });
      row.appendChild(removeBtn);
      section.appendChild(row);
    });

    // Add new condition row
    const addRow = document.createElement('div');
    addRow.style.cssText = 'display:flex;gap:4px;align-items:center;flex-wrap:wrap;margin-top:6px;';

    // Logic selector (AND/OR) - only show if conditions exist
    if (conditions.length > 0) {
      const logicSel = document.createElement('select');
      logicSel.style.cssText = 'padding:4px;border-radius:3px;border:1px solid #555;font-size:11px;background:#333;color:#fff;';
      logicSel.innerHTML = '<option value="and">AND</option><option value="or">OR</option>';
      logicSel.value = logic;
      logicSel.addEventListener('change', () => {
        onChange(conditions, logicSel.value);
        this._updatePreview();
      });
      addRow.appendChild(logicSel);
    }

    // Field select
    const fieldSel = document.createElement('select');
    fieldSel.style.cssText = 'flex:1;min-width:80px;padding:4px;border-radius:3px;border:1px solid #555;font-size:11px;background:#333;color:#fff;';
    fieldSel.innerHTML = '<option value="">Field</option>';
    this.dataFields.forEach(f => {
      fieldSel.innerHTML += `<option value="${f.name}">${f.name}</option>`;
    });
    addRow.appendChild(fieldSel);

    // Operator select
    const opSel = document.createElement('select');
    opSel.style.cssText = 'padding:4px;border-radius:3px;border:1px solid #555;font-size:11px;background:#333;color:#fff;';
    opSel.innerHTML = `
      <option value="=">=</option>
      <option value="!=">!=</option>
      <option value="&gt;">&gt;</option>
      <option value="&lt;">&lt;</option>
      <option value="&gt;=">&gt;=</option>
      <option value="&lt;=">&lt;=</option>
      <option value="contains">contains</option>
      <option value="starts-with">starts-with</option>
    `;
    addRow.appendChild(opSel);

    // Value input
    const valInput = document.createElement('input');
    valInput.type = 'text';
    valInput.placeholder = 'Value';
    valInput.style.cssText = 'flex:1;min-width:60px;padding:4px;border-radius:3px;border:1px solid #555;font-size:11px;background:#333;color:#fff;';
    addRow.appendChild(valInput);

    // Add button
    const addBtn = document.createElement('button');
    addBtn.textContent = 'Add';
    addBtn.style.cssText = 'padding:4px 10px;background:#3b82f6;color:white;border:none;border-radius:3px;cursor:pointer;font-size:11px;font-weight:600;';
    addBtn.addEventListener('click', () => {
      if (!fieldSel.value) return;
      conditions.push({
        field: fieldSel.value,
        operator: opSel.value,
        value: valInput.value
      });
      onChange(conditions, logic);
      this.render(this._container);
    });
    addRow.appendChild(addBtn);

    section.appendChild(addRow);
    return section;
  }

  _buildChooseUI() {
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'display:flex;flex-direction:column;gap:8px;';

    // Each "when" block
    this.whenBlocks.forEach((block, idx) => {
      const blockEl = document.createElement('fieldset');
      blockEl.style.cssText = 'border:1px solid #3b82f6;border-radius:6px;padding:10px;margin:0;';
      const legend = document.createElement('legend');
      legend.style.cssText = 'padding:0 6px;font-weight:600;font-size:12px;color:#3b82f6;display:flex;align-items:center;gap:8px;';
      legend.textContent = `When #${idx + 1}`;

      // Remove button (if more than 1 block)
      if (this.whenBlocks.length > 1) {
        const removeBlock = document.createElement('button');
        removeBlock.textContent = 'Remove';
        removeBlock.style.cssText = 'background:#ef4444;color:white;border:none;border-radius:3px;padding:2px 8px;cursor:pointer;font-size:10px;';
        removeBlock.addEventListener('click', () => {
          this.whenBlocks.splice(idx, 1);
          this.render(this._container);
        });
        legend.appendChild(removeBlock);
      }
      blockEl.appendChild(legend);

      blockEl.appendChild(this._buildConditionBuilder(block.conditions, block.logic, (conds, logic) => {
        block.conditions = conds;
        block.logic = logic;
        this._updatePreview();
      }));

      wrapper.appendChild(blockEl);
    });

    // "Add When" button
    const addWhenBtn = document.createElement('button');
    addWhenBtn.textContent = '+ Add When Block';
    addWhenBtn.style.cssText = 'padding:8px;background:#1e40af;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:600;';
    addWhenBtn.addEventListener('click', () => {
      this.whenBlocks.push({ conditions: [], logic: 'and' });
      this.render(this._container);
    });
    wrapper.appendChild(addWhenBtn);

    // Otherwise note
    const note = document.createElement('div');
    note.style.cssText = 'padding:8px;background:#374151;border-radius:4px;font-size:11px;color:#9ca3af;';
    note.innerHTML = '<strong style="color:#f59e0b;">Otherwise:</strong> Otomatik eklenir — hiçbir koşul doğru değilse gösterilecek varsayılan içerik.';
    wrapper.appendChild(note);

    return wrapper;
  }

  _buildButtons() {
    const area = document.createElement('div');
    area.style.cssText = 'display:flex;gap:8px;padding-top:8px;border-top:1px solid #444;';

    const wrapBtn = this._makeBtn('Wrap Selection', '#3b82f6', () => this._wrapSelection());
    const insertBtn = this._makeBtn('Insert Region', '#10b981', () => this._insertRegion());

    area.appendChild(wrapBtn);
    area.appendChild(insertBtn);
    return area;
  }

  _makeBtn(text, bg, onClick) {
    const btn = document.createElement('button');
    btn.textContent = text;
    btn.style.cssText = `flex:1;padding:10px;background:${bg};color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:600;`;
    btn.addEventListener('click', onClick);
    return btn;
  }

  _updatePreview() {
    const el = document.getElementById('cond-preview');
    if (el) el.textContent = this._generateExpression();
  }

  _buildCondExpr(conditions, logic) {
    if (!conditions || conditions.length === 0) return 'true()';
    const parts = conditions.map(c => {
      if (['contains', 'starts-with', 'ends-with'].includes(c.operator)) {
        return `${c.operator}(${c.field}, '${c.value}')`;
      }
      return `${c.field} ${c.operator} '${c.value}'`;
    });
    return parts.join(` ${(logic || 'and').toUpperCase()} `);
  }

  _generateExpression() {
    if (this.mode === 'simple') {
      const expr = this._buildCondExpr(this.conditions, this.conditionLogic);
      return `<?if:${expr}?>\n  [Content when true]\n<?end if?>`;
    }
    if (this.mode === 'ifelse') {
      const expr = this._buildCondExpr(this.conditions, this.conditionLogic);
      return `<?if:${expr}?>\n  [Content when TRUE]\n<?else?>\n  [Content when FALSE]\n<?end if?>`;
    }
    // choose mode
    let out = '<?choose?>\n';
    this.whenBlocks.forEach((block, i) => {
      const expr = this._buildCondExpr(block.conditions, block.logic);
      out += `<?when:${expr}?>\n  [When #${i + 1} content]\n<?end when?>\n`;
    });
    out += '<?otherwise?>\n  [Default content]\n<?end otherwise?>\n';
    out += '<?end choose?>';
    return out;
  }

  async _wrapSelection() {
    try {
      const expr = this._generateExpression();
      const lines = expr.split('\n');
      // opening tags = all lines before first content placeholder
      // closing tags = all lines after last content placeholder

      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();

        const selectedText = selection.text || '';

        if (this.mode === 'simple') {
          const condExpr = this._buildCondExpr(this.conditions, this.conditionLogic);
          const wrapped = `<?if:${condExpr}?>${selectedText}<?end if?>`;
          selection.insertText(wrapped, 'Replace');
        } else if (this.mode === 'ifelse') {
          const condExpr = this._buildCondExpr(this.conditions, this.conditionLogic);
          const wrapped = `<?if:${condExpr}?>${selectedText}<?else?>[Else content]<?end if?>`;
          selection.insertText(wrapped, 'Replace');
        } else {
          // choose mode - wrap with first when condition
          let wrapped = '<?choose?>';
          this.whenBlocks.forEach((block, i) => {
            const bExpr = this._buildCondExpr(block.conditions, block.logic);
            if (i === 0) {
              wrapped += `<?when:${bExpr}?>${selectedText}<?end when?>`;
            } else {
              wrapped += `<?when:${bExpr}?>[When #${i + 1} content]<?end when?>`;
            }
          });
          wrapped += '<?otherwise?>[Default]<?end otherwise?>';
          wrapped += '<?end choose?>';
          selection.insertText(wrapped, 'Replace');
        }

        await context.sync();
      });
    } catch (err) {
      alert('Error: ' + (err.message || err));
    }
  }

  async _insertRegion() {
    try {
      const fullExpr = this._generateExpression();

      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(fullExpr, 'After');
        await context.sync();
      });
    } catch (err) {
      alert('Error: ' + (err.message || err));
    }
  }

  loadDataFields() {
    this.dataFields = [];
    let tree = this.services.AppState ? this.services.AppState.fieldTree : null;

    // Fallback: restore from localStorage if AppState is empty
    if (!tree) {
      try {
        const stored = localStorage.getItem('bip_fieldTree');
        if (stored) {
          tree = JSON.parse(stored);
          if (this.services.AppState) this.services.AppState.fieldTree = tree;
        }
      } catch (_) {}
    }
    if (!tree) return;

    const collectLeaves = (node) => {
      if (!node) return;
      if (node.children && node.children.length > 0) {
        node.children.forEach(child => collectLeaves(child));
      } else {
        let type = 'text';
        if (node.sampleValue) {
          if (!isNaN(Number(node.sampleValue))) type = 'number';
          else if (/^\d{4}-\d{2}-\d{2}/.test(node.sampleValue)) type = 'date';
        }
        this.dataFields.push({ name: node.name, xpath: node.xpath || node.name, type });
      }
    };

    if (Array.isArray(tree)) tree.forEach(n => collectLeaves(n));
    else collectLeaves(tree);
  }

  onDataLoaded(data, fieldTree) {
    this.loadDataFields();
  }
}

export default ConditionalRegion;
export { ConditionalRegion };
