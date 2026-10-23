/* ============================================================
 * app.js  —— 上传 / 渲染 / 编辑 / 导出 调度
 * ============================================================ */
(function () {
  const FIELDS = window.XF_FIELDS;
  const records = []; // 内存数组

  // DOM 引用
  const fileInput = document.getElementById('fileInput');
  const dropzone = document.getElementById('dropzone');
  const messageArea = document.getElementById('messageArea');
  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');
  const emptyTip = document.getElementById('emptyTip');
  const statText = document.getElementById('statText');
  const modalMask = document.getElementById('modalMask');
  const modalBody = document.getElementById('modalBody');
  const modalIndex = document.getElementById('modalIndex');

  let editingIdx = -1;

  function init() {
    closeModal();
    buildHead();
    renderTable();
    bindEvents();
  }

  /* ---------- 事件绑定 ---------- */
  function bindEvents() {
    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
    ['dragenter', 'dragover'].forEach((ev) =>
      dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((ev) =>
      dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
    dropzone.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));
    dropzone.addEventListener('click', () => fileInput.click());

    document.getElementById('exportBtn').addEventListener('click', doExport);
    document.getElementById('deleteSelBtn').addEventListener('click', deleteSelected);
    document.getElementById('selectAllBtn').addEventListener('click', toggleSelectAll);
    document.getElementById('modalClose').addEventListener('click', closeModal);
    document.getElementById('modalCancel').addEventListener('click', closeModal);
    document.getElementById('modalSave').addEventListener('click', saveModal);
    modalMask.addEventListener('click', (e) => { if (e.target === modalMask) closeModal(); });
  }

  async function handleFiles(fileList) {
    const files = Array.from(fileList).filter((f) => /\.docx?$/i.test(f.name));
    const skipped = Array.from(fileList).filter((f) => !/\.docx?$/i.test(f.name));
    const errors = [];
    skipped.forEach((f) => errors.push({ file: f.name, reason: '非 Word(.docx/.doc) 文件，已跳过' }));

    for (const file of files) {
      try {
        const rec = await window.XFParser.parseFile(file);
        records.push(rec);
      } catch (err) {
        errors.push({ file: file.name, reason: err && err.message ? err.message : String(err) });
      }
    }
    fileInput.value = ''; // 允许重复选同名
    renderTable();
    renderMessages(errors);
  }

  /* ---------- 消息区 ---------- */
  function renderMessages(errors) {
    let errHtml = '';
    let warnHtml = '';

    if (errors && errors.length) {
      const lis = errors.map((e) => `<li>${escapeHtml(e.file)}：${escapeHtml(e.reason)}</li>`).join('');
      errHtml = `<div class="msg-group error"><div class="msg-title">解析失败（已跳过）</div><ul>${lis}</ul></div>`;
    }

    // 汇总所有记录的字段警告
    records.forEach((r, i) => {
      if (r._warnings && r._warnings.length) {
        warnHtml += `<li>第 ${i + 1} 条【${escapeHtml(r._file || '')}】：${r._warnings.map(escapeHtml).join('；')}</li>`;
      }
    });
    const warnBlock = warnHtml
      ? `<div class="msg-group warn"><div class="msg-title">字段提示（数据已入表，请核对）</div><ul>${warnHtml}</ul></div>`
      : '';

    messageArea.innerHTML = errHtml + warnBlock;
    messageArea.hidden = !(errHtml || warnBlock);
  }

  /* ---------- 表头 ---------- */
  function buildHead() {
    let tr = '<tr><th class="col-action sticky-col">选择 / #</th>';
    FIELDS.forEach((f) => {
      tr += `<th>${escapeHtml(f.header)}<span class="en">${escapeHtml(f.field)}</span></th>`;
    });
    tr += '<th class="col-action">操作</th></tr>';
    tableHead.innerHTML = tr;
  }

  /* ---------- 渲染表格 ---------- */
  function renderTable() {
    tableBody.innerHTML = '';
    records.forEach((r, i) => {
      const tr = document.createElement('tr');
      let html = '';
      html += `<td class="col-check sticky-col">
        <label class="rowcheck"><input type="checkbox" data-idx="${i}" ${r._checked ? 'checked' : ''}> ${i + 1}</label></td>`;
      FIELDS.forEach((f) => {
        const val = r[f.field] || '';
        const cls = f.type === 'textarea' ? 'textarea-cell' : '';
        const warn = r._warnings && r._warnings.some((w) => w.startsWith(f.header + '字段为空'));
        html += `<td class="${cls}" title="${escapeHtml(val)}"${warn ? ' style="background:#fffdf0"' : ''}>${escapeHtml(val)}</td>`;
      });
      html += `<td class="col-action">
        <button class="icon-btn edit" data-idx="${i}">编辑</button>
        <button class="icon-btn del" data-idx="${i}">删除</button></td>`;
      tr.innerHTML = html;
      tableBody.appendChild(tr);
    });

    // 绑定行内事件
    tableBody.querySelectorAll('input[type=checkbox]').forEach((cb) =>
      cb.addEventListener('change', (e) => { records[+e.target.dataset.idx]._checked = e.target.checked; updateStat(); }));
    tableBody.querySelectorAll('.icon-btn.edit').forEach((b) =>
      b.addEventListener('click', (e) => openModal(+e.currentTarget.dataset.idx)));
    tableBody.querySelectorAll('.icon-btn.del').forEach((b) =>
      b.addEventListener('click', (e) => { records.splice(+e.currentTarget.dataset.idx, 1); renderTable(); renderMessages(); }));

    emptyTip.style.display = records.length ? 'none' : 'block';
    updateStat();
  }

  function updateStat() {
    const sel = records.filter((r) => r._checked).length;
    statText.textContent = `共 ${records.length} 条 · 已选 ${sel} 条`;
  }

  /* ---------- 全选 / 删除 ---------- */
  function toggleSelectAll() {
    const allChecked = records.length && records.every((r) => r._checked);
    records.forEach((r) => { r._checked = !allChecked; });
    renderTable();
  }
  function deleteSelected() {
    const keep = records.filter((r) => !r._checked);
    if (keep.length === records.length) { alert('没有选中任何记录'); return; }
    if (!confirm(`确认删除选中的 ${records.length - keep.length} 条记录？`)) return;
    records.length = 0;
    keep.forEach((r) => records.push(r));
    renderTable();
    renderMessages();
  }

  /* ---------- 编辑模态框 ---------- */
  function openModal(i) {
    if (!Number.isInteger(i) || i < 0 || i >= records.length || !records[i]) {
      closeModal();
      return;
    }

    editingIdx = i;
    const r = records[i];
    modalIndex.textContent = `#${i + 1} · ${r._file || ''}`;
    let html = '';
    FIELDS.forEach((f) => {
      const val = r[f.field] || '';
      const warn = r._warnings && r._warnings.filter((w) => w.includes(f.header));
      let input;
      if (f.type === 'select') {
        const all = Array.from(new Set([val, ...f.options].filter((x) => x !== undefined)));
        input = `<select data-field="${escapeHtml(f.field)}">` +
          all.map((o) => `<option value="${escapeHtml(o)}"${o === val ? ' selected' : ''}>${escapeHtml(o) || '（空）'}</option>`).join('') +
          `</select>`;
      } else if (f.type === 'textarea') {
        input = `<textarea data-field="${escapeHtml(f.field)}">${escapeHtml(val)}</textarea>`;
      } else {
        input = `<input type="text" data-field="${escapeHtml(f.field)}" value="${escapeHtml(val)}">`;
      }
      html += `<div class="form-row">
        <label>${escapeHtml(f.header)}<span class="en">${escapeHtml(f.field)}</span></label>
        <div class="field">${input}${f.blank ? '<div class="hint">导出时留空</div>' : ''}
        ${warn && warn.length ? `<div class="hint">${warn.map(escapeHtml).join('；')}</div>` : ''}</div>
      </div>`;
    });
    modalBody.innerHTML = html;
    modalMask.hidden = false;
    modalMask.style.display = 'flex';
    modalMask.setAttribute('aria-hidden', 'false');
  }

  function saveModal() {
    if (!Number.isInteger(editingIdx) || editingIdx < 0 || !records[editingIdx]) {
      closeModal();
      return;
    }
    const r = records[editingIdx];
    modalBody.querySelectorAll('[data-field]').forEach((el) => {
      r[el.dataset.field] = el.value;
    });
    // 重新计算派生字段（文档标题）
    r['fd_doc_subject'] = (r['fd_letter_no'] || '') + (r['fd_contact_name'] || '');
    recomputeWarnings(r);
    closeModal();
    renderTable();
    renderMessages();
  }

  function recomputeWarnings(r) {
    const warnings = [];
    FIELDS.forEach((f) => {
      if (f.optional || f.blank) return;
      if (!r[f.field]) warnings.push(f.header + '字段为空');
    });
    if (r['fd_content_type']) {
      const opts = FIELDS.find((f) => f.field === 'fd_content_type').options;
      if (!opts.includes(r['fd_content_type'])) warnings.push('内容分类字段不在可选范围，请手动修改：' + r['fd_content_type']);
    }
    ['fd_letter_date', 'fd_deadline_date'].forEach((df) => {
      if (r[df] && !/^\d{4}-\d{2}-\d{2}$/.test(r[df])) warnings.push((df === 'fd_letter_date' ? '信访日期' : '初件限办日期') + '格式异常：' + r[df]);
    });
    r._warnings = warnings;
  }

  function closeModal() {
    modalMask.hidden = true;
    modalMask.style.display = 'none';
    modalMask.setAttribute('aria-hidden', 'true');
    modalBody.innerHTML = '';
    editingIdx = -1;
  }

  /* ---------- 导出 ---------- */
  async function doExport() {
    const sel = records.filter((r) => r._checked);
    if (!sel.length) { alert('请先勾选要导出的记录'); return; }
    const btn = document.getElementById('exportBtn');
    btn.disabled = true; btn.textContent = '导出中…';
    try {
      const result = await window.XFExporter.exportRecords(sel);
      const name = (result && result.fileName) || result;
      alert('已导出 ' + sel.length + ' 条记录：' + name + '\n文件格式与模板完全一致，可直接上传省一体化平台。');
    } catch (e) {
      alert('导出失败：' + (e && e.message ? e.message : e));
    } finally {
      btn.disabled = false; btn.textContent = '导出 Excel';
    }
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
