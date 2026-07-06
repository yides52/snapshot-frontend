// =========================================================
// Client Setup Page — logic (redesigned)
// =========================================================
 
const db = window.supabase.createClient(window.SUPABASE_URL, window.SUPABASE_ANON_KEY);
 
const params = new URLSearchParams(window.location.search);
const slug = params.get('slug');
 
let currentClient = null;
let currentTemplate = null;
let currentTokens = [];
let currentReports = [];
let currentMappings = {};
let chatHistory = [];
let excelPreviews = {};  // report_id -> parsed sheet data
 
// DOM
const clientTitle = document.getElementById('clientTitle');
const emilyClientName = document.getElementById('emilyClientName');
const tplIcon = document.getElementById('tplIcon');
const templateStatus = document.getElementById('templateStatus');
const templateDrop = document.getElementById('templateDrop');
const templateInput = document.getElementById('templateInput');
const reportsList = document.getElementById('reportsList');
const addReportBtn = document.getElementById('addReportBtn');
const reportInput = document.getElementById('reportInput');
const tokensGrid = document.getElementById('tokensGrid');
 
// =========================================================
// LOAD
// =========================================================
async function loadClient() {
  if (!slug) { clientTitle.textContent = 'Missing client slug'; return; }
  const { data: client, error } = await db.from('clients').select('*').eq('slug', slug).single();
  if (error || !client) { clientTitle.textContent = 'Client not found'; return; }
  currentClient = client;
  clientTitle.textContent = client.name;
  emilyClientName.textContent = `${client.name}'s assistant`;
  document.title = `${client.name} — Emily's Snapshot Studio`;
  await loadExistingData();
  greetUser();
}
 
async function loadExistingData() {
  const { data: templates } = await db.from('client_templates')
    .select('*').eq('client_id', currentClient.id)
    .order('uploaded_at', { ascending: false }).limit(1);
 
  if (templates && templates.length > 0) {
    currentTemplate = templates[0];
    currentTokens = (currentTemplate.tokens || []).filter(t => t.token);
    const pngMarker = (currentTemplate.tokens || []).find(t => t._png_path);
    if (pngMarker) currentTemplate.png_path = pngMarker._png_path;
    if (!currentTemplate.png_path && currentTemplate.file_path) {
      currentTemplate.png_path = currentTemplate.file_path.replace(/\.docx$/i, '.preview.png');
    }
    showTemplateStatus();
    loadTemplatePreview();
  }
 
  const { data: reports } = await db.from('client_reports')
    .select('*').eq('client_id', currentClient.id)
    .order('uploaded_at', { ascending: true });
  currentReports = reports || [];
  renderReports();
 
  const { data: mappings } = await db.from('client_mappings')
    .select('*').eq('client_id', currentClient.id);
  currentMappings = {};
  if (mappings) mappings.forEach(m => currentMappings[m.token] = m);
 
  renderTokens();
}
 
// =========================================================
// TEMPLATE
// =========================================================
templateDrop.addEventListener('click', () => templateInput.click());
templateInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleTemplate(e.target.files[0]);
});
 
async function handleTemplate(file) {
  if (!file.name.endsWith('.docx')) { alert('Please upload a .docx file'); return; }
  const { tokens, rawText } = await parseDocx(file);
  const path = `${currentClient.id}/templates/${Date.now()}_${file.name}`;
  const { error: upErr } = await db.storage.from('client-files').upload(path, file);
  if (upErr) { alert('Upload failed: ' + upErr.message); return; }
 
  const tokenObjs = tokens.map(t => ({ token: t, label: t.replace(/_/g, ' '), is_manual: false }));
  const { data: newTemplate, error: dbErr } = await db.from('client_templates')
    .insert({ client_id: currentClient.id, file_path: path, tokens: tokenObjs })
    .select().single();
  if (dbErr) { alert('DB error: ' + dbErr.message); return; }
 
  currentTemplate = newTemplate;
  currentTemplate.raw_text = rawText;
  currentTokens = tokenObjs;
  showTemplateStatus();
  renderTokens();
 
  emilySay(`Uploading and generating preview…`);
  try {
    const res = await fetch(`${window.RENDER_API_URL}/render-template`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template_path: path })
    });
    const data = await res.json();
    if (data.png_path) {
      currentTemplate.png_path = data.png_path;
      await db.from('client_templates')
        .update({ tokens: [...tokenObjs, { _png_path: data.png_path }] })
        .eq('id', newTemplate.id);
      loadTemplatePreview();
    }
  } catch (e) { console.warn('Preview render failed:', e); }
 
  if (tokens.length === 0) {
    emilySay(`Uploaded! Ask me to propose the tokens.`);
  } else {
    emilySay(`Detected <strong>${tokens.length} tokens</strong>. Upload reports next if you haven't.`);
  }
}
 
async function parseDocx(file) {
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file('word/document.xml').async('string');
  const rawText = _cleanDocxText(docXml);
  const matches = rawText.match(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g) || [];
  const unique = [...new Set(matches.map(m => m.replace(/[\{\}\s]/g, '')))];
  return { tokens: unique, rawText };
}
 
function _cleanDocxText(xml) {
  let cleaned = xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '');
  const paragraphs = cleaned.split(/<w:p[\s>]/);
  const lines = [];
  for (const p of paragraphs) {
    const runs = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const text = runs.map(r => r.replace(/<w:t[^>]*>/, '').replace(/<\/w:t>/, '')).join('').trim();
    if (text) lines.push(text);
  }
  return lines.join('\n');
}
 
function showTemplateStatus() {
  const name = currentTemplate.file_path.split('/').pop().replace(/^\d+_/, '');
  templateStatus.innerHTML = `
    <div class="template-uploaded">
      <div style="min-width:0;flex:1;">
        <div class="filename">✓ ${name}</div>
        <div class="filemeta">${currentTokens.length} tokens</div>
      </div>
      <button class="replace-link" onclick="replaceTemplate()">Replace</button>
    </div>
  `;
  tplIcon.style.display = 'flex';
}
window.replaceTemplate = () => templateInput.click();
 
async function loadTemplatePreview() {
  if (!currentTemplate?.png_path) return;
  try {
    const { data, error } = await db.storage
      .from('client-files')
      .createSignedUrl(currentTemplate.png_path, 3600);
    if (error) throw error;
    document.getElementById('templateModalImg').src = data.signedUrl;
  } catch (e) { console.warn('Preview load failed:', e); }
}
 
// Template preview modal
tplIcon.addEventListener('click', () => {
  document.getElementById('tplModal').style.display = 'flex';
});
document.getElementById('tplModalClose').addEventListener('click', () => {
  document.getElementById('tplModal').style.display = 'none';
});
document.getElementById('tplModal').addEventListener('click', (e) => {
  if (e.target.id === 'tplModal') e.currentTarget.style.display = 'none';
});
 
// =========================================================
// REPORTS
// =========================================================
addReportBtn.addEventListener('click', () => reportInput.click());
reportInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleReport(e.target.files[0]);
});
 
async function handleReport(file) {
  const reportName = prompt(`Name this report:`, file.name.replace(/\.[^.]+$/, ''));
  if (!reportName) return;
  const path = `${currentClient.id}/samples/${Date.now()}_${file.name}`;
  const { error: upErr } = await db.storage.from('client-files').upload(path, file);
  if (upErr) { alert('Upload failed: ' + upErr.message); return; }
  const { data: newRow, error: dbErr } = await db.from('client_reports')
    .insert({ client_id: currentClient.id, report_name: reportName, sample_path: path, columns: [] })
    .select().single();
  if (dbErr) { alert('DB error: ' + dbErr.message); return; }
  currentReports.push(newRow);
  renderReports();
}
 
function renderReports() {
  reportsList.innerHTML = '';
  currentReports.forEach(r => {
    const chip = document.createElement('div');
    chip.className = 'report-chip';
    chip.innerHTML = `<span>📊 ${r.report_name}</span> <button class="chip-x" data-id="${r.id}">×</button>`;
    reportsList.appendChild(chip);
  });
  reportsList.querySelectorAll('.chip-x').forEach(x => {
    x.addEventListener('click', async (e) => {
      const id = e.target.dataset.id;
      if (!confirm('Remove this report?')) return;
      await db.from('client_reports').delete().eq('id', id);
      currentReports = currentReports.filter(r => r.id !== id);
      renderReports();
    });
  });
}
 
// =========================================================
// TOKENS
// =========================================================
function renderTokens() {
  tokensGrid.innerHTML = '';
  if (!currentTokens || currentTokens.length === 0) {
    tokensGrid.innerHTML = `
      <div class="tokens-empty">
        <div class="empty-icon">✨</div>
        <div class="empty-text">No tokens yet</div>
        <div class="empty-sub">Ask Emily to propose them → click her bubble bottom-right</div>
      </div>
    `;
    return;
  }
  currentTokens.forEach(t => {
    const mapping = currentMappings[t.token];
    const card = document.createElement('div');
    card.className = 'token-card';
    const mapText = mapping
      ? (mapping.operation === 'manual' ? 'Manual entry' : `${_reportName(mapping.report_id) || '?'} · ${mapping.operation}`)
      : 'Unmapped';
    const mapClass = mapping ? 'tc-mapping' : 'tc-mapping unmapped';
    card.innerHTML = `
      <div>
        <div class="tc-label">${t.label || t.token}</div>
        <div class="tc-name">{{${t.token}}}</div>
        ${t.note ? `<div class="tc-note">${t.note}</div>` : ''}
      </div>
      <div class="${mapClass}">${mapText}</div>
    `;
    card.addEventListener('click', () => openTokenModal(t));
    tokensGrid.appendChild(card);
  });
 
  // Add card
  const addCard = document.createElement('div');
  addCard.className = 'token-card add-card';
  addCard.innerHTML = `<div><div class="plus">+</div><div>Add token</div></div>`;
  addCard.addEventListener('click', () => {
    const name = prompt('Token name (snake_case, no spaces):');
    if (!name) return;
    const clean = name.toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (currentTokens.find(t => t.token === clean)) { alert('Token already exists.'); return; }
    addTokenLocal(clean);
  });
  tokensGrid.appendChild(addCard);
}
 
function _reportName(reportId) {
  const r = currentReports.find(r => r.id === reportId);
  return r?.report_name;
}
 
async function addTokenLocal(name) {
  const updated = [...currentTokens, { token: name, label: name.replace(/_/g, ' '), is_manual: false }];
  const marker = (currentTemplate.tokens || []).find(t => t._png_path);
  const toSave = [...updated];
  if (marker) toSave.push(marker);
  await db.from('client_templates').update({ tokens: toSave }).eq('id', currentTemplate.id);
  currentTokens = updated;
  renderTokens();
}
 
// =========================================================
// TOKEN EDITOR MODAL
// =========================================================
const tokenModal = document.getElementById('tokenModal');
const tokenModalTitle = document.getElementById('tokenModalTitle');
const tokenModalNote = document.getElementById('tokenModalNote');
const tokenSource = document.getElementById('tokenSource');
const tokenOperation = document.getElementById('tokenOperation');
const tokenConfig = document.getElementById('tokenConfig');
const excelPreviewDiv = document.getElementById('excelPreview');
const excelPreviewTable = document.getElementById('excelPreviewTable');
let editingToken = null;
 
document.getElementById('tokenModalClose').addEventListener('click', closeTokenModal);
document.getElementById('tokenCancelBtn').addEventListener('click', closeTokenModal);
tokenModal.addEventListener('click', (e) => { if (e.target === tokenModal) closeTokenModal(); });
 
document.getElementById('tokenSaveBtn').addEventListener('click', saveTokenMapping);
document.getElementById('tokenDeleteBtn').addEventListener('click', deleteToken);
 
tokenSource.addEventListener('change', () => { refreshExcelPreview(); refreshConfigFields(); });
tokenOperation.addEventListener('change', () => { refreshConfigFields(); refreshExcelPreview(); });
 
function openTokenModal(token) {
  editingToken = token;
  const mapping = currentMappings[token.token] || {};
  tokenModalTitle.textContent = `{{${token.token}}}`;
  tokenModalNote.textContent = token.note || token.label || '';
 
  // Populate source dropdown
  tokenSource.innerHTML = '<option value="">— manual —</option>' +
    currentReports.map(r =>
      `<option value="${r.id}" ${mapping.report_id === r.id ? 'selected' : ''}>${r.report_name}</option>`
    ).join('');
 
  tokenOperation.value = mapping.operation || 'manual';
  refreshConfigFields(mapping.config);
  refreshExcelPreview();
 
  tokenModal.style.display = 'flex';
}
 
function closeTokenModal() {
  tokenModal.style.display = 'none';
  editingToken = null;
}
 
function refreshConfigFields(existing) {
  const op = tokenOperation.value;
  const cfg = existing || currentMappings[editingToken?.token]?.config || {};
  let html = '';
  if (op === 'cell') {
    html = `
      <label>Sheet <input type="text" id="cfg_sheet" value="${cfg.sheet || 'Sheet1'}"></label>
      <label>Column (A, B, C...) <input type="text" id="cfg_col" value="${cfg.col || 'A'}"></label>
      <label>Row (number) <input type="number" id="cfg_row" value="${cfg.row || 1}"></label>
    `;
  } else if (op === 'value_at_label') {
    html = `
      <label>Label column (name from header) <input type="text" id="cfg_label_col" value="${cfg.label_col || ''}" placeholder="e.g. Customer"></label>
      <label>Label to match <input type="text" id="cfg_label_value" value="${cfg.label_value || ''}" placeholder="e.g. Total"></label>
      <label>Value column <input type="text" id="cfg_value_col" value="${cfg.value_col || ''}" placeholder="e.g. Amount"></label>
    `;
  } else if (op === 'sum') {
    html = `<label>Column to sum <input type="text" id="cfg_column" value="${cfg.column || ''}" placeholder="e.g. Amount"></label>`;
  } else if (op === 'section_sum') {
    html = `
      <label>Column to sum <input type="text" id="cfg_column" value="${cfg.column || ''}" placeholder="e.g. Amount"></label>
      <label>Section starts at label <input type="text" id="cfg_section_start" value="${cfg.section_start || ''}"></label>
      <label>Section ends at label (or blank) <input type="text" id="cfg_section_end" value="${cfg.section_end || ''}"></label>
    `;
  } else if (op === 'count') {
    html = `<label>Column to count non-empty rows <input type="text" id="cfg_column" value="${cfg.column || ''}" placeholder="e.g. Customer"></label>`;
  }
  tokenConfig.innerHTML = html;
  tokenConfig.style.display = html ? 'block' : 'none';
}
 
function _readConfigFields() {
  const op = tokenOperation.value;
  const cfg = {};
  if (op === 'cell') {
    cfg.sheet = document.getElementById('cfg_sheet')?.value;
    cfg.col = document.getElementById('cfg_col')?.value;
    cfg.row = parseInt(document.getElementById('cfg_row')?.value) || 1;
  } else if (op === 'value_at_label') {
    cfg.label_col = document.getElementById('cfg_label_col')?.value;
    cfg.label_value = document.getElementById('cfg_label_value')?.value;
    cfg.value_col = document.getElementById('cfg_value_col')?.value;
  } else if (op === 'sum') {
    cfg.column = document.getElementById('cfg_column')?.value;
  } else if (op === 'section_sum') {
    cfg.column = document.getElementById('cfg_column')?.value;
    cfg.section_start = document.getElementById('cfg_section_start')?.value;
    cfg.section_end = document.getElementById('cfg_section_end')?.value;
  } else if (op === 'count') {
    cfg.column = document.getElementById('cfg_column')?.value;
  }
  return cfg;
}
 
async function saveTokenMapping() {
  if (!editingToken) return;
  const reportId = tokenSource.value || null;
  const operation = tokenOperation.value;
  const config = _readConfigFields();
  const row = {
    client_id: currentClient.id,
    token: editingToken.token,
    report_id: reportId,
    operation,
    config
  };
  const { data, error } = await db.from('client_mappings')
    .upsert(row, { onConflict: 'client_id,token' }).select().single();
  if (error) { alert('Save failed: ' + error.message); return; }
  currentMappings[editingToken.token] = data;
  closeTokenModal();
  renderTokens();
}
 
async function deleteToken() {
  if (!editingToken) return;
  if (!confirm(`Delete {{${editingToken.token}}}?`)) return;
  // Remove from template tokens
  const marker = (currentTemplate.tokens || []).find(t => t._png_path);
  const remaining = currentTokens.filter(t => t.token !== editingToken.token);
  const toSave = [...remaining];
  if (marker) toSave.push(marker);
  await db.from('client_templates').update({ tokens: toSave }).eq('id', currentTemplate.id);
  await db.from('client_mappings').delete()
    .eq('client_id', currentClient.id).eq('token', editingToken.token);
  currentTokens = remaining;
  delete currentMappings[editingToken.token];
  closeTokenModal();
  renderTokens();
}
 
async function refreshExcelPreview() {
  const reportId = tokenSource.value;
  if (!reportId) { excelPreviewDiv.style.display = 'none'; return; }
  const report = currentReports.find(r => r.id === reportId);
  if (!report) { excelPreviewDiv.style.display = 'none'; return; }
 
  excelPreviewDiv.style.display = 'block';
  excelPreviewTable.innerHTML = '<div style="color:rgba(254,243,199,0.5);font-size:12px;">Loading preview…</div>';
 
  // Cache excel parsing per report
  if (!excelPreviews[reportId]) {
    try {
      const res = await fetch(`${window.RENDER_API_URL}/parse-excel`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sample_path: report.sample_path })
      });
      const data = await res.json();
      excelPreviews[reportId] = data;
    } catch (e) {
      excelPreviewTable.innerHTML = `<div style="color:#fb7185;font-size:12px;">Couldn't load: ${e.message}</div>`;
      return;
    }
  }
 
  const parsed = excelPreviews[reportId];
  if (!parsed?.sheets || parsed.sheets.length === 0) {
    excelPreviewTable.innerHTML = '<div style="color:#fb7185;font-size:12px;">No data.</div>';
    return;
  }
 
  const sheet = parsed.sheets[0];
  const cols = sheet.columns || [];
  const rows = (sheet.first_rows || []).slice(1, 9);   // skip header row, show 8 rows
 
  // Determine highlight (column name from config)
  const op = tokenOperation.value;
  const cfg = _readConfigFields();
  let highlightCol = cfg.column || cfg.value_col;
  const highlightColIdx = cols.findIndex(c => c === highlightCol);
 
  let html = '<table><thead><tr>';
  cols.forEach((c, i) => {
    html += `<th class="${i === highlightColIdx ? 'hl-cell' : ''}">${c || ''}</th>`;
  });
  html += '</tr></thead><tbody>';
  rows.forEach(r => {
    html += '<tr>';
    r.forEach((cell, i) => {
      html += `<td class="${i === highlightColIdx ? 'hl-cell' : ''}">${cell || ''}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  excelPreviewTable.innerHTML = html;
}
 
// =========================================================
// EMILY CHAT
// =========================================================
const emilyBubble = document.getElementById('emilyBubble');
const emilyPanel = document.getElementById('emilyPanel');
const emilyMessages = document.getElementById('emilyMessages');
const emilyInput = document.getElementById('emilyInput');
const emilySend = document.getElementById('emilySend');
const emilyClose = document.getElementById('emilyClose');
const emilyExpand = document.getElementById('emilyExpand');
const emilyReset = document.getElementById('emilyReset');
 
emilyBubble.addEventListener('click', () => {
  emilyPanel.classList.add('open');
  emilyBubble.style.display = 'none';
  emilyInput.focus();
});
emilyClose.addEventListener('click', () => {
  emilyPanel.classList.remove('open');
  emilyBubble.style.display = 'block';
});
emilyExpand.addEventListener('click', () => emilyPanel.classList.toggle('expanded'));
emilyReset.addEventListener('click', async () => {
  if (!confirm('Start fresh chat?')) return;
  await db.from('client_chat_messages').delete().eq('client_id', currentClient.id);
  emilyMessages.innerHTML = '';
  chatHistory = [];
  greetUser();
});
emilySend.addEventListener('click', sendEmilyMessage);
emilyInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendEmilyMessage(); }
});
 
function emilySay(html) {
  const msg = document.createElement('div');
  msg.className = 'msg emily';
  msg.innerHTML = html;
  emilyMessages.appendChild(msg);
  emilyMessages.scrollTop = emilyMessages.scrollHeight;
  const text = msg.textContent;
  chatHistory.push({ role: 'assistant', content: text });
  saveChatMessage('assistant', text);
}
 
function userSay(text) {
  const msg = document.createElement('div');
  msg.className = 'msg user';
  msg.textContent = text;
  emilyMessages.appendChild(msg);
  emilyMessages.scrollTop = emilyMessages.scrollHeight;
  chatHistory.push({ role: 'user', content: text });
  saveChatMessage('user', text);
}
 
async function saveChatMessage(role, content) {
  if (!currentClient) return;
  await db.from('client_chat_messages').insert({ client_id: currentClient.id, role, content });
}
 
async function loadChatHistory() {
  if (!currentClient) return false;
  const { data } = await db.from('client_chat_messages')
    .select('*').eq('client_id', currentClient.id)
    .order('created_at', { ascending: true }).limit(50);
  if (data && data.length > 0) {
    data.forEach(m => {
      const el = document.createElement('div');
      el.className = `msg ${m.role === 'user' ? 'user' : 'emily'}`;
      el.textContent = m.content;
      emilyMessages.appendChild(el);
      chatHistory.push({ role: m.role === 'user' ? 'user' : 'assistant', content: m.content });
    });
    emilyMessages.scrollTop = emilyMessages.scrollHeight;
    return true;
  }
  return false;
}
 
async function greetUser() {
  const had = await loadChatHistory();
  if (had) return;
  const first = `Hey! 👋 I'm Emily for <strong>${currentClient.name}</strong>. `;
  if (!currentTemplate) emilySay(first + `Upload your Word template — no tokens needed, we'll add them together.`);
  else if (currentTokens.length === 0) emilySay(first + `Want me to propose tokens for this template?`);
  else if (currentReports.length === 0) emilySay(first + `${currentTokens.length} tokens ready. Upload sample reports.`);
  else emilySay(first + `We're set up — ask me anything.`);
}
 
async function sendEmilyMessage() {
  const text = emilyInput.value.trim();
  if (!text) return;
  userSay(text);
  emilyInput.value = '';
 
  const typing = document.createElement('div');
  typing.className = 'msg emily typing';
  typing.textContent = 'Emily is thinking…';
  emilyMessages.appendChild(typing);
  emilyMessages.scrollTop = emilyMessages.scrollHeight;
 
  const coldTimer = setTimeout(() => {
    typing.textContent = 'Emily is waking up (cold start, ~50s)…';
  }, 4000);
 
  try {
    const shortHistory = chatHistory.slice(-8).map(m => ({
      role: m.role,
      content: (m.content || '').slice(0, 800)
    }));
 
    const body = {
      client_id: currentClient.id,
      template_id: currentTemplate?.id || null,
      client_name: currentClient.name,
      template_png_path: currentTemplate?.png_path || null,
      report_names: currentReports.map(r => r.report_name).slice(0, 10),
      sample_paths: currentReports.map(r => r.sample_path).slice(0, 5),
      current_tokens: currentTokens.map(t => t.token).slice(0, 30),
      current_mappings: Object.values(currentMappings).slice(0, 30).map(m => ({
        token: m.token, operation: m.operation
      })),
      history: shortHistory
    };
 
    const res = await fetch(`${window.RENDER_API_URL}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    clearTimeout(coldTimer);
    typing.remove();
 
    if (!res.ok) {
      const t = await res.text();
      emilySay(`⚠️ ${t.slice(0, 200)}`);
      return;
    }
    const data = await res.json();
    emilySay(renderMarkdownLite(data.reply || '(no reply)'));
    if (data.actions && data.actions.length > 0) {
      const s = data.actions.map(a => `• ${a.result}`).join('<br>');
      emilySay(`<em style="opacity:.7">✓ ${s}</em>`);
      await loadExistingData();
    }
  } catch (err) {
    clearTimeout(coldTimer);
    typing.remove();
    emilySay(`⚠️ Couldn't reach engine: ${err.message}`);
  }
}
 
function renderMarkdownLite(text) {
  const escaped = text
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return escaped
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(?!\s)([^*]+?)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\n/g, '<br>');
}
 
// =========================================================
// INIT
// =========================================================
loadClient();
 
