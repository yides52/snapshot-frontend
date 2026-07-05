
// =========================================================
// Client Setup Page — logic (Step 5: Emily wired to Anthropic)
// =========================================================
 
const db = window.supabase.createClient(window.SUPABASE_URL, window.SUPABASE_ANON_KEY);
 
const params = new URLSearchParams(window.location.search);
const slug = params.get('slug');
 
let currentClient = null;
let currentTemplate = null;
let currentTokens = [];
let currentReports = [];
let currentMappings = {};
let chatHistory = [];   // [{role: 'user'|'assistant', content: '...'}]
 
const clientTitle = document.getElementById('clientTitle');
const emilyClientName = document.getElementById('emilyClientName');
const templateDrop = document.getElementById('templateDrop');
const templateInput = document.getElementById('templateInput');
const templateStatus = document.getElementById('templateStatus');
const reportDrop = document.getElementById('reportDrop');
const reportInput = document.getElementById('reportInput');
const reportsList = document.getElementById('reportsList');
const tokensList = document.getElementById('tokensList');
 
// =========================================================
// LOAD CLIENT
// =========================================================
async function loadClient() {
  if (!slug) { clientTitle.textContent = 'Missing client slug'; return; }
 
  const { data: client, error } = await db
    .from('clients').select('*').eq('slug', slug).single();
 
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
    currentTokens = currentTemplate.tokens || [];
    // Re-fetch raw text from Storage so Emily has it
    try {
      const { data: fileBlob } = await db.storage
        .from('client-files').download(currentTemplate.file_path);
      if (fileBlob) {
        currentTemplate.raw_text = await extractRawTextFromBlob(fileBlob);
      }
    } catch (e) { console.warn('Template text load skipped:', e); }
    showTemplateStatus(currentTemplate);
    renderTokens();
  }
 
  const { data: reports } = await db.from('client_reports')
    .select('*').eq('client_id', currentClient.id)
    .order('uploaded_at', { ascending: true });
  currentReports = reports || [];
  renderReports();
 
  const { data: mappings } = await db.from('client_mappings')
    .select('*').eq('client_id', currentClient.id);
  if (mappings) {
    mappings.forEach(m => currentMappings[m.token] = m);
    renderTokens();
  }
}
 
// =========================================================
// TEMPLATE UPLOAD
// =========================================================
templateDrop.addEventListener('click', () => templateInput.click());
templateDrop.addEventListener('dragover', (e) => { e.preventDefault(); templateDrop.classList.add('dragover'); });
templateDrop.addEventListener('dragleave', () => templateDrop.classList.remove('dragover'));
templateDrop.addEventListener('drop', (e) => {
  e.preventDefault(); templateDrop.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleTemplate(e.dataTransfer.files[0]);
});
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
  showTemplateStatus(currentTemplate);
  renderTokens();
 
  if (tokens.length === 0) {
    emilySay(`Got your template! I don't see any <code>{{tokens}}</code> yet — no worries. Say "walk me through it" and I'll go section by section suggesting what to tokenize.`);
  } else {
    emilySay(`Great, I detected <strong>${tokens.length} tokens</strong> in your template. Upload some sample reports next so we can start mapping.`);
  }
}
 
async function parseDocx(file) {
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file('word/document.xml').async('string');
  const rawText = docXml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  const matches = rawText.match(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g) || [];
  const unique = [...new Set(matches.map(m => m.replace(/[\{\}\s]/g, '')))];
  return { tokens: unique, rawText };
}
 
async function extractRawTextFromBlob(blob) {
  const buf = await blob.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file('word/document.xml').async('string');
  return docXml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}
 
function showTemplateStatus(t) {
  const name = t.file_path.split('/').pop().replace(/^\d+_/, '');
  templateStatus.style.display = 'flex';
  templateStatus.innerHTML = `
    <div>
      <div class="filename">✓ ${name}</div>
      <div class="filemeta">${currentTokens.length} tokens • uploaded</div>
    </div>
    <button class="ghost-btn" style="padding:6px 12px;font-size:12px;" onclick="replaceTemplate()">Replace</button>
  `;
}
window.replaceTemplate = () => templateInput.click();
 
// =========================================================
// REPORT UPLOAD
// =========================================================
reportDrop.addEventListener('click', () => reportInput.click());
reportDrop.addEventListener('dragover', (e) => { e.preventDefault(); reportDrop.classList.add('dragover'); });
reportDrop.addEventListener('dragleave', () => reportDrop.classList.remove('dragover'));
reportDrop.addEventListener('drop', (e) => {
  e.preventDefault(); reportDrop.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleReport(e.dataTransfer.files[0]);
});
reportInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleReport(e.target.files[0]);
});
 
async function handleReport(file) {
  const reportName = prompt(`Name this report (e.g. "AR Aging", "P&L", "Estimates"):`, file.name.replace(/\.[^.]+$/, ''));
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
  renderTokens();
  emilySay(`Got the <strong>${reportName}</strong> report. Ask me to look at it and I'll help figure out what maps where.`);
}
 
function renderReports() {
  reportsList.innerHTML = '';
  currentReports.forEach(r => {
    const row = document.createElement('div');
    row.className = 'report-row';
    row.innerHTML = `
      <span>📊</span>
      <input type="text" value="${r.report_name}" data-id="${r.id}" />
      <button onclick="deleteReport('${r.id}')" title="Remove">×</button>
    `;
    reportsList.appendChild(row);
  });
  reportsList.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('change', async (e) => {
      const id = e.target.dataset.id;
      await db.from('client_reports').update({ report_name: e.target.value }).eq('id', id);
    });
  });
}
 
window.deleteReport = async (id) => {
  if (!confirm('Remove this report?')) return;
  await db.from('client_reports').delete().eq('id', id);
  currentReports = currentReports.filter(r => r.id !== id);
  renderReports(); renderTokens();
};
 
// =========================================================
// TOKENS + MAPPING
// =========================================================
function renderTokens() {
  if (!currentTokens || currentTokens.length === 0) {
    tokensList.innerHTML = `
      <div class="tokens-empty">
        ${currentTemplate
          ? 'No tokens in your template yet. Ask Emily to help tokenize it — click her bubble in the bottom right.'
          : 'Upload a template to detect tokens.'}
      </div>`;
    return;
  }
 
  tokensList.innerHTML = '';
  currentTokens.forEach(t => {
    const mapping = currentMappings[t.token] || {};
    const row = document.createElement('div');
    row.className = 'token-row';
 
    const reportOptions = currentReports.map(r =>
      `<option value="${r.id}" ${mapping.report_id === r.id ? 'selected' : ''}>${r.report_name}</option>`
    ).join('');
 
    row.innerHTML = `
      <div class="token-name">{{${t.token}}}</div>
      <select data-token="${t.token}" data-field="report_id">
        <option value="">— manual —</option>
        ${reportOptions}
      </select>
      <select data-token="${t.token}" data-field="operation">
        <option value="manual" ${mapping.operation === 'manual' ? 'selected' : ''}>Type manually</option>
        <option value="cell" ${mapping.operation === 'cell' ? 'selected' : ''}>Single cell</option>
        <option value="value_at_label" ${mapping.operation === 'value_at_label' ? 'selected' : ''}>Value at label</option>
        <option value="sum" ${mapping.operation === 'sum' ? 'selected' : ''}>Column sum</option>
        <option value="section_sum" ${mapping.operation === 'section_sum' ? 'selected' : ''}>Section sum</option>
        <option value="count" ${mapping.operation === 'count' ? 'selected' : ''}>Row count</option>
      </select>
      <div class="token-status">${mapping.id ? '✓' : '—'}</div>
    `;
    tokensList.appendChild(row);
  });
 
  tokensList.querySelectorAll('select').forEach(sel => {
    sel.addEventListener('change', (e) => saveMappingChange(e.target));
  });
}
 
async function saveMappingChange(el) {
  const token = el.dataset.token;
  const field = el.dataset.field;
  const value = el.value;
 
  const existing = currentMappings[token] || {
    client_id: currentClient.id, token, operation: 'manual',
    report_id: null, config: {}
  };
  existing[field] = value === '' ? null : value;
  if (!existing.operation) existing.operation = 'manual';
 
  const { data, error } = await db.from('client_mappings')
    .upsert(existing, { onConflict: 'client_id,token' }).select().single();
  if (error) { console.error('Mapping save error:', error); return; }
  currentMappings[token] = data;
  renderTokens();
}
 
// =========================================================
// EMILY CHAT BUBBLE
// =========================================================
const emilyBubble = document.getElementById('emilyBubble');
const emilyPanel = document.getElementById('emilyPanel');
const emilyMessages = document.getElementById('emilyMessages');
const emilyInput = document.getElementById('emilyInput');
const emilySend = document.getElementById('emilySend');
const emilyClose = document.getElementById('emilyClose');
const emilyExpand = document.getElementById('emilyExpand');
 
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
emilySend.addEventListener('click', sendEmilyMessage);
emilyInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendEmilyMessage(); }
});
 
function emilySay(html, saveHistory = true) {
  const msg = document.createElement('div');
  msg.className = 'msg emily';
  msg.innerHTML = html;
  emilyMessages.appendChild(msg);
  emilyMessages.scrollTop = emilyMessages.scrollHeight;
  const text = msg.textContent;
  if (saveHistory) {
    chatHistory.push({ role: 'assistant', content: text });
    saveChatMessage('assistant', text);
  }
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
  await db.from('client_chat_messages').insert({
    client_id: currentClient.id, role, content
  });
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
  const hadHistory = await loadChatHistory();
  if (hadHistory) return;
 
  const first = `Hey! 👋 I'm Emily, your mapping assistant for <strong>${currentClient.name}</strong>. `;
  if (!currentTemplate) {
    emilySay(first + `Start by uploading your Word template — no tokens needed, we'll add them together.`);
  } else if (currentTokens.length === 0) {
    emilySay(first + `I see your template. Want me to walk through it and suggest what should be tokenized?`);
  } else if (currentReports.length === 0) {
    emilySay(first + `I see <strong>${currentTokens.length} tokens</strong>. Upload some sample reports so we can map them.`);
  } else {
    emilySay(first + `We've got ${currentTokens.length} tokens and ${currentReports.length} reports. Ready to map — ask me anything.`);
  }
}
 
// =========================================================
// SEND MESSAGE — calls the Render /chat endpoint
// =========================================================
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
 
  // If Render is cold (spun down), first request takes ~50s
  const coldTimer = setTimeout(() => {
    typing.textContent = 'Emily is waking up (this is the first message — takes ~50s)…';
  }, 4000);
 
  try {
    // Keep payload small so Render's edge doesn't 403
    const shortHistory = chatHistory.slice(-8).map(m => ({
      role: m.role,
      content: (m.content || '').slice(0, 800)
    }));
 
    const body = {
      client_name: currentClient.name,
      template_text: (currentTemplate?.raw_text || '').slice(0, 2000),
      report_names: currentReports.map(r => r.report_name).slice(0, 10),
      sample_paths: currentReports.map(r => r.sample_path).slice(0, 5),
      current_tokens: currentTokens.map(t => t.token).slice(0, 30),
      current_mappings: Object.values(currentMappings).slice(0, 30).map(m => ({
        token: m.token,
        operation: m.operation
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
      const errText = await res.text();
      emilySay(`⚠️ Something went wrong: ${errText.slice(0, 200)}`);
      return;
    }
 
    const data = await res.json();
    const rendered = renderMarkdownLite(data.reply || '(no reply)');
    emilySay(rendered);
  } catch (err) {
    clearTimeout(coldTimer);
    typing.remove();
    emilySay(`⚠️ Couldn't reach the engine. Check that Render is running. Error: ${err.message}`);
  }
}
 
// Minimal markdown → HTML: bold, italics, code, line breaks
function renderMarkdownLite(text) {
  const escaped = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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
 
