// =========================================================
// Emily's Snapshot Studio — dashboard logic
// =========================================================

const db = window.supabase.createClient(
  window.SUPABASE_URL,
  window.SUPABASE_ANON_KEY
);

const honeycombEl = document.getElementById('honeycomb');
const emptyStateEl = document.getElementById('emptyState');
const clientCountEl = document.getElementById('clientCount');
const readyCountEl = document.getElementById('readyCount');
const setupCountEl = document.getElementById('setupCount');

const modal = document.getElementById('modalBackdrop');
const nameInput = document.getElementById('newClientName');
const slugInput = document.getElementById('newClientSlug');

document.getElementById('addClientBtn').addEventListener('click', openModal);
document.getElementById('cancelBtn').addEventListener('click', closeModal);
document.getElementById('saveClientBtn').addEventListener('click', createClient);
document.getElementById('refreshBtn').addEventListener('click', loadClients);

nameInput.addEventListener('input', () => {
  slugInput.value = nameInput.value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
});

async function loadClients() {
  honeycombEl.innerHTML = '';

  const { data: clients, error } = await db
    .from('clients')
    .select(`
      id, slug, name, created_at,
      client_templates(id),
      client_mappings(id)
    `)
    .order('created_at', { ascending: true });

  if (error) {
    console.error('Load error:', error);
    honeycombEl.innerHTML = `<p style="color:#ffb020">Error loading clients: ${error.message}</p>`;
    return;
  }

  const total = clients.length;
  const ready = clients.filter(c =>
    c.client_templates?.length > 0 && c.client_mappings?.length > 0
  ).length;
  const setup = total - ready;

  clientCountEl.textContent = total;
  readyCountEl.textContent = ready;
  setupCountEl.textContent = setup;

  clients.forEach(client => {
    honeycombEl.appendChild(buildHex(client));
  });

  honeycombEl.appendChild(buildAddHex());
  emptyStateEl.style.display = 'none';
}

function buildHex(client) {
  const initials = client.name
    .split(/\s+/)
    .map(w => w[0])
    .join('')
    .slice(0, 3)
    .toUpperCase();

  const isReady =
    client.client_templates?.length > 0 && client.client_mappings?.length > 0;

  const wrap = document.createElement('div');
  wrap.className = 'hex-wrap';
  wrap.innerHTML = `
    <div class="hex">
      <div class="hex-status ${isReady ? 'ready' : 'setup'}" title="${isReady ? 'Ready' : 'Needs setup'}"></div>
      <div class="hex-inner">
        <div class="hex-initials">${initials}</div>
        <div class="hex-name">${client.name}</div>
      </div>
    </div>
  `;
  wrap.addEventListener('click', () => openClient(client.slug));
  return wrap;
}

function buildAddHex() {
  const wrap = document.createElement('div');
  wrap.className = 'hex-wrap add-hex';
  wrap.innerHTML = `
    <div class="hex">
      <div class="hex-inner">
        <div class="hex-initials">+</div>
        <div class="hex-name">Add Client</div>
      </div>
    </div>
  `;
  wrap.addEventListener('click', openModal);
  return wrap;
}

function openClient(slug) {
  window.location.href = `/client.html?slug=${slug}`;
}

function openModal() {
  modal.style.display = 'flex';
  nameInput.value = '';
  slugInput.value = '';
  nameInput.focus();
}

function closeModal() {
  modal.style.display = 'none';
}

async function createClient() {
  const name = nameInput.value.trim();
  const slug = slugInput.value.trim();

  if (!name || !slug) {
    alert('Both fields are required.');
    return;
  }
  if (!/^[a-z0-9-]+$/.test(slug)) {
    alert('Slug can only contain lowercase letters, numbers, and hyphens.');
    return;
  }

  const { error } = await db.from('clients').insert({ name, slug });

  if (error) {
    alert('Error: ' + error.message);
    return;
  }

  closeModal();
  loadClients();
}

modal.addEventListener('click', (e) => {
  if (e.target === modal) closeModal();
});

loadClients();
