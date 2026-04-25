const state = {
  clients: [],
  products: [],
  selectedClients: new Map(),
  workspaceItems: new Map(),
  selectedWorkspaceKeys: new Set(),
  filteredClients: [],
  filteredProducts: [],
  deferredPrompt: null,
  globalDates: {
    from: '',
    to: '',
  },
  openSection: 'clients',
};

const CLIENT_RESULT_LIMIT = 8;
const PRODUCT_RESULT_LIMIT = window.innerWidth <= 720 ? 28 : 90;

const DATASET_CONFIG = {
  clients: {
    baseName: 'klienci',
    normalize: normalizeClients,
  },
  products: {
    baseName: 'towar',
    normalize: normalizeProducts,
  },
};

const dom = {
  progressHeadline: document.getElementById('progressHeadline'),
  progressPercent: document.getElementById('progressPercent'),
  progressHint: document.getElementById('progressHint'),
  progressBackBtn: document.getElementById('progressBackBtn'),
  clientSearch: document.getElementById('clientSearch'),
  clientList: document.getElementById('clientList'),
  clientResultsInfo: document.getElementById('clientResultsInfo'),
  clientCountBadge: document.getElementById('clientCountBadge'),
  clientSectionInfo: document.getElementById('clientSectionInfo'),
  selectedClientsChips: document.getElementById('selectedClientsChips'),
  selectVisibleClientsBtn: document.getElementById('selectVisibleClientsBtn'),
  clearClientsBtn: document.getElementById('clearClientsBtn'),
  goProductsBtn: document.getElementById('goProductsBtn'),

  productSearch: document.getElementById('productSearch'),
  productList: document.getElementById('productList'),
  productResultsInfo: document.getElementById('productResultsInfo'),
  productSectionInfo: document.getElementById('productSectionInfo'),
  addVisibleProductsBtn: document.getElementById('addVisibleProductsBtn'),
  goWorkspaceBtn: document.getElementById('goWorkspaceBtn'),

  workspaceCards: document.getElementById('workspaceCards'),
  workspaceHelperText: document.getElementById('workspaceHelperText'),
  workspaceCountBadge: document.getElementById('workspaceCountBadge'),
  exportRowsBadge: document.getElementById('exportRowsBadge'),
  exportSectionInfo: document.getElementById('exportSectionInfo'),
  exportBtn: document.getElementById('stickyExportBtn'),
  exportReadinessHint: document.getElementById('exportReadinessHint'),
  stickyActionBar: document.getElementById('stickyActionBar'),
  stickyRowsCount: document.getElementById('stickyRowsCount'),
  stickySelectedCount: document.getElementById('stickySelectedCount'),
  clearWorkspaceBtn: document.getElementById('clearWorkspaceBtn'),
  selectAllWorkspaceBtn: document.getElementById('selectAllWorkspaceBtn'),
  clearWorkspaceSelectionBtn: document.getElementById('clearWorkspaceSelectionBtn'),
  applyBulkRefundBtn: document.getElementById('applyBulkRefundBtn'),
  bulkAcPrice: document.getElementById('bulkAcPrice'),
  bulkRefundValue: document.getElementById('bulkRefundValue'),
  bulkSelectionInfo: document.getElementById('bulkSelectionInfo'),

  reloadBtn: document.getElementById('reloadBtn'),
  installBtn: document.getElementById('installBtn'),
  toast: document.getElementById('toast'),
  refundDateFrom: document.getElementById('refundDateFrom'),
  refundDateTo: document.getElementById('refundDateTo'),
  sectionCards: [...document.querySelectorAll('.section-card')],
  sectionButtons: [...document.querySelectorAll('[data-open]')],
};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  bindEvents();
  registerServiceWorker();
  setupInstallPrompt();
  await loadAllData();
  updateStepLayout();
}

function bindEvents() {
  document.addEventListener('keydown', handleGlobalShortcuts);
  dom.clientSearch.addEventListener('input', renderClients);
  dom.productSearch.addEventListener('input', renderProducts);

  dom.refundDateFrom.addEventListener('input', event => {
    state.globalDates.from = event.target.value;
    updateExportSummary();
  });

  dom.refundDateTo.addEventListener('input', event => {
    state.globalDates.to = event.target.value;
    updateExportSummary();
  });

  dom.selectVisibleClientsBtn.addEventListener('click', () => {
    const visible = getFilteredClients().slice(0, CLIENT_RESULT_LIMIT);
    visible.forEach(client => state.selectedClients.set(client._key, client));
    renderClients();
    renderSelectedClients();
    updateExportSummary();
    if (state.selectedClients.size) openSection('products');
    showToast(`Zaznaczono ${visible.length} widocznych kontrahentów.`);
  });

  dom.clearClientsBtn.addEventListener('click', () => {
    state.selectedClients.clear();
    renderClients();
    renderSelectedClients();
    updateExportSummary();
    showToast('Wyczyszczono wybór kontrahentów.');
  });

  dom.addVisibleProductsBtn.addEventListener('click', () => {
    const visible = getFilteredProducts().slice(0, PRODUCT_RESULT_LIMIT);
    let added = 0;
    visible.forEach(item => {
      if (!state.workspaceItems.has(item._key)) {
        state.workspaceItems.set(item._key, createWorkspaceItem(item));
        state.selectedWorkspaceKeys.add(item._key);
        added += 1;
      }
    });
    renderProducts();
    renderWorkspace();
    updateExportSummary();
    if (added) openSection('workspace');
    showToast(added ? `Dodano ${added} widocznych indeksów.` : 'Brak nowych indeksów do dodania.');
  });

  dom.clearWorkspaceBtn.addEventListener('click', () => {
    state.workspaceItems.clear();
    state.selectedWorkspaceKeys.clear();
    renderProducts();
    renderWorkspace();
    updateExportSummary();
    showToast('Wyczyszczono listę refundacji.');
  });

  dom.selectAllWorkspaceBtn.addEventListener('click', () => {
    state.selectedWorkspaceKeys = new Set(state.workspaceItems.keys());
    renderWorkspace();
    updateExportSummary();
    showToast('Zaznaczono wszystkie indeksy do zmiany zbiorczej.');
  });

  dom.clearWorkspaceSelectionBtn.addEventListener('click', () => {
    state.selectedWorkspaceKeys.clear();
    renderWorkspace();
    updateExportSummary();
  });

  dom.applyBulkRefundBtn.addEventListener('click', applyBulkRefund);
  dom.exportBtn.addEventListener('click', exportData);

  dom.reloadBtn.addEventListener('click', async () => {
    await loadAllData(true);
    showToast('Wczytano ponownie pliki z folderu data.');
  });

  dom.installBtn.addEventListener('click', installApp);
  dom.goProductsBtn.addEventListener('click', () => openSection('products'));
  dom.goWorkspaceBtn.addEventListener('click', () => openSection('workspace'));
  dom.progressBackBtn.addEventListener('click', goToPreviousSection);

  dom.sectionButtons.forEach(button => {
    button.addEventListener('click', () => openSection(button.dataset.open));
  });
}

async function loadAllData(forceReload = false) {
  try {
    const [clientsData, productsData] = await Promise.all([
      loadDataset('clients', forceReload),
      loadDataset('products', forceReload),
    ]);

    state.clients = clientsData.rows;
    state.products = productsData.rows;

    pruneSelections();
    syncDateInputs();
    renderClients();
    renderProducts();
    renderSelectedClients();
    renderWorkspace();
    updateAccordionState();
    updateExportSummary();
  } catch (error) {
    console.error(error);
    showToast(error.message || 'Błąd wczytywania danych. Sprawdź pliki w folderze data.');
  }
}

async function loadDataset(kind, forceReload = false) {
  const config = DATASET_CONFIG[kind];
  const candidates = ['xlsx', 'xls', 'json'];
  const cacheMode = forceReload ? 'no-store' : 'default';

  for (const extension of candidates) {
    const url = `./data/${config.baseName}.${extension}${forceReload ? `?t=${Date.now()}` : ''}`;
    try {
      const response = await fetch(url, { cache: cacheMode });
      if (!response.ok) continue;

      let rawRows = [];
      if (extension === 'json') {
        rawRows = await response.json();
      } else {
        rawRows = await parseWorkbookBuffer(await response.arrayBuffer());
      }

      return { rows: config.normalize(rawRows) };
    } catch (error) {
      console.warn(`Pominięto ${url}`, error);
    }
  }

  throw new Error(`Nie znaleziono pliku ${config.baseName}.xlsx, ${config.baseName}.xls ani ${config.baseName}.json w folderze data.`);
}

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function headerKey(value) {
  return normalizeText(value).replace(/[^a-z0-9]/g, '');
}

function pickValue(row, aliases) {
  const map = new Map(Object.entries(row).map(([key, value]) => [headerKey(key), value]));
  for (const alias of aliases) {
    const found = map.get(headerKey(alias));
    if (found !== undefined && found !== null && String(found).trim() !== '') {
      return String(found).trim();
    }
  }
  return '';
}

function normalizeClients(rows) {
  const items = rows.map(row => {
    const Numer = pickValue(row, ['Numer', 'numer', 'Nr', 'nr']);
    const Logo = pickValue(row, ['Logo', 'logo']);
    const Nazwa = pickValue(row, ['Nazwa', 'nazwa']);
    const display = [Numer, Logo, Nazwa].filter(Boolean).join(' | ');

    return {
      _key: [Numer, Logo, Nazwa].join('||'),
      Numer,
      Logo,
      Nazwa,
      display,
      searchText: normalizeText(display),
    };
  }).filter(item => item.display);

  return dedupeBy(items, '_key').sort((a, b) => a.display.localeCompare(b.display, 'pl'));
}

const BARCODE_COLUMN_ALIASES = [
  'KodKreskowy',
  'KodPaskowy',
  'kodpaskowy',
  'Kod paskowy',
  'kod paskowy',
  'Kod paskowy EAN',
  'kod paskowy ean',
  'Kod kreskowy',
  'kod kreskowy',
  'Kod kreskowy EAN',
  'kod kreskowy ean',
  'Kod EAN',
  'EAN',
  'ean',
  'EAN13',
  'GTIN',
  'Barcode',
  'barcode',
];

function pickBarcodeValue(row) {
  const directMatch = pickValue(row, BARCODE_COLUMN_ALIASES);
  if (directMatch) return directMatch;

  const fallbackHeaders = ['ean', 'kresk', 'paskow', 'barcode', 'gtin'];
  for (const [key, value] of Object.entries(row)) {
    if (value === undefined || value === null) continue;
    const trimmed = String(value).trim();
    if (!trimmed) continue;

    const normalizedHeader = headerKey(key);
    if (fallbackHeaders.some(fragment => normalizedHeader.includes(fragment))) {
      return trimmed;
    }
  }

  return '';
}

function normalizeProducts(rows) {
  const items = rows.map(row => {
    const NazwaZnacznika = pickValue(row, ['NazwaZnacznika', 'Nazwa znacznika', 'Producent', 'producent']);
    const KodWlasny = pickValue(row, ['KodWlasny', 'Kod własny', 'kodwlasny', 'kod']);
    const Nazwa = pickValue(row, ['Nazwa', 'nazwa']);
    const IdTowaru = pickValue(row, ['IdTowaru', 'ID towaru', 'Id towaru', 'ID', 'id']);
    const KodKreskowy = pickBarcodeValue(row);
    return {
      _key: [KodWlasny, Nazwa, NazwaZnacznika, IdTowaru, KodKreskowy].join('||'),
      NazwaZnacznika,
      KodWlasny,
      Nazwa,
      IdTowaru,
      KodKreskowy,
      searchText: normalizeText(`${KodWlasny} ${Nazwa} ${NazwaZnacznika} ${IdTowaru} ${KodKreskowy}`),
    };
  }).filter(item => item.KodWlasny || item.Nazwa);

  return dedupeBy(items, '_key').sort((a, b) => {
    const byProducer = String(a.NazwaZnacznika).localeCompare(String(b.NazwaZnacznika), 'pl');
    if (byProducer !== 0) return byProducer;
    return String(a.Nazwa).localeCompare(String(b.Nazwa), 'pl');
  });
}

function dedupeBy(items, key) {
  const map = new Map();
  items.forEach(item => {
    if (!map.has(item[key])) map.set(item[key], item);
  });
  return [...map.values()];
}

function pruneSelections() {
  const validClientKeys = new Set(state.clients.map(item => item._key));
  [...state.selectedClients.keys()].forEach(key => {
    if (!validClientKeys.has(key)) state.selectedClients.delete(key);
  });

  const validProductKeys = new Set(state.products.map(item => item._key));
  [...state.workspaceItems.keys()].forEach(key => {
    if (!validProductKeys.has(key)) state.workspaceItems.delete(key);
  });
  [...state.selectedWorkspaceKeys].forEach(key => {
    if (!validProductKeys.has(key) || !state.workspaceItems.has(key)) state.selectedWorkspaceKeys.delete(key);
  });
}

function syncDateInputs() {
  dom.refundDateFrom.value = state.globalDates.from;
  dom.refundDateTo.value = state.globalDates.to;
}

function getFilteredClients() {
  const query = normalizeText(dom.clientSearch.value);
  const filtered = state.clients.filter(item => {
    if (state.selectedClients.has(item._key)) return false;
    if (!query) return true;
    return item.searchText.includes(query);
  });

  state.filteredClients = filtered;
  return filtered;
}

function toggleClient(itemKey, checked) {
  const item = state.clients.find(client => client._key === itemKey);
  if (!item) return;
  if (checked) state.selectedClients.set(itemKey, item);
  else state.selectedClients.delete(itemKey);
  renderClients();
  renderSelectedClients();
  updateExportSummary();
}

function renderClients() {
  const filtered = getFilteredClients();
  const visible = filtered.slice(0, CLIENT_RESULT_LIMIT);
  dom.clientList.innerHTML = '';
  dom.clientResultsInfo.textContent = filtered.length > CLIENT_RESULT_LIMIT
    ? `${formatNumber(filtered.length)} wyników, pokazano ${formatNumber(CLIENT_RESULT_LIMIT)}`
    : `${formatNumber(filtered.length)} wyników`;

  if (!filtered.length) {
    dom.clientList.innerHTML = '<div class="empty-state">Brak kontrahentów dla podanego filtra.</div>';
    updateClientBadge();
    return;
  }

  visible.forEach(item => {
    const isSelected = state.selectedClients.has(item._key);
    const row = document.createElement('article');
    row.className = `result-item client-item ${isSelected ? 'is-selected-client' : ''}`;
    row.innerHTML = `
      <label class="client-row-inner">
        <input class="client-check" type="checkbox" ${isSelected ? 'checked' : ''}>
        <span class="result-text">
          <span class="result-title">${escapeHtml(item.display)}</span>
        </span>
      </label>
    `;

    const checkbox = row.querySelector('input');
    checkbox.addEventListener('change', event => toggleClient(item._key, event.target.checked));
    row.addEventListener('click', event => {
      if (event.target.closest('input')) return;
      checkbox.checked = !checkbox.checked;
      toggleClient(item._key, checkbox.checked);
    });

    dom.clientList.appendChild(row);
  });

  if (filtered.length > CLIENT_RESULT_LIMIT) {
    const note = document.createElement('div');
    note.className = 'empty-state';
    note.textContent = 'Pokazano część wyników. Zawęź frazę, żeby szybciej wybrać kontrahenta.';
    dom.clientList.appendChild(note);
  }

  updateClientBadge();
}

function renderSelectedClients() {
  const items = [...state.selectedClients.values()].sort((a, b) => a.display.localeCompare(b.display, 'pl'));
  dom.selectedClientsChips.innerHTML = '';

  if (!items.length) {
    dom.selectedClientsChips.innerHTML = '<div class="empty-state">Nie wybrano jeszcze żadnego kontrahenta.</div>';
    dom.clientSectionInfo.textContent = 'Szukaj po numerze, logo lub nazwie';
    return;
  }

  dom.clientSectionInfo.textContent = `${formatNumber(items.length)} kontrahentów wybranych`;

  items.forEach(item => {
    const chip = document.createElement('div');
    chip.className = 'chip selected-chip';
    chip.innerHTML = `
      <span>${escapeHtml(item.display)}</span>
      <button type="button" aria-label="Usuń kontrahenta">×</button>
    `;
    chip.querySelector('button').addEventListener('click', () => {
      state.selectedClients.delete(item._key);
      renderClients();
      renderSelectedClients();
      updateExportSummary();
    });
    dom.selectedClientsChips.appendChild(chip);
  });
}

function getFilteredProducts() {
  const tokens = normalizeText(dom.productSearch.value)
    .split(/\s+/)
    .filter(Boolean);

  const filtered = state.products.filter(item => {
    if (state.workspaceItems.has(item._key)) return false;
    if (!tokens.length) return true;
    return tokens.every(token => item.searchText.includes(token));
  });

  state.filteredProducts = filtered;
  return filtered;
}

function renderProducts() {
  const filtered = getFilteredProducts();
  const visible = filtered.slice(0, PRODUCT_RESULT_LIMIT);
  dom.productList.innerHTML = '';
  dom.productResultsInfo.textContent = filtered.length > PRODUCT_RESULT_LIMIT
    ? `${formatNumber(filtered.length)} wyników, pokazano ${formatNumber(PRODUCT_RESULT_LIMIT)}`
    : `${formatNumber(filtered.length)} wyników`;

  if (!filtered.length) {
    dom.productList.innerHTML = '<div class="empty-state">Brak indeksów dla wybranego filtra.</div>';
    updateWorkspaceBadge();
    return;
  }

  visible.forEach(item => {
    const row = document.createElement('article');
    row.className = 'result-item product-item compact-product-item';
    row.innerHTML = `
      <div class="product-card-topline compact-topline-row">
        <span class="producer-pill">${escapeHtml(item.NazwaZnacznika || 'Brak producenta')}</span>
        ${item.KodKreskowy ? `<span class="barcode-inline" title="Kod kreskowy">EAN: ${escapeHtml(item.KodKreskowy)}</span>` : ''}
        <span class="code-pill">${escapeHtml(item.KodWlasny || 'Brak kodu')}</span>
      </div>
      <div class="product-main-row">
        <div class="product-name full-visible">${escapeHtml(item.Nazwa || 'Brak nazwy')}</div>
        <div class="product-card-actions compact-product-actions">
          <button type="button" class="add-btn small-btn compact-add-btn">Dodaj</button>
        </div>
      </div>
    `;

    row.querySelector('button').addEventListener('click', () => addProductToWorkspace(item));
    dom.productList.appendChild(row);
  });

  if (filtered.length > PRODUCT_RESULT_LIMIT) {
    const note = document.createElement('div');
    note.className = 'empty-state';
    note.textContent = 'Pokazano część wyników. Dopisz więcej liter, żeby szybciej trafić w indeks.';
    dom.productList.appendChild(note);
  }

  updateWorkspaceBadge();
}

function createWorkspaceItem(item) {
  return {
    _key: item._key,
    KodWlasny: item.KodWlasny,
    Nazwa: item.Nazwa,
    NazwaZnacznika: item.NazwaZnacznika,
    acPrice: '',
    refundValue: '',
  };
}

function addProductToWorkspace(item) {
  if (state.workspaceItems.has(item._key)) return;
  state.workspaceItems.set(item._key, createWorkspaceItem(item));
  state.selectedWorkspaceKeys.add(item._key);
  renderProducts();
  renderWorkspace();
  updateExportSummary();
  showToast(`Dodano indeks ${item.KodWlasny || item.Nazwa}.`);
}

function toggleWorkspaceSelection(itemKey, checked) {
  if (checked) state.selectedWorkspaceKeys.add(itemKey);
  else state.selectedWorkspaceKeys.delete(itemKey);
  renderWorkspace();
  updateExportSummary();
}

function renderWorkspace() {
  dom.workspaceCards.innerHTML = '';

  if (!state.workspaceItems.size) {
    dom.workspaceCards.innerHTML = '<div class="empty-state">Nie dodano jeszcze żadnych indeksów.</div>';
    dom.workspaceHelperText.textContent = 'Dodane: 0';
    dom.bulkSelectionInfo.textContent = '0 zaznaczonych indeksów';
    updateWorkspaceBadge();
    return;
  }

  const items = [...state.workspaceItems.values()].sort((a, b) => String(a.Nazwa).localeCompare(String(b.Nazwa), 'pl'));
  dom.workspaceHelperText.textContent = `Dodane: ${formatNumber(items.length)}`;
  dom.bulkSelectionInfo.textContent = `${formatNumber(state.selectedWorkspaceKeys.size)} zaznaczonych indeksów`;

  items.forEach(item => {
    const isSelected = state.selectedWorkspaceKeys.has(item._key);
    const card = document.createElement('article');
    card.className = `workspace-card compact-workspace-card ${isSelected ? 'selected' : ''}`;
    card.innerHTML = `
      <div class="workspace-top compact-workspace-top">
        <label class="workspace-check-wrap" title="Zaznacz do zmiany zbiorczej">
          <input class="workspace-check" type="checkbox" ${isSelected ? 'checked' : ''}>
        </label>
        <div class="workspace-product-meta compact-meta-row">
          <span class="producer-pill">${escapeHtml(item.NazwaZnacznika || 'Brak producenta')}</span>
          <span class="code-pill">${escapeHtml(item.KodWlasny || 'Brak kodu')}</span>
        </div>
        <button class="remove-btn small-btn compact-remove-btn" type="button">Usuń</button>
      </div>
      <button class="workspace-select-surface" type="button" aria-label="Zaznacz indeks do zmiany zbiorczej">
        <div class="workspace-title">${escapeHtml(item.Nazwa || 'Brak nazwy')}</div>
      </button>
      <div class="workspace-fields compact-workspace-fields">
        <label class="field-card compact-field">
          <span>Cena AC</span>
          <input class="search-input ac-price compact-input" type="number" min="0" step="0.01" inputmode="decimal" value="${escapeHtml(item.acPrice)}" placeholder="Np. 12,99">
        </label>
        <label class="field-card compact-field">
          <span>Refundacja</span>
          <input class="search-input refund-value compact-input" type="number" min="0" step="0.01" inputmode="decimal" value="${escapeHtml(item.refundValue)}" placeholder="Np. 2,50">
        </label>
      </div>
    `;

    const checkbox = card.querySelector('.workspace-check');
    const surface = card.querySelector('.workspace-select-surface');

    checkbox.addEventListener('change', event => toggleWorkspaceSelection(item._key, event.target.checked));
    surface.addEventListener('click', () => {
      checkbox.checked = !checkbox.checked;
      toggleWorkspaceSelection(item._key, checkbox.checked);
    });

    card.querySelector('.remove-btn').addEventListener('click', () => {
      state.workspaceItems.delete(item._key);
      state.selectedWorkspaceKeys.delete(item._key);
      renderProducts();
      renderWorkspace();
      updateExportSummary();
    });

    card.querySelector('.ac-price').addEventListener('input', event => {
      item.acPrice = event.target.value;
      updateExportSummary();
    });

    card.querySelector('.refund-value').addEventListener('input', event => {
      item.refundValue = event.target.value;
      updateExportSummary();
    });

    dom.workspaceCards.appendChild(card);
  });

  updateWorkspaceBadge();
}

function applyBulkRefund() {
  if (!state.selectedWorkspaceKeys.size) {
    showToast('Najpierw zaznacz indeksy do zmiany zbiorczej.');
    return;
  }

  const rawAcPrice = String(dom.bulkAcPrice.value ?? '').replace(',', '.').trim();
  const rawRefund = String(dom.bulkRefundValue.value ?? '').replace(',', '.').trim();

  const hasAcPrice = rawAcPrice !== '';
  const hasRefund = rawRefund !== '';

  if (!hasAcPrice && !hasRefund) {
    showToast('Podaj cenę AC lub refundację do zastosowania.');
    return;
  }

  if (hasAcPrice && Number.isNaN(Number(rawAcPrice))) {
    showToast('Podaj poprawną cenę AC.');
    return;
  }

  if (hasRefund && Number.isNaN(Number(rawRefund))) {
    showToast('Podaj poprawną refundację.');
    return;
  }

  state.selectedWorkspaceKeys.forEach(key => {
    const item = state.workspaceItems.get(key);
    if (!item) return;
    if (hasAcPrice) item.acPrice = rawAcPrice;
    if (hasRefund) item.refundValue = rawRefund;
  });

  renderWorkspace();
  updateExportSummary();
  showToast(`Zastosowano zmiany dla ${formatNumber(state.selectedWorkspaceKeys.size)} indeksów.`);
}

function updateClientBadge() {
  dom.clientCountBadge.textContent = formatNumber(state.selectedClients.size);
}

function updateWorkspaceBadge() {
  dom.workspaceCountBadge.textContent = formatNumber(state.workspaceItems.size);
  dom.productSectionInfo.textContent = `${formatNumber(state.workspaceItems.size)} indeksów dodanych do refundacji`;
}

function parseDecimal(value) {
  const raw = String(value ?? '').replace(',', '.').trim();
  if (!raw) return null;
  const num = Number(raw);
  return Number.isFinite(num) ? num : null;
}

function hasRefundValue(item) {
  return parseDecimal(item.refundValue) !== null;
}

function exportNumber(value) {
  const parsed = parseDecimal(value);
  return parsed === null ? '' : parsed;
}

function sanitizeFilenamePart(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-{2,}/g, '-')
    .toLowerCase();
}

function buildExportFilename(extension, clients = [...state.selectedClients.values()]) {
  const contractorNames = clients
    .map(client => client.Nazwa || client.Logo || client.Numer)
    .filter(Boolean);

  const uniqueContractorNames = [...new Set(contractorNames)];
  let clientPart = 'klient';

  if (uniqueContractorNames.length === 1) {
    clientPart = uniqueContractorNames[0];
  } else if (uniqueContractorNames.length === 2) {
    clientPart = `${uniqueContractorNames[0]} i ${uniqueContractorNames[1]}`;
  } else if (uniqueContractorNames.length > 2) {
    clientPart = `${uniqueContractorNames[0]} +${uniqueContractorNames.length - 1}`;
  }

  const safeClientPart = sanitizeFilenamePart(clientPart) || 'klient';
  const safeFrom = sanitizeFilenamePart(state.globalDates.from) || 'brak-daty-od';
  const safeTo = sanitizeFilenamePart(state.globalDates.to) || 'brak-daty-do';
  let filename = `refundacje-${safeClientPart}-${safeFrom}-${safeTo}`;

  if (filename.length > 120) {
    filename = filename.slice(0, 120).replace(/-+$/g, '');
  }

  return `${filename}.${extension}`;
}

function buildExportRows() {
  const clients = [...state.selectedClients.values()];
  return buildExportRowsForClients(clients);
}

function buildExportRowsForClients(clients) {
  const products = [...state.workspaceItems.values()].filter(item => hasRefundValue(item));
  const rows = [];

  clients.forEach(client => {
    products.forEach(product => {
      rows.push({
        numer: client.Numer,
        logo: client.Logo,
        KodWlasny: product.KodWlasny,
        NazwaZnacznika: product.NazwaZnacznika,
        Nazwa: product.Nazwa,
        refundacja_od: state.globalDates.from,
        refundacja_do: state.globalDates.to,
        'cena AC': exportNumber(product.acPrice),
        refundacja: exportNumber(product.refundValue),
      });
    });
  });

  return rows;
}

function updateExportSummary() {
  const rows = buildExportRows();
  dom.exportRowsBadge.textContent = formatNumber(rows.length);
  dom.stickyRowsCount.textContent = formatNumber(rows.length);
  dom.stickySelectedCount.textContent = formatNumber(state.selectedWorkspaceKeys.size);
  dom.exportSectionInfo.textContent = `${formatNumber(rows.length)} wierszy gotowych do eksportu`;
  updateProgressStatus(rows.length);
  updateExportState();
}

function validateBeforeExport() {
  if (!state.selectedClients.size) return 'Najpierw wybierz co najmniej jednego kontrahenta.';
  if (!state.workspaceItems.size) return 'Najpierw dodaj co najmniej jeden indeks.';
  if (!state.globalDates.from || !state.globalDates.to) return 'Uzupełnij zakres dat refundacji: od i do.';
  if (state.globalDates.from > state.globalDates.to) return 'Data „od” nie może być późniejsza niż data „do”.';

  const invalidItems = [...state.workspaceItems.values()].filter(item => !hasRefundValue(item));
  if (invalidItems.length) return 'Uzupełnij refundację dla wszystkich indeksów w obszarze roboczym.';

  return '';
}

function exportData() {
  const validationError = validateBeforeExport();
  if (validationError) {
    showToast(validationError);
    return;
  }

  const rows = buildExportRows();
  if (!rows.length) {
    showToast('Brak danych do eksportu.');
    return;
  }

  const headers = ['numer', 'logo', 'KodWlasny', 'NazwaZnacznika', 'Nazwa', 'refundacja_od', 'refundacja_do', 'cena AC', 'refundacja'];
  const selectedClients = [...state.selectedClients.values()];

  if (window.XLSX) {
    if (selectedClients.length > 1) {
      selectedClients.forEach(client => {
        const clientRows = buildExportRowsForClients([client]);
        const worksheet = XLSX.utils.json_to_sheet(clientRows, { header: headers });
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Refundacje');
        XLSX.writeFile(workbook, buildExportFilename('xlsx', [client]));
      });
      showToast(`Wyeksportowano ${formatNumber(selectedClients.length)} plików (po 1 na kontrahenta).`);
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Refundacje');
    XLSX.writeFile(workbook, buildExportFilename('xlsx', selectedClients));
    showToast(`Wyeksportowano ${formatNumber(rows.length)} wierszy.`);
    return;
  }

  if (selectedClients.length > 1) {
    selectedClients.forEach(client => {
      const clientRows = buildExportRowsForClients([client]);
      const csv = createCsv(clientRows);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = buildExportFilename('csv', [client]);
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    });
    showToast(`Wyeksportowano ${formatNumber(selectedClients.length)} plików (po 1 na kontrahenta).`);
    return;
  }

  const csv = createCsv(rows);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = buildExportFilename('csv', selectedClients);
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showToast(`Wyeksportowano ${formatNumber(rows.length)} wierszy.`);
}

function createCsv(rows) {
  const headers = ['numer', 'logo', 'KodWlasny', 'NazwaZnacznika', 'Nazwa', 'refundacja_od', 'refundacja_do', 'cena AC', 'refundacja'];
  const lines = [headers.join(';')];
  rows.forEach(row => {
    lines.push([
      row.numer,
      row.logo,
      row.KodWlasny,
      row.NazwaZnacznika,
      row.Nazwa,
      row.refundacja_od,
      row.refundacja_do,
      row['cena AC'],
      row.refundacja,
    ].map(value => `"${String(value ?? '').replace(/"/g, '""')}"`).join(';'));
  });
  return `\uFEFF${lines.join('\n')}`;
}

async function parseWorkbookBuffer(buffer) {
  if (!window.XLSX) {
    throw new Error('Nie udało się załadować silnika XLS/XLSX. Otwórz aplikację z internetem i spróbuj ponownie.');
  }

  const workbook = XLSX.read(buffer, { type: 'array', raw: false, cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];
  const worksheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(worksheet, { defval: '' });
}

function openSection(section) {
  state.openSection = section;
  updateAccordionState();
  updateStepLayout();
  updateProgressStatus(buildExportRows().length);
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function goToPreviousSection() {
  const sectionOrder = ['clients', 'products', 'workspace'];
  const currentIndex = sectionOrder.indexOf(state.openSection);
  if (currentIndex <= 0) return;
  openSection(sectionOrder[currentIndex - 1]);
}

function updateAccordionState() {
  dom.sectionCards.forEach(card => {
    card.classList.toggle('active', card.dataset.section === state.openSection);
  });
}

function updateStepLayout() {
  const inWorkspace = state.openSection === 'workspace';
  dom.stickyActionBar.classList.toggle('hidden-step', !inWorkspace);
  document.body.classList.toggle('workspace-step-active', inWorkspace);
}

function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) return;
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(error => console.warn('Service Worker:', error));
  });
}

function setupInstallPrompt() {
  window.addEventListener('beforeinstallprompt', event => {
    event.preventDefault();
    state.deferredPrompt = event;
    dom.installBtn.classList.remove('hidden');
  });

  window.addEventListener('appinstalled', () => {
    state.deferredPrompt = null;
    dom.installBtn.classList.add('hidden');
    showToast('Aplikacja została zainstalowana.');
  });
}

async function installApp() {
  if (!state.deferredPrompt) {
    showToast('Opcja instalacji nie jest teraz dostępna.');
    return;
  }

  state.deferredPrompt.prompt();
  await state.deferredPrompt.userChoice;
  state.deferredPrompt = null;
  dom.installBtn.classList.add('hidden');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatNumber(value) {
  return new Intl.NumberFormat('pl-PL').format(Number(value || 0));
}

let toastTimer = null;
function showToast(message) {
  dom.toast.textContent = message;
  dom.toast.classList.add('show');
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => dom.toast.classList.remove('show'), 2600);
}

function updateProgressStatus(exportRowsCount = 0) {
  const sectionOrder = ['clients', 'products', 'workspace'];
  const currentIndex = sectionOrder.indexOf(state.openSection);
  const canGoBack = currentIndex > 0;
  const previousLabel = canGoBack
    ? (sectionOrder[currentIndex - 1] === 'clients' ? 'Kontrahenci' : 'Indeksy')
    : '';
  dom.progressBackBtn.disabled = !canGoBack;
  dom.progressBackBtn.classList.toggle('hidden', !canGoBack);
  dom.progressBackBtn.textContent = canGoBack ? `← ${previousLabel}` : '← Wstecz';

  const hasClients = state.selectedClients.size > 0;
  const hasProducts = state.workspaceItems.size > 0;
  const hasDates = Boolean(state.globalDates.from && state.globalDates.to && state.globalDates.from <= state.globalDates.to);
  const hasRefunds = !hasProducts || [...state.workspaceItems.values()].every(item => hasRefundValue(item));

  let doneSteps = 0;
  if (hasClients) doneSteps += 1;
  if (hasProducts) doneSteps += 1;
  if (hasDates && hasRefunds && exportRowsCount > 0) doneSteps += 1;

  const percent = Math.round((doneSteps / 3) * 100);
  dom.progressPercent.textContent = `${percent}%`;

  if (!hasClients) {
    dom.progressHeadline.textContent = 'Krok 1 z 3: wybierz kontrahentów';
    dom.progressHint.textContent = 'Wybierz minimum 1 kontrahenta, aby przejść dalej.';
    return;
  }
  if (!hasProducts) {
    dom.progressHeadline.textContent = 'Krok 2 z 3: dodaj indeksy do refundacji';
    dom.progressHint.textContent = 'Wyszukaj indeks i dodaj go przyciskiem „Dodaj” lub „Dodaj widoczne”.';
    return;
  }
  if (!hasDates || !hasRefunds || !exportRowsCount) {
    dom.progressHeadline.textContent = 'Krok 3 z 3: uzupełnij daty i refundacje';
    dom.progressHint.textContent = 'Podaj zakres dat i refundację dla każdego indeksu, aby odblokować eksport.';
    return;
  }
  dom.progressHeadline.textContent = 'Gotowe do eksportu';
  dom.progressHint.textContent = 'Wszystkie kroki zakończone. Możesz wyeksportować plik skrótem Ctrl/Cmd + Enter.';
}

function updateExportState() {
  const validationMessage = validateBeforeExport();
  const isReady = !validationMessage;
  dom.exportBtn.disabled = !isReady;
  dom.exportReadinessHint.textContent = isReady ? 'Gotowe: możesz wyeksportować dane.' : validationMessage;
}

function handleGlobalShortcuts(event) {
  if (event.defaultPrevented) return;
  const target = event.target;
  const isTypingContext = target instanceof HTMLElement && (
    target.tagName === 'INPUT' ||
    target.tagName === 'TEXTAREA' ||
    target.isContentEditable
  );

  if (event.key === '/' && !isTypingContext) {
    event.preventDefault();
    const inputBySection = {
      clients: dom.clientSearch,
      products: dom.productSearch,
      workspace: dom.bulkRefundValue,
    };
    inputBySection[state.openSection]?.focus();
    return;
  }

  const isSubmitShortcut = (event.ctrlKey || event.metaKey) && event.key === 'Enter';
  if (!isSubmitShortcut) return;
  event.preventDefault();
  if (dom.exportBtn.disabled) {
    showToast(validateBeforeExport() || 'Eksport jest jeszcze niedostępny.');
    return;
  }
  exportData();
}
