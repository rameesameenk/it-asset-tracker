const SAVED_BOQS_KEY = "boq_saved_items_v1";
const APPROVED_BOQS_KEY = "boq_approved_items_v1";
const CLIENTS_KEY = "boq_clients_v1";
const PROJECTS_KEY = "boq_projects_v1";
const SUGGEST_MEMORY_KEY = "boq_suggest_memory_v1";
const USERS_KEY = "boq_users_v1";
const USER_PROFILES_KEY = "boq_user_profiles_v1";
const SESSION_KEY = "boq_session_v1";

const loginPage = document.getElementById("login-page");
const loginForm = document.getElementById("login-form");
const loginUsernameInput = document.getElementById("login-username");
const loginPasswordInput = document.getElementById("login-password");
const loginStatus = document.getElementById("login-status");
const projectNameInput = document.getElementById("project-name");
const projectNameSuggestions = document.getElementById("project-name-suggestions");
const clientNameInput = document.getElementById("client-name");
const clientNameSuggestions = document.getElementById("client-name-suggestions");
const preparedByInput = document.getElementById("prepared-by");
const boqDateInput = document.getElementById("boq-date");
const currencySelect = document.getElementById("currency");

const addItemBtn = document.getElementById("add-item");
const clearItemsBtn = document.getElementById("clear-items");
const itemsTbody = document.getElementById("items-tbody");
const itemStatus = document.getElementById("item-status");

const discountInput = document.getElementById("discount");
const taxInput = document.getElementById("tax");
const contingencyInput = document.getElementById("contingency");

const subtotalValue = document.getElementById("subtotal-value");
const discountValue = document.getElementById("discount-value");
const taxValue = document.getElementById("tax-value");
const contingencyValue = document.getElementById("contingency-value");
const grandTotal = document.getElementById("grand-total");

const saveBtn = document.getElementById("save-boq");
const saveAsNewBtn = document.getElementById("save-boq-new");
const loadBtn = document.getElementById("load-boq");
const printBoqBtn = document.getElementById("print-boq");
const exportCsvBtn = document.getElementById("export-csv");
const exportJsonBtn = document.getElementById("export-json");
const summaryStatus = document.getElementById("summary-status");
const userCorner = document.getElementById("user-corner");
const userMenuTrigger = document.getElementById("user-menu-trigger");
const userMenu = document.getElementById("user-menu");
const userPhoto = document.getElementById("user-photo");
const userNameLabel = document.getElementById("user-name-label");
const userPhotoBtn = document.getElementById("user-photo-btn");
const userPhotoInput = document.getElementById("user-photo-input");
const userLogoutBtn = document.getElementById("user-logout-btn");
const homePage = document.getElementById("home-page");
const homeCreateBoqBtn = document.getElementById("home-create-boq");
const homeClientProjectsBtn = document.getElementById("home-client-projects");
const homeSavedBoqsBtn = document.getElementById("home-saved-boqs");
const homeSettingsBtn = document.getElementById("home-settings");
const clientProjectPage = document.getElementById("client-project-page");
const cpHomeBtn = document.getElementById("cp-home");
const cpClientNameInput = document.getElementById("cp-client-name");
const cpAddClientBtn = document.getElementById("cp-add-client");
const cpProjectClientSelect = document.getElementById("cp-project-client");
const cpProjectNameInput = document.getElementById("cp-project-name");
const cpProjectNoteInput = document.getElementById("cp-project-note-input");
const cpAddNoteBtn = document.getElementById("cp-add-note");
const cpNoteList = document.getElementById("cp-note-list");
const cpProjectRemarkInput = document.getElementById("cp-project-remark-input");
const cpAddRemarkBtn = document.getElementById("cp-add-remark");
const cpRemarkList = document.getElementById("cp-remark-list");
const cpAddProjectBtn = document.getElementById("cp-add-project");
const cpClientTbody = document.getElementById("cp-client-tbody");
const cpProjectTbody = document.getElementById("cp-project-tbody");
const cpStatus = document.getElementById("cp-status");
const viewProjectNotesBtn = document.getElementById("view-project-notes");
const projectNotesPanel = document.getElementById("project-notes-panel");
const projectNotesText = document.getElementById("project-notes-text");
const projectRemarksText = document.getElementById("project-remarks-text");
const editorHomeBtn = document.getElementById("editor-home");
const openSavedTopBtn = document.getElementById("open-saved-top");
const editorSettingsBtn = document.getElementById("editor-settings");
const editorPage = document.getElementById("editor-page");
const savedBoqPage = document.getElementById("saved-boq-page");
const backToEditorBtn = document.getElementById("back-to-editor");
const savedHomeBtn = document.getElementById("saved-home");
const savedSettingsBtn = document.getElementById("saved-settings");
const savedBoqTbody = document.getElementById("saved-boq-tbody");
const savedBoqStatus = document.getElementById("saved-boq-status");
const approvedBoqTbody = document.getElementById("approved-boq-tbody");
const approvedBoqStatus = document.getElementById("approved-boq-status");
const savedClientButtons = document.getElementById("saved-client-buttons");
const savedProjectButtons = document.getElementById("saved-project-buttons");
const savedPreparedByFilter = document.getElementById("saved-prepared-by-filter");
const savedSearchInput = document.getElementById("saved-search-input");
const settingsPage = document.getElementById("settings-page");
const settingsBackBtn = document.getElementById("settings-back");
const settingsHomeBtn = document.getElementById("settings-home");
const downloadBackupBtn = document.getElementById("download-backup");
const restoreBackupBtn = document.getElementById("restore-backup");
const backupFileInput = document.getElementById("backup-file-input");
const adminUserSection = document.getElementById("admin-user-section");
const adminNewUsername = document.getElementById("admin-new-username");
const adminNewPassword = document.getElementById("admin-new-password");
const adminNewFirstName = document.getElementById("admin-new-first-name");
const adminNewLastName = document.getElementById("admin-new-last-name");
const adminNewRole = document.getElementById("admin-new-role");
const adminAddUserBtn = document.getElementById("admin-add-user-btn");
const adminUserTbody = document.getElementById("admin-user-tbody");
const settingsStatus = document.getElementById("settings-status");

const UNIT_OPTIONS = ["Nos", "m", "m2", "m3", "kg", "ton", "Ltr", "sqft", "set", "day", "job"];

let state = {
  projectName: "",
  clientName: "",
  preparedBy: "",
  date: "",
  currency: "INR",
  discount: 0,
  tax: 0,
  contingency: 0,
  items: []
};
let savedBoqs = [];
let approvedBoqs = [];
let clients = [];
let projects = [];
let activeSavedBoqId = null;
let activeBoqNumber = "";
let activeProjectId = "";
let savedSelectedClientId = "";
let savedSelectedProjectId = "";
let settingsReturnPage = "home";
let users = [];
let userProfiles = {};
let suggestMemory = { clients: {}, projects: {} };
let sessionUser = "";
let draftProjectNotes = [];
let draftProjectRemarks = [];

function uid() {
  return window.crypto?.randomUUID
    ? window.crypto.randomUUID()
    : `id_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function escapeHtml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function currencySymbol(code) {
  if (code === "EUR") return "€";
  if (code === "INR") return "₹";
  if (code === "AED") return "د.إ";
  return "$";
}

function setStatus(el, message, type = "ok") {
  el.textContent = message;
  el.className = `status ${type}`;
}

function loadSuggestMemory() {
  try {
    const raw = JSON.parse(localStorage.getItem(SUGGEST_MEMORY_KEY) || "{}");
    suggestMemory = raw && typeof raw === "object" ? raw : {};
  } catch {
    suggestMemory = {};
  }
  if (!suggestMemory.clients || typeof suggestMemory.clients !== "object") suggestMemory.clients = {};
  if (!suggestMemory.projects || typeof suggestMemory.projects !== "object") suggestMemory.projects = {};
}

function persistSuggestMemory() {
  localStorage.setItem(SUGGEST_MEMORY_KEY, JSON.stringify(suggestMemory));
}

function normalizeKey(text) {
  return String(text || "").trim().toLowerCase();
}

function rememberSuggestion(kind, text, increment = 1) {
  const key = normalizeKey(text);
  if (!key) return;
  if (!suggestMemory[kind]) suggestMemory[kind] = {};
  suggestMemory[kind][key] = (Number(suggestMemory[kind][key]) || 0) + increment;
}

function topSuggestedValues(kind, candidates, query = "") {
  const q = normalizeKey(query);
  const unique = new Map();
  candidates.forEach((raw) => {
    const value = String(raw || "").trim();
    if (!value) return;
    const key = normalizeKey(value);
    if (!q || key.includes(q)) {
      if (!unique.has(key)) unique.set(key, value);
    }
  });
  const memoryRows = Object.keys(suggestMemory[kind] || {})
    .filter((key) => !q || key.includes(q))
    .map((key) => ({ key, value: unique.get(key) || key }));
  memoryRows.forEach((row) => {
    if (!unique.has(row.key)) unique.set(row.key, row.value);
  });
  return Array.from(unique.entries())
    .map(([key, value]) => ({ key, value, score: Number(suggestMemory[kind]?.[key] || 0) }))
    .sort((a, b) => b.score - a.score || a.value.localeCompare(b.value))
    .slice(0, 20)
    .map((row) => row.value);
}

function renderSuggestList(box, items, field) {
  if (!box) return;
  if (!items.length) {
    box.innerHTML = "";
    box.classList.add("hidden");
    return;
  }
  box.innerHTML = items
    .map((text) => `<button type="button" class="suggest-item" data-field="${field}" data-value="${escapeHtml(text)}">${escapeHtml(text)}</button>`)
    .join("");
  box.classList.remove("hidden");
}

function seedDefaultUser() {
  try {
    const raw = JSON.parse(localStorage.getItem(USERS_KEY) || "[]");
    users = Array.isArray(raw) ? raw : [];
  } catch {
    users = [];
  }
  users = users
    .map((row) => ({
      username: String(row?.username || "").trim(),
      password: String(row?.password || ""),
      firstName: String(row?.firstName || "").trim(),
      lastName: String(row?.lastName || "").trim(),
      role: String(row?.role || "").toLowerCase() === "admin" ? "admin" : "user"
    }))
    .filter((row) => row.username && row.password);
  if (!users.some((row) => row.username.toLowerCase() === "ramees")) {
    users.push({ username: "ramees", password: "IT@Admin", firstName: "Ramees", lastName: "", role: "admin" });
  } else {
    users = users.map((row) => {
      if (row.username.toLowerCase() !== "ramees") return row;
      return { ...row, role: "admin", firstName: row.firstName || "Ramees" };
    });
  }
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
}

function currentUserRecord() {
  return users.find((u) => u.username === sessionUser) || null;
}

function isAdminUser() {
  return currentUserRecord()?.role === "admin";
}

function userDisplayName(user) {
  const first = String(user?.firstName || "").trim();
  const last = String(user?.lastName || "").trim();
  const joined = `${first} ${last}`.trim();
  return joined || String(user?.username || "");
}

function refreshPreparedBy() {
  const user = currentUserRecord();
  const name = userDisplayName(user);
  if (preparedByInput) preparedByInput.value = name;
}

function renderUserTable() {
  if (!adminUserTbody) return;
  const rows = users
    .slice()
    .sort((a, b) => a.username.localeCompare(b.username))
    .map((u) => `
      <tr>
        <td>${escapeHtml(u.username)}</td>
        <td>${escapeHtml(u.firstName || "-")}</td>
        <td>${escapeHtml(u.lastName || "-")}</td>
        <td>${escapeHtml(u.role)}</td>
      </tr>
    `).join("");
  adminUserTbody.innerHTML = rows || '<tr><td colspan="4">No users found.</td></tr>';
}

function addUserByAdmin() {
  if (!isAdminUser()) {
    setStatus(settingsStatus, "Only admin can add users.", "err");
    return;
  }
  const username = String(adminNewUsername.value || "").trim();
  const password = String(adminNewPassword.value || "");
  const firstName = String(adminNewFirstName.value || "").trim();
  const lastName = String(adminNewLastName.value || "").trim();
  const role = String(adminNewRole.value || "user") === "admin" ? "admin" : "user";
  if (!username || !password || !firstName) {
    setStatus(settingsStatus, "Username, password and first name are required.", "err");
    return;
  }
  if (users.some((u) => u.username.toLowerCase() === username.toLowerCase())) {
    setStatus(settingsStatus, "Username already exists.", "err");
    return;
  }
  users.push({ username, password, firstName, lastName, role });
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
  adminNewUsername.value = "";
  adminNewPassword.value = "";
  adminNewFirstName.value = "";
  adminNewLastName.value = "";
  if (adminNewRole) adminNewRole.value = "user";
  renderUserTable();
  setStatus(settingsStatus, `User added: ${username}`, "ok");
}

function loadUserProfiles() {
  try {
    const raw = JSON.parse(localStorage.getItem(USER_PROFILES_KEY) || "{}");
    userProfiles = raw && typeof raw === "object" ? raw : {};
  } catch {
    userProfiles = {};
  }
}

function persistUserProfiles() {
  localStorage.setItem(USER_PROFILES_KEY, JSON.stringify(userProfiles));
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Failed to read image file."));
    reader.readAsDataURL(file);
  });
}

function refreshUserCorner() {
  const show = isAuthenticated();
  userCorner?.classList.toggle("hidden", !show);
  if (!show) return;
  if (userNameLabel) userNameLabel.textContent = userDisplayName(currentUserRecord());
  const photo = userProfiles?.[sessionUser]?.photo || "assets/sribal-logo.png";
  if (userPhoto) userPhoto.src = photo;
}

function closeUserMenu() {
  userMenu?.classList.add("hidden");
}

function isAuthenticated() {
  return !!sessionUser;
}

function tryLogin(username, password) {
  const u = String(username || "").trim();
  const p = String(password || "");
  const found = users.find((row) => row.username.toLowerCase() === u.toLowerCase() && row.password === p);
  if (!found) return false;
  sessionUser = found.username;
  localStorage.setItem(SESSION_KEY, sessionUser);
  refreshPreparedBy();
  refreshUserCorner();
  return true;
}

function logout() {
  sessionUser = "";
  localStorage.removeItem(SESSION_KEY);
  loginForm.reset();
  setStatus(loginStatus, "", "ok");
  closeUserMenu();
  refreshUserCorner();
  showPage("login");
}

function blankState() {
  return {
    projectId: "",
    projectName: "",
    clientName: "",
    preparedBy: "",
    date: "",
    currency: "INR",
    discount: 0,
    tax: 0,
    contingency: 0,
    items: []
  };
}

function cloneState(input) {
  return {
    projectId: String(input?.projectId || ""),
    projectName: String(input?.projectName || ""),
    clientName: String(input?.clientName || ""),
    preparedBy: String(input?.preparedBy || ""),
    date: String(input?.date || ""),
    currency: "INR",
    discount: toNumber(input?.discount),
    tax: toNumber(input?.tax),
    contingency: toNumber(input?.contingency),
    items: Array.isArray(input?.items) ? input.items.map((row) => ({
      id: row.id || uid(),
      description: String(row.description || ""),
      unit: String(row.unit || "Nos"),
      qty: toNumber(row.qty),
      rate: toNumber(row.rate),
      waste: toNumber(row.waste)
    })) : []
  };
}

function getBoqLabel(boqState) {
  const name = String(boqState?.projectName || "").trim();
  return name || "Untitled BOQ";
}

function addItem(partial = {}) {
  state.items.push({
    id: partial.id || uid(),
    description: partial.description || "",
    unit: partial.unit || "Nos",
    qty: toNumber(partial.qty),
    rate: toNumber(partial.rate),
    waste: toNumber(partial.waste)
  });
}

function itemAmount(item) {
  const base = toNumber(item.qty) * toNumber(item.rate);
  const wasteFactor = 1 + toNumber(item.waste) / 100;
  return base * wasteFactor;
}

function formatMoney(amount) {
  return `${currencySymbol(state.currency)} ${toNumber(amount).toFixed(2)}`;
}

function unitOptionsMarkup(selectedUnit) {
  const selected = String(selectedUnit || "Nos");
  const options = UNIT_OPTIONS.includes(selected) ? UNIT_OPTIONS : [...UNIT_OPTIONS, selected];
  return options
    .map((unit) => `<option value="${escapeHtml(unit)}" ${unit === selected ? "selected" : ""}>${escapeHtml(unit)}</option>`)
    .join("");
}

function computeTotals() {
  const subtotal = state.items.reduce((sum, item) => sum + itemAmount(item), 0);
  const discountAmount = subtotal * (toNumber(state.discount) / 100);
  const netAfterDiscount = subtotal - discountAmount;
  const taxAmount = netAfterDiscount * (toNumber(state.tax) / 100);
  const contingencyAmount = netAfterDiscount * (toNumber(state.contingency) / 100);
  const grand = netAfterDiscount + taxAmount + contingencyAmount;

  return { subtotal, discountAmount, taxAmount, contingencyAmount, grand };
}

function renderItems() {
  if (!state.items.length) {
    itemsTbody.innerHTML = '<tr><td colspan="8">No items. Click "Add Item".</td></tr>';
    renderTotals();
    return;
  }

  itemsTbody.innerHTML = state.items.map((item, i) => {
    const amount = itemAmount(item);
    return `
      <tr data-id="${item.id}">
        <td>${i + 1}</td>
        <td>
          <textarea class="cell-input desc" data-field="description" rows="3" maxlength="500">${escapeHtml(item.description)}</textarea>
        </td>
        <td>
          <select class="cell-input" data-field="unit">
            ${unitOptionsMarkup(item.unit)}
          </select>
        </td>
        <td>
          <input class="cell-input" data-field="qty" type="number" min="0" step="0.01" value="${item.qty}" />
        </td>
        <td>
          <input class="cell-input" data-field="rate" type="number" min="0" step="0.01" value="${item.rate}" />
        </td>
        <td>
          <input class="cell-input" data-field="waste" type="number" min="0" step="0.01" value="${item.waste}" />
        </td>
        <td class="amount-cell">${formatMoney(amount)}</td>
        <td>
          <button type="button" class="remove-row" data-action="remove">Remove</button>
        </td>
      </tr>
    `;
  }).join("");

  renderTotals();
}

function renderTotals() {
  const totals = computeTotals();
  subtotalValue.textContent = formatMoney(totals.subtotal);
  discountValue.textContent = formatMoney(totals.discountAmount);
  taxValue.textContent = formatMoney(totals.taxAmount);
  contingencyValue.textContent = formatMoney(totals.contingencyAmount);
  grandTotal.textContent = formatMoney(totals.grand);
}

function refreshVisibleRowAmounts() {
  const rows = itemsTbody.querySelectorAll("tr[data-id]");
  rows.forEach((row) => {
    const id = row.dataset.id;
    const item = state.items.find((x) => x.id === id);
    const amountCell = row.querySelector(".amount-cell");
    if (!item || !amountCell) return;
    amountCell.textContent = formatMoney(itemAmount(item));
  });
}

function syncHeaderAndSummaryState() {
  const typedClient = String(clientNameInput.value || "").trim();
  const pickedClient = clients.find((c) => c.name.toLowerCase() === typedClient.toLowerCase());
  const typedProject = String(projectNameInput.value || "").trim();
  const pickedProject = projects.find((p) => (
    p.name.toLowerCase() === typedProject.toLowerCase()
    && (!pickedClient || p.clientId === pickedClient.id)
  ));
  activeProjectId = pickedProject?.id || "";
  state.projectId = activeProjectId;
  state.projectName = pickedProject?.name || typedProject;
  state.clientName = pickedClient?.name || typedClient;
  state.preparedBy = preparedByInput.value.trim();
  state.date = boqDateInput.value;
  state.currency = currencySelect.value;
  state.discount = toNumber(discountInput.value);
  state.tax = toNumber(taxInput.value);
  state.contingency = toNumber(contingencyInput.value);
}

function applyStateToInputs() {
  const selectedProject = projects.find((p) => p.id === (state.projectId || ""));
  const selectedClientId = selectedProject?.clientId || clients.find((c) => c.name === state.clientName)?.id || "";
  const selectedClient = clients.find((c) => c.id === selectedClientId);
  renderEditorClientOptions();
  if (selectedClient) clientNameInput.value = selectedClient.name;
  else clientNameInput.value = state.clientName || "";
  renderEditorProjectOptions();
  if (state.projectId && projects.some((p) => p.id === state.projectId)) {
    projectNameInput.value = projectNameById(state.projectId);
  } else if (state.projectName) {
    const byName = projects.find((p) => p.name === state.projectName && (!selectedClientId || p.clientId === selectedClientId));
    if (byName) {
      projectNameInput.value = byName.name;
      activeProjectId = byName.id;
      state.projectId = byName.id;
    }
  } else {
    projectNameInput.value = "";
  }
  preparedByInput.value = state.preparedBy;
  boqDateInput.value = state.date;
  currencySelect.value = state.currency;
  discountInput.value = state.discount;
  taxInput.value = state.tax;
  contingencyInput.value = state.contingency;
}

function loadSavedBoqs() {
  try {
    const raw = JSON.parse(localStorage.getItem(SAVED_BOQS_KEY) || "[]");
    savedBoqs = Array.isArray(raw)
      ? raw.map((row) => ({
        id: row?.id || uid(),
        boqNumber: String(row?.boqNumber || "").trim(),
        savedAt: row?.savedAt || new Date().toISOString(),
        status: String(row?.status || "Saved"),
        negotiationNote: String(row?.negotiationNote || "").trim(),
        sourceId: String(row?.sourceId || ""),
        state: cloneState(row?.state || {})
      }))
      : [];
  } catch {
    savedBoqs = [];
  }
}

function persistSavedBoqs() {
  localStorage.setItem(SAVED_BOQS_KEY, JSON.stringify(savedBoqs));
}

function loadApprovedBoqs() {
  try {
    const raw = JSON.parse(localStorage.getItem(APPROVED_BOQS_KEY) || "[]");
    approvedBoqs = Array.isArray(raw)
      ? raw.map((row) => ({
        id: row?.id || uid(),
        boqNumber: String(row?.boqNumber || "").trim(),
        approvedAt: row?.approvedAt || new Date().toISOString(),
        approvedFromId: String(row?.approvedFromId || ""),
        status: "Approved",
        negotiationNote: String(row?.negotiationNote || "").trim(),
        sourceId: String(row?.sourceId || ""),
        state: cloneState(row?.state || {})
      }))
      : [];
  } catch {
    approvedBoqs = [];
  }
}

function persistApprovedBoqs() {
  localStorage.setItem(APPROVED_BOQS_KEY, JSON.stringify(approvedBoqs));
}

function loadClientProjectData() {
  try {
    const rawClients = JSON.parse(localStorage.getItem(CLIENTS_KEY) || "[]");
    clients = Array.isArray(rawClients) ? rawClients : [];
  } catch {
    clients = [];
  }
  try {
    const rawProjects = JSON.parse(localStorage.getItem(PROJECTS_KEY) || "[]");
    projects = Array.isArray(rawProjects)
      ? rawProjects.map((row) => ({
        id: row.id || uid(),
        name: String(row.name || "").trim(),
        clientId: row.clientId || "",
        notes: normalizeTextList(row.notes),
        remarks: normalizeTextList(row.remarks)
      })).filter((p) => p.name && p.clientId)
      : [];
  } catch {
    projects = [];
  }
}

function persistClientProjectData() {
  localStorage.setItem(CLIENTS_KEY, JSON.stringify(clients));
  localStorage.setItem(PROJECTS_KEY, JSON.stringify(projects));
}

function clientNameById(clientId) {
  const client = clients.find((c) => c.id === clientId);
  return client ? client.name : "-";
}

function projectNameById(projectId) {
  const project = projects.find((p) => p.id === projectId);
  return project ? project.name : "-";
}

function normalizeTextList(value) {
  if (Array.isArray(value)) {
    return value.map((v) => String(v || "").trim()).filter(Boolean);
  }
  const single = String(value || "").trim();
  return single ? [single] : [];
}

function renderDraftProjectMetaLists() {
  const noteRows = draftProjectNotes
    .map((text, i) => `<button type="button" class="ghost mini" data-kind="note" data-index="${i}" title="Remove">${escapeHtml(text)} ×</button>`)
    .join("");
  cpNoteList.innerHTML = noteRows || '<span class="muted">No notes added.</span>';

  const remarkRows = draftProjectRemarks
    .map((text, i) => `<button type="button" class="ghost mini" data-kind="remark" data-index="${i}" title="Remove">${escapeHtml(text)} ×</button>`)
    .join("");
  cpRemarkList.innerHTML = remarkRows || '<span class="muted">No remarks added.</span>';
}

function renderClientProjectOptions() {
  const options = clients
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map((client) => `<option value="${client.id}">${escapeHtml(client.name)}</option>`)
    .join("");
  cpProjectClientSelect.innerHTML = `<option value="">Select client</option>${options}`;
}

function renderClientTable() {
  if (!clients.length) {
    cpClientTbody.innerHTML = '<tr><td colspan="3">No clients added.</td></tr>';
    return;
  }
  const rows = clients
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map((client) => {
      const totalProjects = projects.filter((p) => p.clientId === client.id).length;
      return `
        <tr data-id="${client.id}">
          <td>${escapeHtml(client.name)}</td>
          <td>${totalProjects}</td>
          <td class="row-actions">
            <button type="button" class="danger" data-action="delete-client">Delete</button>
          </td>
        </tr>
      `;
    }).join("");
  cpClientTbody.innerHTML = rows;
}

function renderProjectTable() {
  if (!projects.length) {
    cpProjectTbody.innerHTML = '<tr><td colspan="3">No projects added.</td></tr>';
    return;
  }
  const rows = projects
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map((project) => `
      <tr data-id="${project.id}">
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(clientNameById(project.clientId))}</td>
        <td class="row-actions">
          <button type="button" class="secondary" data-action="open-project">Open In BOQ</button>
          <button type="button" class="danger" data-action="delete-project">Delete</button>
        </td>
      </tr>
    `).join("");
  cpProjectTbody.innerHTML = rows;
}

function renderClientProjectPage() {
  renderClientProjectOptions();
  renderClientTable();
  renderProjectTable();
  renderEditorClientOptions();
  renderEditorProjectOptions();
  renderSavedFilterClients();
  renderSavedFilterProjects();
}

function renderEditorClientOptions() {
  const q = String(clientNameInput?.value || "").trim();
  const names = clients
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map((c) => c.name);
  const items = topSuggestedValues("clients", names, q);
  renderSuggestList(clientNameSuggestions, items, "client");
}

function renderEditorProjectOptions() {
  const typedClient = String(clientNameInput?.value || "").trim().toLowerCase();
  const client = clients.find((c) => c.name.toLowerCase() === typedClient);
  const q = String(projectNameInput?.value || "").trim();
  const names = projects
    .filter((p) => (!client || p.clientId === client.id))
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map((p) => p.name);
  const items = topSuggestedValues("projects", names, q);
  renderSuggestList(projectNameSuggestions, items, "project");
}

function showPage(name) {
  if (name !== "login" && !isAuthenticated()) {
    userCorner?.classList.add("hidden");
    loginPage.classList.remove("hidden");
    homePage.classList.add("hidden");
    clientProjectPage.classList.add("hidden");
    editorPage.classList.add("hidden");
    savedBoqPage.classList.add("hidden");
    settingsPage.classList.add("hidden");
    return;
  }
  loginPage.classList.toggle("hidden", name !== "login");
  homePage.classList.toggle("hidden", name !== "home");
  clientProjectPage.classList.toggle("hidden", name !== "client-project");
  editorPage.classList.toggle("hidden", name !== "editor");
  savedBoqPage.classList.toggle("hidden", name !== "saved");
  settingsPage.classList.toggle("hidden", name !== "settings");
  adminUserSection?.classList.toggle("hidden", !isAdminUser());
  if (name === "settings" && isAdminUser()) renderUserTable();
  closeUserMenu();
  refreshUserCorner();
}

function refreshProjectNotesView() {
  const typedProject = String(projectNameInput?.value || "").trim().toLowerCase();
  const typedClient = String(clientNameInput?.value || "").trim().toLowerCase();
  const project = projects.find((p) => (
    p.id === activeProjectId
    || (typedProject && p.name.toLowerCase() === typedProject && (!typedClient || clientNameById(p.clientId).toLowerCase() === typedClient))
  ));
  if (!project) {
    viewProjectNotesBtn?.classList.add("hidden");
    projectNotesPanel?.classList.add("hidden");
    if (projectNotesText) projectNotesText.textContent = "-";
    if (projectRemarksText) projectRemarksText.textContent = "-";
    return;
  }
  viewProjectNotesBtn?.classList.remove("hidden");
  if (projectNotesText) {
    const notes = normalizeTextList(project.notes);
    projectNotesText.textContent = notes.length ? notes.map((n, i) => `${i + 1}. ${n}`).join("\n") : "-";
  }
  if (projectRemarksText) {
    const remarks = normalizeTextList(project.remarks);
    projectRemarksText.textContent = remarks.length ? remarks.map((n, i) => `${i + 1}. ${n}`).join("\n") : "-";
  }
}

function resetEditorStatus() {
  setStatus(summaryStatus, "", "ok");
  setStatus(itemStatus, "", "ok");
}

function startNewBoq() {
  activeSavedBoqId = null;
  activeBoqNumber = "";
  activeProjectId = "";
  state = blankState();
  applyStateToInputs();
  refreshPreparedBy();
  renderItems();
  refreshProjectNotesView();
  resetEditorStatus();
  showPage("editor");
}

function applyState(stateInput, source = "Saved BOQ") {
  state = cloneState(stateInput);
  activeProjectId = state.projectId || "";
  applyStateToInputs();
  refreshPreparedBy();
  renderItems();
  refreshProjectNotesView();
  setStatus(summaryStatus, `${source} loaded.`, "ok");
}

function boqNumberValue(value) {
  const text = String(value || "").trim().toUpperCase();
  const match = text.match(/BOQ-(\d+)/);
  return match ? Number(match[1]) : 0;
}

function formatBoqNumber(value) {
  return `BOQ-${String(Math.max(1, Number(value) || 1)).padStart(4, "0")}`;
}

function nextBoqNumber() {
  const maxSaved = savedBoqs.reduce((max, row) => Math.max(max, boqNumberValue(row?.boqNumber)), 0);
  const maxApproved = approvedBoqs.reduce((max, row) => Math.max(max, boqNumberValue(row?.boqNumber)), 0);
  return formatBoqNumber(Math.max(maxSaved, maxApproved) + 1);
}

function projectKeyFromState(boqState) {
  const s = cloneState(boqState);
  if (s.projectId) return `id:${s.projectId}`;
  return `name:${normalizeKey(s.clientName)}|${normalizeKey(s.projectName)}`;
}

function openSavedBoqPage() {
  renderSavedFilterClients();
  renderSavedFilterProjects();
  renderSavedPreparedByFilter();
  renderSavedBoqList();
  renderApprovedBoqList();
  showPage("saved");
}

function renderSavedPreparedByFilter() {
  if (!savedPreparedByFilter) return;
  const prev = savedPreparedByFilter.value;
  const names = Array.from(
    new Set(
      [...savedBoqs, ...approvedBoqs]
        .map((row) => String(row?.state?.preparedBy || "").trim())
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));
  savedPreparedByFilter.innerHTML = `<option value="">All</option>${names
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("")}`;
  if (names.includes(prev)) savedPreparedByFilter.value = prev;
}

function renderSavedFilterClients() {
  const rows = clients
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));
  if (!rows.length) {
    savedClientButtons.innerHTML = "<p>No clients found.</p>";
    savedSelectedClientId = "";
    return;
  }
  if (!rows.some((c) => c.id === savedSelectedClientId)) {
    savedSelectedClientId = rows[0].id;
  }
  savedClientButtons.innerHTML = rows
    .map((c) => `<button type="button" class="${c.id === savedSelectedClientId ? "" : "secondary"}" data-action="pick-client" data-id="${c.id}">${escapeHtml(c.name)}</button>`)
    .join("");
}

function renderSavedFilterProjects() {
  const rows = projects
    .filter((p) => !savedSelectedClientId || p.clientId === savedSelectedClientId)
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));
  if (!rows.length) {
    savedProjectButtons.innerHTML = "<p>No projects under selected client.</p>";
    savedSelectedProjectId = "";
    return;
  }
  if (!rows.some((p) => p.id === savedSelectedProjectId)) {
    savedSelectedProjectId = rows[0].id;
  }
  savedProjectButtons.innerHTML = rows
    .map((p) => `<button type="button" class="${p.id === savedSelectedProjectId ? "" : "secondary"}" data-action="pick-project" data-id="${p.id}">${escapeHtml(p.name)}</button>`)
    .join("");
}

function openClientProjectPage() {
  renderClientProjectPage();
  setStatus(cpStatus, "", "ok");
  showPage("client-project");
}

function openSettingsPage(from = "home") {
  settingsReturnPage = from;
  setStatus(settingsStatus, "", "ok");
  showPage("settings");
}

function buildBackupPayload() {
  return {
    backupVersion: 1,
    generatedAt: new Date().toISOString(),
    users: users.map((u) => ({
      username: String(u.username || "").trim(),
      password: String(u.password || ""),
      firstName: String(u.firstName || "").trim(),
      lastName: String(u.lastName || "").trim(),
      role: u.role === "admin" ? "admin" : "user"
    })),
    clients: clients.map((c) => ({ id: c.id || uid(), name: String(c.name || "") })),
    projects: projects.map((p) => ({
      id: p.id || uid(),
      name: String(p.name || ""),
      clientId: p.clientId || "",
      notes: normalizeTextList(p.notes),
      remarks: normalizeTextList(p.remarks)
    })),
    savedBoqs: savedBoqs.map((entry) => ({
      id: entry.id || uid(),
      boqNumber: String(entry.boqNumber || "").trim(),
      savedAt: entry.savedAt || new Date().toISOString(),
      status: String(entry.status || "Saved"),
      negotiationNote: String(entry.negotiationNote || "").trim(),
      sourceId: String(entry.sourceId || ""),
      state: cloneState(entry.state)
    })),
    approvedBoqs: approvedBoqs.map((entry) => ({
      id: entry.id || uid(),
      boqNumber: String(entry.boqNumber || "").trim(),
      approvedAt: entry.approvedAt || new Date().toISOString(),
      approvedFromId: String(entry.approvedFromId || ""),
      negotiationNote: String(entry.negotiationNote || "").trim(),
      sourceId: String(entry.sourceId || ""),
      state: cloneState(entry.state)
    }))
  };
}

function downloadBackup() {
  const payload = buildBackupPayload();
  const stamp = new Date().toISOString().replaceAll(":", "-");
  downloadFile(`boq-backup-${stamp}.json`, JSON.stringify(payload, null, 2), "application/json");
  const total = payload.savedBoqs.length + payload.approvedBoqs.length;
  setStatus(settingsStatus, `Backup downloaded (${total} BOQs: ${payload.savedBoqs.length} saved, ${payload.approvedBoqs.length} approved).`, "ok");
}

async function restoreBackupFromFile() {
  const file = backupFileInput.files?.[0];
  if (!file) {
    setStatus(settingsStatus, "Select a backup JSON file first.", "err");
    return;
  }

  try {
    const raw = JSON.parse(await file.text());
    if (Array.isArray(raw.users)) {
      users = raw.users
        .map((u) => ({
          username: String(u?.username || "").trim(),
          password: String(u?.password || ""),
          firstName: String(u?.firstName || "").trim(),
          lastName: String(u?.lastName || "").trim(),
          role: String(u?.role || "").toLowerCase() === "admin" ? "admin" : "user"
        }))
        .filter((u) => u.username && u.password);
      seedDefaultUser();
    }
    const nextClients = Array.isArray(raw.clients)
      ? raw.clients.map((row) => ({ id: row.id || uid(), name: String(row.name || "").trim() })).filter((c) => c.name)
      : [];
    const validClientIds = new Set(nextClients.map((c) => c.id));
    const nextProjects = Array.isArray(raw.projects)
      ? raw.projects
        .map((row) => ({
          id: row.id || uid(),
          name: String(row.name || "").trim(),
          clientId: row.clientId || "",
          notes: normalizeTextList(row.notes),
          remarks: normalizeTextList(row.remarks)
        }))
        .filter((p) => p.name && validClientIds.has(p.clientId))
      : [];
    const rows = Array.isArray(raw.savedBoqs) ? raw.savedBoqs : [];
    const nextSaved = rows.map((row) => ({
      id: row.id || uid(),
      boqNumber: String(row.boqNumber || "").trim(),
      savedAt: row.savedAt || new Date().toISOString(),
      status: String(row.status || "Saved"),
      negotiationNote: String(row.negotiationNote || "").trim(),
      sourceId: String(row.sourceId || ""),
      state: cloneState(row.state || {})
    }));
    const approvedRows = Array.isArray(raw.approvedBoqs) ? raw.approvedBoqs : [];
    const nextApproved = approvedRows.map((row) => ({
      id: row.id || uid(),
      boqNumber: String(row.boqNumber || "").trim(),
      approvedAt: row.approvedAt || new Date().toISOString(),
      approvedFromId: String(row.approvedFromId || ""),
      status: "Approved",
      negotiationNote: String(row.negotiationNote || "").trim(),
      sourceId: String(row.sourceId || ""),
      state: cloneState(row.state || {})
    }));
    clients = nextClients;
    projects = nextProjects;
    savedBoqs = nextSaved;
    approvedBoqs = nextApproved;
    activeSavedBoqId = null;
    persistClientProjectData();
    persistSavedBoqs();
    persistApprovedBoqs();
    renderClientProjectPage();
    renderSavedPreparedByFilter();
    renderSavedBoqList();
    renderApprovedBoqList();
    setStatus(settingsStatus, `Backup restored (${savedBoqs.length} saved, ${approvedBoqs.length} approved, ${clients.length} clients, ${projects.length} projects).`, "ok");
  } catch (error) {
    setStatus(settingsStatus, `Restore failed: ${error.message}`, "err");
  } finally {
    backupFileInput.value = "";
  }
}

function matchesSavedFilters(boqState) {
  const selectedClientId = savedSelectedClientId;
  const selectedProjectId = savedSelectedProjectId;
  const selectedPreparedBy = String(savedPreparedByFilter?.value || "").trim().toLowerCase();
  const q = String(savedSearchInput?.value || "").trim().toLowerCase();
  const s = cloneState(boqState);
  const project = projects.find((p) => p.id === (s.projectId || ""));
  const clientId = project?.clientId || clients.find((c) => c.name === s.clientName)?.id || "";
  if (selectedClientId && clientId !== selectedClientId) return false;
  if (selectedProjectId) {
    const byId = (s.projectId || "") === selectedProjectId;
    const byName = s.projectName.toLowerCase() === projectNameById(selectedProjectId).toLowerCase();
    if (!byId && !byName) return false;
  }
  if (selectedPreparedBy && String(s.preparedBy || "").trim().toLowerCase() !== selectedPreparedBy) return false;
  if (!q) return true;
  const hay = `${s.projectName} ${s.clientName} ${s.preparedBy}`.toLowerCase();
  return hay.includes(q);
}

function boqGrandTotal(boqState) {
  const s = cloneState(boqState);
  const subtotal = s.items.reduce((sum, item) => sum + itemAmount(item), 0);
  const discountAmount = subtotal * (toNumber(s.discount) / 100);
  const netAfterDiscount = subtotal - discountAmount;
  const taxAmount = netAfterDiscount * (toNumber(s.tax) / 100);
  const contingencyAmount = netAfterDiscount * (toNumber(s.contingency) / 100);
  return netAfterDiscount + taxAmount + contingencyAmount;
}

function renderSavedBoqList() {
  if (!savedBoqs.length) {
    savedBoqTbody.innerHTML = '<tr><td colspan="10">No saved BOQs found.</td></tr>';
    setStatus(savedBoqStatus, "Save a BOQ from the editor to view it here.", "err");
    return;
  }

  const rows = [...savedBoqs]
    .filter((entry) => matchesSavedFilters(entry.state))
    .sort((a, b) => new Date(b.savedAt || 0) - new Date(a.savedAt || 0));

  if (!rows.length) {
    savedBoqTbody.innerHTML = '<tr><td colspan="10">No BOQs for selected filters.</td></tr>';
    setStatus(savedBoqStatus, "No matching BOQs found.", "err");
    return;
  }

  savedBoqTbody.innerHTML = rows.map((entry) => {
    const boqState = cloneState(entry.state);
    const itemCount = boqState.items.length;
    const total = boqGrandTotal(boqState);
    const savedAt = entry.savedAt ? new Date(entry.savedAt).toLocaleString() : "-";
    const status = String(entry.status || "Saved");
    const negotiationText = String(entry.negotiationNote || "").trim();
    const boqNumber = String(entry.boqNumber || "-");
    return `
      <tr data-id="${entry.id}">
        <td>${escapeHtml(boqNumber)}</td>
        <td>${escapeHtml(getBoqLabel(boqState))}</td>
        <td>${escapeHtml(boqState.clientName || "-")}</td>
        <td>${escapeHtml(boqState.date || "-")}</td>
        <td>${escapeHtml(savedAt)}</td>
        <td>${itemCount}</td>
        <td><span class="status-pill">${escapeHtml(status)}</span></td>
        <td>${escapeHtml(negotiationText || "-")}</td>
        <td>₹ ${total.toFixed(2)}</td>
        <td class="row-actions">
          <button type="button" class="secondary" data-action="open">Open</button>
          <button type="button" class="secondary" data-action="negotiate">Negotiate</button>
          <button type="button" class="secondary" data-action="approve">Approve</button>
          <button type="button" class="danger" data-action="delete">Delete</button>
        </td>
      </tr>
    `;
  }).join("");
  setStatus(savedBoqStatus, `${rows.length} BOQ(s) found.`, "ok");
}

function renderApprovedBoqList() {
  if (!approvedBoqTbody || !approvedBoqStatus) return;
  if (!approvedBoqs.length) {
    approvedBoqTbody.innerHTML = '<tr><td colspan="8">No approved BOQs found.</td></tr>';
    setStatus(approvedBoqStatus, "Approved BOQs will appear here.", "err");
    return;
  }

  const rows = [...approvedBoqs]
    .filter((entry) => matchesSavedFilters(entry.state))
    .sort((a, b) => new Date(b.approvedAt || 0) - new Date(a.approvedAt || 0));

  if (!rows.length) {
    approvedBoqTbody.innerHTML = '<tr><td colspan="8">No approved BOQs for selected filters.</td></tr>';
    setStatus(approvedBoqStatus, "No matching approved BOQs found.", "err");
    return;
  }

  approvedBoqTbody.innerHTML = rows.map((entry) => {
    const boqState = cloneState(entry.state);
    const approvedAt = entry.approvedAt ? new Date(entry.approvedAt).toLocaleString() : "-";
    const total = boqGrandTotal(boqState);
    const boqNumber = String(entry.boqNumber || "-");
    return `
      <tr data-id="${entry.id}">
        <td>${escapeHtml(boqNumber)}</td>
        <td>${escapeHtml(getBoqLabel(boqState))}</td>
        <td>${escapeHtml(boqState.clientName || "-")}</td>
        <td>${escapeHtml(boqState.date || "-")}</td>
        <td>${escapeHtml(approvedAt)}</td>
        <td>${escapeHtml(boqState.preparedBy || "-")}</td>
        <td>₹ ${total.toFixed(2)}</td>
        <td class="row-actions">
          <button type="button" class="secondary" data-action="open-approved">Open</button>
          <button type="button" class="danger" data-action="delete-approved">Delete</button>
        </td>
      </tr>
    `;
  }).join("");
  setStatus(approvedBoqStatus, `${rows.length} approved BOQ(s) found.`, "ok");
}

function saveState(asNew = false) {
  syncHeaderAndSummaryState();
  rememberSuggestion("clients", state.clientName, 2);
  rememberSuggestion("projects", state.projectName, 2);
  persistSuggestMemory();
  const snapshot = cloneState(state);
  const nowIso = new Date().toISOString();
  if (asNew) {
    activeSavedBoqId = null;
    activeBoqNumber = "";
  }

  if (activeSavedBoqId) {
    const idx = savedBoqs.findIndex((x) => x.id === activeSavedBoqId);
    if (idx >= 0) {
      const currentNumber = savedBoqs[idx].boqNumber || activeBoqNumber || nextBoqNumber();
      savedBoqs[idx] = {
        ...savedBoqs[idx],
        boqNumber: currentNumber,
        state: snapshot,
        savedAt: nowIso,
        status: savedBoqs[idx].status || "Saved"
      };
      activeBoqNumber = currentNumber;
    } else {
      activeSavedBoqId = null;
      activeBoqNumber = "";
    }
  }

  if (!activeSavedBoqId) {
    const id = uid();
    const boqNumber = activeBoqNumber || nextBoqNumber();
    savedBoqs.push({
      id,
      boqNumber,
      state: snapshot,
      savedAt: nowIso,
      status: "Saved",
      negotiationNote: "",
      sourceId: ""
    });
    activeSavedBoqId = id;
    activeBoqNumber = boqNumber;
  }

  persistSavedBoqs();
  renderSavedPreparedByFilter();
  renderSavedBoqList();
  setStatus(summaryStatus, `${asNew ? "Saved as new" : "Saved"}: ${getBoqLabel(state)}.`, "ok");
}

function negotiateSavedBoq(id) {
  const entry = savedBoqs.find((row) => row.id === id);
  if (!entry) return;
  const noteInput = window.prompt("Enter client negotiation note:", entry.negotiationNote || "");
  if (noteInput === null) return;
  const negotiationNote = String(noteInput || "").trim();
  const nowIso = new Date().toISOString();
  const newEntry = {
    id: uid(),
    boqNumber: String(entry.boqNumber || nextBoqNumber()),
    state: cloneState(entry.state),
    savedAt: nowIso,
    status: "Negotiation",
    negotiationNote: negotiationNote || "Client negotiation requested.",
    sourceId: entry.id
  };
  savedBoqs.push(newEntry);
  activeSavedBoqId = newEntry.id;
  activeBoqNumber = newEntry.boqNumber;
  persistSavedBoqs();
  renderSavedPreparedByFilter();
  renderSavedBoqList();
  setStatus(savedBoqStatus, "Negotiated BOQ version created.", "ok");
}

function approveSavedBoq(id) {
  const entry = savedBoqs.find((row) => row.id === id);
  if (!entry) return;
  if (!window.confirm("Approve this BOQ and move it to Approved BOQs?")) return;
  const approvedAtIso = new Date().toISOString();
  const approvedProjectKey = projectKeyFromState(entry.state);
  const approvedBoqNumber = String(entry.boqNumber || nextBoqNumber());
  const approvedEntry = {
    id: uid(),
    boqNumber: approvedBoqNumber,
    approvedAt: approvedAtIso,
    approvedFromId: entry.id,
    status: "Approved",
    negotiationNote: String(entry.negotiationNote || "").trim(),
    sourceId: String(entry.sourceId || entry.id),
    state: cloneState(entry.state)
  };
  approvedBoqs.push(approvedEntry);
  savedBoqs = savedBoqs
    .map((row) => {
      if (row.id === id) return null;
      if (projectKeyFromState(row.state) !== approvedProjectKey) return row;
      return { ...row, status: "Rejected", negotiationNote: row.negotiationNote || "Rejected after another BOQ was approved." };
    })
    .filter(Boolean);
  if (activeSavedBoqId === id) activeSavedBoqId = null;
  if (activeBoqNumber === approvedBoqNumber) activeBoqNumber = "";
  persistSavedBoqs();
  persistApprovedBoqs();
  renderSavedPreparedByFilter();
  renderSavedBoqList();
  renderApprovedBoqList();
  setStatus(savedBoqStatus, "BOQ approved and moved to Approved BOQs.", "ok");
}

function toCsvRow(values) {
  return values.map((v) => `"${String(v ?? "").replaceAll('"', '""')}"`).join(",");
}

function exportCsv() {
  syncHeaderAndSummaryState();
  const totals = computeTotals();

  const rows = [
    toCsvRow(["Project Name", state.projectName]),
    toCsvRow(["Client", state.clientName]),
    toCsvRow(["Prepared By", state.preparedBy]),
    toCsvRow(["Date", state.date]),
    toCsvRow(["Currency", state.currency]),
    "",
    toCsvRow(["#", "Description", "Unit", "Qty", "Rate", "Waste %", "Amount"])
  ];

  state.items.forEach((item, i) => {
    rows.push(toCsvRow([
      i + 1,
      item.description,
      item.unit,
      item.qty,
      item.rate,
      item.waste,
      itemAmount(item).toFixed(2)
    ]));
  });

  rows.push("");
  rows.push(toCsvRow(["Subtotal", totals.subtotal.toFixed(2)]));
  rows.push(toCsvRow(["Discount Amount", totals.discountAmount.toFixed(2)]));
  rows.push(toCsvRow(["GST Amount", totals.taxAmount.toFixed(2)]));
  rows.push(toCsvRow(["Contingency Amount", totals.contingencyAmount.toFixed(2)]));
  rows.push(toCsvRow(["Grand Total", totals.grand.toFixed(2)]));

  downloadFile(`boq-${state.date || new Date().toISOString().slice(0, 10)}.csv`, rows.join("\n"), "text/csv");
  setStatus(summaryStatus, "CSV exported.", "ok");
}

function exportJson() {
  syncHeaderAndSummaryState();
  downloadFile(`boq-${state.date || new Date().toISOString().slice(0, 10)}.json`, JSON.stringify({
    ...state,
    totals: computeTotals(),
    exportedAt: new Date().toISOString()
  }, null, 2), "application/json");
  setStatus(summaryStatus, "JSON exported.", "ok");
}

function printBoq() {
  syncHeaderAndSummaryState();
  const totals = computeTotals();
  const symbol = currencySymbol(state.currency);
  const logoUrl = new URL("assets/sribal-logo.png", window.location.href).href;
  const rowsHtml = state.items.length
    ? state.items.map((item, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${escapeHtml(item.description || "-")}</td>
        <td>${escapeHtml(item.unit || "-")}</td>
        <td>${toNumber(item.qty).toFixed(2)}</td>
        <td>${toNumber(item.rate).toFixed(2)}</td>
        <td>${toNumber(item.waste).toFixed(2)}</td>
        <td>${itemAmount(item).toFixed(2)}</td>
      </tr>
    `).join("")
    : '<tr><td colspan="7">No BOQ items.</td></tr>';

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    setStatus(summaryStatus, "Pop-up blocked. Allow pop-ups to print BOQ.", "err");
    return;
  }

  const html = `
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="UTF-8" />
      <title>BOQ Print</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 24px; color: #111; }
        .head { display: flex; justify-content: space-between; align-items: center; gap: 16px; margin-bottom: 14px; }
        .head h1 { margin: 0; font-size: 24px; }
        .meta { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px 18px; margin: 14px 0; }
        .meta p { margin: 0; font-size: 13px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #333; padding: 7px; font-size: 12px; text-align: left; vertical-align: top; }
        th { background: #f2f2f2; }
        .totals { margin-top: 14px; width: 320px; margin-left: auto; }
        .totals p { display: flex; justify-content: space-between; margin: 0; border: 1px solid #333; border-top: none; padding: 7px; font-size: 12px; }
        .totals p:first-child { border-top: 1px solid #333; }
        .totals p.grand { font-weight: 700; font-size: 13px; }
        .logo { width: 220px; max-width: 40%; height: auto; }
        @media print { body { margin: 10mm; } }
      </style>
    </head>
    <body>
      <div class="head">
        <div>
          <h1>Bill Of Quantities (BOQ)</h1>
          <p>Generated on: ${new Date().toLocaleString()}</p>
        </div>
        <img class="logo" src="${logoUrl}" alt="Sribal" />
      </div>

      <div class="meta">
        <p><strong>Project:</strong> ${escapeHtml(state.projectName || "-")}</p>
        <p><strong>Client:</strong> ${escapeHtml(state.clientName || "-")}</p>
        <p><strong>Prepared By:</strong> ${escapeHtml(state.preparedBy || "-")}</p>
        <p><strong>Date:</strong> ${escapeHtml(state.date || "-")}</p>
        <p><strong>Currency:</strong> ${escapeHtml(state.currency || "INR")}</p>
      </div>

      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Description</th>
            <th>Unit</th>
            <th>Qty</th>
            <th>Rate</th>
            <th>Waste %</th>
            <th>Amount</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>

      <div class="totals">
        <p><span>Subtotal</span><strong>${symbol} ${totals.subtotal.toFixed(2)}</strong></p>
        <p><span>Discount Amount</span><strong>${symbol} ${totals.discountAmount.toFixed(2)}</strong></p>
        <p><span>GST Amount</span><strong>${symbol} ${totals.taxAmount.toFixed(2)}</strong></p>
        <p><span>Contingency Amount</span><strong>${symbol} ${totals.contingencyAmount.toFixed(2)}</strong></p>
        <p class="grand"><span>Grand Total</span><strong>${symbol} ${totals.grand.toFixed(2)}</strong></p>
      </div>
    </body>
    </html>
  `;

  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
  setStatus(summaryStatus, "Print dialog opened.", "ok");
}

function addClient() {
  const name = String(cpClientNameInput.value || "").trim();
  if (!name) {
    setStatus(cpStatus, "Client name is required.", "err");
    return;
  }
  const exists = clients.some((c) => c.name.toLowerCase() === name.toLowerCase());
  if (exists) {
    setStatus(cpStatus, "Client already exists.", "err");
    return;
  }
  clients.push({ id: uid(), name });
  rememberSuggestion("clients", name, 3);
  persistSuggestMemory();
  persistClientProjectData();
  cpClientNameInput.value = "";
  renderClientProjectPage();
  setStatus(cpStatus, `Client added: ${name}`, "ok");
}

function addProject() {
  const clientId = cpProjectClientSelect.value;
  const name = String(cpProjectNameInput.value || "").trim();
  const notes = [...draftProjectNotes];
  const remarks = [...draftProjectRemarks];
  if (!clientId) {
    setStatus(cpStatus, "Select a client.", "err");
    return;
  }
  if (!name) {
    setStatus(cpStatus, "Project name is required.", "err");
    return;
  }
  const exists = projects.some((p) => p.clientId === clientId && p.name.toLowerCase() === name.toLowerCase());
  if (exists) {
    setStatus(cpStatus, "Project already exists under this client.", "err");
    return;
  }
  projects.push({ id: uid(), name, clientId, notes, remarks });
  rememberSuggestion("projects", name, 3);
  persistSuggestMemory();
  persistClientProjectData();
  cpProjectNameInput.value = "";
  cpProjectNoteInput.value = "";
  cpProjectRemarkInput.value = "";
  draftProjectNotes = [];
  draftProjectRemarks = [];
  renderDraftProjectMetaLists();
  renderClientProjectPage();
  setStatus(cpStatus, `Project added: ${name}`, "ok");
}

function addDraftProjectNote() {
  const value = String(cpProjectNoteInput.value || "").trim();
  if (!value) return;
  draftProjectNotes.push(value);
  cpProjectNoteInput.value = "";
  renderDraftProjectMetaLists();
}

function addDraftProjectRemark() {
  const value = String(cpProjectRemarkInput.value || "").trim();
  if (!value) return;
  draftProjectRemarks.push(value);
  cpProjectRemarkInput.value = "";
  renderDraftProjectMetaLists();
}

function deleteClient(clientId) {
  const client = clients.find((c) => c.id === clientId);
  if (!client) return;
  if (!window.confirm(`Delete client "${client.name}" and all its projects?`)) return;
  const removedProjectIds = new Set(projects.filter((p) => p.clientId === clientId).map((p) => p.id));
  clients = clients.filter((c) => c.id !== clientId);
  projects = projects.filter((p) => p.clientId !== clientId);
  if (removedProjectIds.has(activeProjectId)) {
    activeProjectId = "";
    refreshProjectNotesView();
  }
  persistClientProjectData();
  renderClientProjectPage();
  setStatus(cpStatus, `Client deleted: ${client.name}`, "ok");
}

function deleteProject(projectId) {
  const project = projects.find((p) => p.id === projectId);
  if (!project) return;
  if (!window.confirm(`Delete project "${project.name}"?`)) return;
  projects = projects.filter((p) => p.id !== projectId);
  if (activeProjectId === projectId) {
    activeProjectId = "";
    refreshProjectNotesView();
  }
  persistClientProjectData();
  renderClientProjectPage();
  setStatus(cpStatus, `Project deleted: ${project.name}`, "ok");
}

function openProjectInEditor(projectId) {
  const project = projects.find((p) => p.id === projectId);
  if (!project) return;
  const client = clients.find((c) => c.id === project.clientId);
  startNewBoq();
  activeProjectId = project.id;
  renderEditorClientOptions();
  clientNameInput.value = client?.name || "";
  renderEditorProjectOptions();
  projectNameInput.value = project.name;
  rememberSuggestion("clients", client?.name || "", 1);
  rememberSuggestion("projects", project.name, 1);
  persistSuggestMemory();
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
  setStatus(summaryStatus, `Project loaded: ${project.name}`, "ok");
}

function downloadFile(name, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

addItemBtn.addEventListener("click", () => {
  addItem();
  renderItems();
  setStatus(itemStatus, "Item added.", "ok");
});

clearItemsBtn.addEventListener("click", () => {
  if (!window.confirm("Clear all BOQ items?")) return;
  state.items = [];
  renderItems();
  setStatus(itemStatus, "All items cleared.", "ok");
});

itemsTbody.addEventListener("input", (event) => {
  const input = event.target.closest("input[data-field], textarea[data-field]");
  if (!input) return;
  const row = input.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  const item = state.items.find((x) => x.id === id);
  if (!item) return;
  const field = input.dataset.field;

  if (field === "description" || field === "unit") item[field] = input.value;
  else item[field] = toNumber(input.value);

  const amountCell = row.querySelector(".amount-cell");
  if (amountCell) amountCell.textContent = formatMoney(itemAmount(item));
  renderTotals();
});

itemsTbody.addEventListener("change", (event) => {
  const select = event.target.closest("select[data-field='unit']");
  if (!select) return;
  const row = select.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  const item = state.items.find((x) => x.id === id);
  if (!item) return;
  item.unit = select.value;
});

itemsTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='remove']");
  if (!btn) return;
  const row = btn.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  state.items = state.items.filter((x) => x.id !== id);
  renderItems();
  setStatus(itemStatus, "Item removed.", "ok");
});

[projectNameInput, clientNameInput, preparedByInput, boqDateInput, currencySelect, discountInput, taxInput, contingencyInput]
  .forEach((el) => {
    el.addEventListener("input", () => {
      syncHeaderAndSummaryState();
      if (el === currencySelect) refreshVisibleRowAmounts();
      renderTotals();
    });
    el.addEventListener("change", () => {
      syncHeaderAndSummaryState();
      if (el === currencySelect) refreshVisibleRowAmounts();
      renderTotals();
    });
  });

clientNameInput?.addEventListener("input", () => {
  renderEditorClientOptions();
  renderEditorProjectOptions();
  const typedClient = String(clientNameInput.value || "").trim().toLowerCase();
  const selectedClient = clients.find((c) => c.name.toLowerCase() === typedClient);
  const selectedProject = projects.find((p) => p.name.toLowerCase() === String(projectNameInput.value || "").trim().toLowerCase());
  if (selectedProject && selectedClient && selectedProject.clientId !== selectedClient.id) {
    projectNameInput.value = "";
    activeProjectId = "";
  }
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
});

projectNameInput?.addEventListener("input", () => {
  renderEditorProjectOptions();
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
});

clientNameInput?.addEventListener("focus", renderEditorClientOptions);
projectNameInput?.addEventListener("focus", renderEditorProjectOptions);

clientNameSuggestions?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-field='client'][data-value]");
  if (!btn) return;
  clientNameInput.value = btn.dataset.value || "";
  clientNameSuggestions.classList.add("hidden");
  renderEditorProjectOptions();
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
});

projectNameSuggestions?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-field='project'][data-value]");
  if (!btn) return;
  projectNameInput.value = btn.dataset.value || "";
  projectNameSuggestions.classList.add("hidden");
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
});

document.addEventListener("click", (event) => {
  if (!clientNameInput?.contains(event.target) && !clientNameSuggestions?.contains(event.target)) {
    clientNameSuggestions?.classList.add("hidden");
  }
  if (!projectNameInput?.contains(event.target) && !projectNameSuggestions?.contains(event.target)) {
    projectNameSuggestions?.classList.add("hidden");
  }
});

projectNameInput?.addEventListener("change", () => {
  const selectedProject = projects.find((p) => p.name.toLowerCase() === String(projectNameInput.value || "").trim().toLowerCase());
  if (selectedProject) {
    clientNameInput.value = clientNameById(selectedProject.clientId);
    activeProjectId = selectedProject.id;
  } else {
    activeProjectId = "";
  }
  renderEditorProjectOptions();
  syncHeaderAndSummaryState();
  refreshProjectNotesView();
});

savedSearchInput?.addEventListener("input", () => {
  renderSavedBoqList();
  renderApprovedBoqList();
});
savedPreparedByFilter?.addEventListener("change", () => {
  renderSavedBoqList();
  renderApprovedBoqList();
});

saveBtn.addEventListener("click", () => saveState(false));
saveAsNewBtn?.addEventListener("click", () => saveState(true));
loadBtn.addEventListener("click", openSavedBoqPage);
openSavedTopBtn?.addEventListener("click", openSavedBoqPage);
printBoqBtn?.addEventListener("click", printBoq);
exportCsvBtn.addEventListener("click", exportCsv);
exportJsonBtn.addEventListener("click", exportJson);
homeCreateBoqBtn?.addEventListener("click", startNewBoq);
homeClientProjectsBtn?.addEventListener("click", openClientProjectPage);
homeSavedBoqsBtn?.addEventListener("click", openSavedBoqPage);
homeSettingsBtn?.addEventListener("click", () => openSettingsPage("home"));
cpHomeBtn?.addEventListener("click", () => showPage("home"));
cpAddClientBtn?.addEventListener("click", addClient);
cpAddProjectBtn?.addEventListener("click", addProject);
cpAddNoteBtn?.addEventListener("click", addDraftProjectNote);
cpAddRemarkBtn?.addEventListener("click", addDraftProjectRemark);
cpProjectNoteInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  addDraftProjectNote();
});
cpProjectRemarkInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  addDraftProjectRemark();
});
editorHomeBtn?.addEventListener("click", () => showPage("home"));
editorSettingsBtn?.addEventListener("click", () => openSettingsPage("editor"));
savedHomeBtn?.addEventListener("click", () => showPage("home"));
backToEditorBtn.addEventListener("click", () => showPage("editor"));
savedSettingsBtn?.addEventListener("click", () => openSettingsPage("saved"));
settingsBackBtn?.addEventListener("click", () => showPage(settingsReturnPage));
settingsHomeBtn?.addEventListener("click", () => showPage("home"));
downloadBackupBtn?.addEventListener("click", downloadBackup);
restoreBackupBtn?.addEventListener("click", restoreBackupFromFile);
adminAddUserBtn?.addEventListener("click", addUserByAdmin);
viewProjectNotesBtn?.addEventListener("click", () => {
  projectNotesPanel?.classList.toggle("hidden");
});
userLogoutBtn?.addEventListener("click", logout);
userPhotoBtn?.addEventListener("click", () => userPhotoInput?.click());
userMenuTrigger?.addEventListener("click", () => {
  userMenu?.classList.toggle("hidden");
});
userPhotoInput?.addEventListener("change", async () => {
  const file = userPhotoInput.files?.[0];
  if (!file || !sessionUser) return;
  try {
    const dataUrl = await readFileAsDataUrl(file);
    if (!userProfiles[sessionUser]) userProfiles[sessionUser] = {};
    userProfiles[sessionUser].photo = dataUrl;
    persistUserProfiles();
    refreshUserCorner();
  } catch {
    // non-blocking UI operation
  } finally {
    userPhotoInput.value = "";
  }
});
document.addEventListener("click", (event) => {
  if (!userCorner?.contains(event.target)) closeUserMenu();
});
loginForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  const ok = tryLogin(loginUsernameInput.value, loginPasswordInput.value);
  if (!ok) {
    setStatus(loginStatus, "Invalid username or password.", "err");
    return;
  }
  setStatus(loginStatus, "Login successful.", "ok");
  showPage("home");
});

savedBoqTbody.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;
  const row = btn.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  const entry = savedBoqs.find((x) => x.id === id);
  if (!entry) return;

  if (btn.dataset.action === "open") {
    activeSavedBoqId = id;
    activeBoqNumber = String(entry.boqNumber || "");
    applyState(entry.state, `Loaded ${getBoqLabel(entry.state)}`);
    showPage("editor");
    return;
  }

  if (btn.dataset.action === "negotiate") {
    negotiateSavedBoq(id);
    return;
  }

  if (btn.dataset.action === "approve") {
    approveSavedBoq(id);
    return;
  }

  if (btn.dataset.action === "delete") {
    if (!window.confirm("Delete this saved BOQ?")) return;
    savedBoqs = savedBoqs.filter((x) => x.id !== id);
    if (activeSavedBoqId === id) activeSavedBoqId = null;
    persistSavedBoqs();
    renderSavedPreparedByFilter();
    renderSavedBoqList();
    renderApprovedBoqList();
  }
});

approvedBoqTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;
  const row = btn.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  const entry = approvedBoqs.find((x) => x.id === id);
  if (!entry) return;

  if (btn.dataset.action === "open-approved") {
    activeSavedBoqId = null;
    activeBoqNumber = "";
    applyState(entry.state, `Loaded approved ${getBoqLabel(entry.state)}`);
    showPage("editor");
    return;
  }

  if (btn.dataset.action === "delete-approved") {
    if (!window.confirm("Delete this approved BOQ?")) return;
    approvedBoqs = approvedBoqs.filter((x) => x.id !== id);
    persistApprovedBoqs();
    renderSavedPreparedByFilter();
    renderApprovedBoqList();
  }
});

savedClientButtons?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='pick-client'][data-id]");
  if (!btn) return;
  savedSelectedClientId = btn.dataset.id || "";
  savedSelectedProjectId = "";
  renderSavedFilterClients();
  renderSavedFilterProjects();
  renderSavedBoqList();
  renderApprovedBoqList();
});

savedProjectButtons?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='pick-project'][data-id]");
  if (!btn) return;
  savedSelectedProjectId = btn.dataset.id || "";
  renderSavedFilterProjects();
  renderSavedBoqList();
  renderApprovedBoqList();
});

cpClientTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action='delete-client']");
  if (!btn) return;
  const row = btn.closest("tr[data-id]");
  if (!row) return;
  deleteClient(row.dataset.id);
});

cpProjectTbody?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;
  const row = btn.closest("tr[data-id]");
  if (!row) return;
  const id = row.dataset.id;
  if (btn.dataset.action === "open-project") {
    openProjectInEditor(id);
    return;
  }
  if (btn.dataset.action === "delete-project") {
    deleteProject(id);
  }
});

cpNoteList?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-kind='note'][data-index]");
  if (!btn) return;
  const index = Number(btn.dataset.index);
  if (!Number.isInteger(index) || index < 0) return;
  draftProjectNotes = draftProjectNotes.filter((_, i) => i !== index);
  renderDraftProjectMetaLists();
});

cpRemarkList?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-kind='remark'][data-index]");
  if (!btn) return;
  const index = Number(btn.dataset.index);
  if (!Number.isInteger(index) || index < 0) return;
  draftProjectRemarks = draftProjectRemarks.filter((_, i) => i !== index);
  renderDraftProjectMetaLists();
});

loadSavedBoqs();
loadApprovedBoqs();
loadClientProjectData();
loadSuggestMemory();
renderClientProjectPage();
renderDraftProjectMetaLists();
seedDefaultUser();
loadUserProfiles();
sessionUser = localStorage.getItem(SESSION_KEY) || "";
state = blankState();
applyStateToInputs();
refreshPreparedBy();
renderItems();
showPage(isAuthenticated() ? "home" : "login");
