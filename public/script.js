const productForm = document.getElementById("productForm");
const editingIdInput = document.getElementById("editingId");

const categorySuggestions = document.getElementById("categorySuggestions");
const clientSuggestions = document.getElementById("clientSuggestions");

let categoryHistory = [];
let clientHistory = [];

const itemNo = document.getElementById("itemNo");
const itemName = document.getElementById("itemName");
const category = document.getElementById("category");
const client = document.getElementById("client");

const capitalCurrency = document.getElementById("capitalCurrency");
const capitalPrice = document.getElementById("capitalPrice");

const wholesaleCurrency = document.getElementById("wholesaleCurrency");
const wholesalePrice = document.getElementById("wholesalePrice");

const descEn = document.getElementById("descEn");
const descLa = document.getElementById("descLa");

const photosInput = document.getElementById("photos");
const photoPreview = document.getElementById("photoPreview");

const saveBtn = document.getElementById("saveBtn");
const clearBtn = document.getElementById("clearBtn");
const downloadExcelBtn = document.getElementById("downloadExcelBtn");

const searchInput = document.getElementById("searchInput");
const categoryFilter = document.getElementById("categoryFilter");
const clientFilter = document.getElementById("clientFilter");
const priceFilter = document.getElementById("priceFilter");
const sortBy = document.getElementById("sortBy");
const resetFilterBtn = document.getElementById("resetFilterBtn");

const productList = document.getElementById("productList");

/* =========================
   DETAIL MODAL ELEMENTS
========================= */
const detailModal = document.getElementById("detailModal");
const modalBackdrop = document.getElementById("modalBackdrop");
const modalClose = document.getElementById("modalClose");
const modalEditBtn = document.getElementById("modalEditBtn");
const modalDeleteBtn = document.getElementById("modalDeleteBtn");

const modalItemNo = document.getElementById("modalItemNo");
const modalItemName = document.getElementById("modalItemName");
const modalCategory = document.getElementById("modalCategory");
const modalClient = document.getElementById("modalClient");
const modalCapital = document.getElementById("modalCapital");
const modalWholesale = document.getElementById("modalWholesale");
const modalDescEn = document.getElementById("modalDescEn");
const modalDescLa = document.getElementById("modalDescLa");
const modalGallery = document.getElementById("modalGallery");
const modalCustomerLogo = document.getElementById("modalCustomerLogo");

/* =========================
   CUSTOMER ELEMENTS
========================= */
const customerForm = document.getElementById("customerForm");
const editingCustomerIdInput = document.getElementById("editingCustomerId");
const customerNameInput = document.getElementById("customerName");
const customerLogoInput = document.getElementById("customerLogo");
const customerLogoPreview = document.getElementById("customerLogoPreview");
const saveCustomerBtn = document.getElementById("saveCustomerBtn");
const cancelCustomerBtn = document.getElementById("cancelCustomerBtn");
const customerListEl = document.getElementById("customerList");

let allProducts = [];
let allCustomers = [];
let editingId = null;
let editingCustomerId = null;
let currentModalItem = null;
let existingPhotos = [];

const currencySymbol = {
  LAK: "₭",
  THB: "฿",
  USD: "$",
  CNY: "¥"
};

/* =========================
   HELPERS
========================= */
function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function uniqueCI(arr) {
  const map = new Map();
  arr.forEach(v => {
    const text = String(v || "").trim();
    if (!text) return;
    const key = text.toLowerCase();
    if (!map.has(key)) map.set(key, text);
  });
  return [...map.values()];
}

function normalizeCurrencyCode(code) {
  return String(code || "").trim().toUpperCase() || "LAK";
}

function getSavedCategories(items) {
  return uniqueCI([...categoryHistory, ...items.map(item => item.category || "")]).sort((a, b) => a.localeCompare(b));
}

function getSavedClients(items) {
  return uniqueCI([...clientHistory, ...items.map(item => item.client || "")]).sort((a, b) => a.localeCompare(b));
}

function stripCurrencyPrefix(value) {
  return String(value ?? "")
    .replace(/^(LAK|THB|USD|CNY)\s*/i, "")
    .replace(/^[₭฿$¥]\s*/i, "")
    .trim();
}

function cleanMoneyInput(value) {
  return stripCurrencyPrefix(value).replace(/,/g, "").trim();
}

function formatNumberWithCommas(value) {
  const cleaned = cleanMoneyInput(value);
  if (!cleaned) return "";
  if (!/^\d*\.?\d*$/.test(cleaned)) return "";
  const parts = cleaned.split(".");
  const intPart = parts[0] || "0";
  const decimalPart = parts[1];
  const formattedInt = Number(intPart).toLocaleString("en-US");
  if (decimalPart !== undefined) return `${formattedInt}.${decimalPart}`;
  return formattedInt;
}

function formatMoney(value) {
  const cleaned = cleanMoneyInput(value);
  const num = Number(cleaned);
  if (Number.isNaN(num)) return "0";
  return num.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function getNumeric(value) {
  const num = Number(cleanMoneyInput(value));
  return Number.isNaN(num) ? 0 : num;
}

function getItemNoNumber(itemNoValue) {
  const match = String(itemNoValue || "").match(/(\d+)/);
  return match ? Number(match[1]) : 0;
}

function getCurrencySymbol(code) {
  return currencySymbol[normalizeCurrencyCode(code)] || "";
}

function buildPriceText(currency, value) {
  const code = normalizeCurrencyCode(currency);
  const formatted = formatMoney(value);
  const symbol = getCurrencySymbol(code);
  if (symbol) return `${symbol}${formatted}`;
  if (code) return `${code} ${formatted}`;
  return formatted;
}

function normalizePhotos(photos) {
  if (!photos) return [];
  if (Array.isArray(photos)) return photos.filter(Boolean);
  if (typeof photos === "string") {
    const trimmed = photos.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.filter(Boolean);
    } catch (_) {}
    return trimmed.split(",").map(v => v.trim()).filter(Boolean);
  }
  return [];
}

function formatDisplay(input, currencySelect) {
  const raw = cleanMoneyInput(input.value);
  if (!raw) { input.value = ""; return; }
  input.value = `${normalizeCurrencyCode(currencySelect.value)} ${formatNumberWithCommas(raw)}`;
}

function refreshPriceWithCurrency(input, currencySelect) {
  formatDisplay(input, currencySelect);
}

function syncEditingIdInput() {
  if (editingIdInput) editingIdInput.value = editingId || "";
}

async function parseJsonSafe(res) {
  const text = await res.text();
  try { return text ? JSON.parse(text) : {}; }
  catch (err) { return { success: false, message: text || "Server returned invalid JSON" }; }
}

/* =========================
   CUSTOMER LOGO LOOKUP
========================= */
function getCustomerLogo(clientName) {
  if (!clientName) return "";
  const key = String(clientName).trim().toLowerCase();
  const found = allCustomers.find(c => String(c.name).trim().toLowerCase() === key);
  return found ? found.logoUrl : "";
}

/* =========================
   CUSTOMER MANAGEMENT
========================= */
function resetCustomerForm() {
  customerForm.reset();
  editingCustomerId = null;
  editingCustomerIdInput.value = "";
  customerLogoPreview.innerHTML = "";
  saveCustomerBtn.textContent = "Add Customer";
}

function renderCustomers() {
  if (!customerListEl) return;
  if (!allCustomers.length) {
    customerListEl.innerHTML = `<div class="customer-empty">No customers yet. Add one above.</div>`;
    return;
  }
  customerListEl.innerHTML = allCustomers.map(c => `
    <div class="customer-item">
      <div class="customer-item-logo">
        ${c.logoUrl
          ? `<img src="${escapeHtml(c.logoUrl)}" alt="${escapeHtml(c.name)}" onerror="this.outerHTML='<div class=\\"cust-no-logo\\">${escapeHtml(c.name.charAt(0))}</div>'">`
          : `<div class="cust-no-logo">${escapeHtml(c.name.charAt(0).toUpperCase())}</div>`
        }
      </div>
      <span class="customer-item-name">${escapeHtml(c.name)}</span>
      <div class="customer-item-actions">
        <button class="btn btn-secondary btn-sm" onclick="startEditCustomer('${escapeHtml(c.id)}')">Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteCustomer('${escapeHtml(c.id)}')">Delete</button>
      </div>
    </div>
  `).join("");
}

async function loadCustomers() {
  try {
    const res = await fetch("/api/customers", { headers: { "Accept": "application/json" } });
    const data = await parseJsonSafe(res);
    allCustomers = Array.isArray(data.customers) ? data.customers : [];
    renderCustomers();
  } catch (err) {
    console.error("Load customers failed:", err);
    allCustomers = [];
    renderCustomers();
  }
}

function startEditCustomer(id) {
  const c = allCustomers.find(x => x.id === id);
  if (!c) return;
  editingCustomerId = c.id;
  editingCustomerIdInput.value = c.id;
  customerNameInput.value = c.name;
  customerLogoPreview.innerHTML = c.logoUrl
    ? `<img src="${escapeHtml(c.logoUrl)}" alt="current logo" class="current-logo-preview" />`
    : "";
  saveCustomerBtn.textContent = "Update Customer";
  customerNameInput.scrollIntoView({ behavior: "smooth", block: "center" });
  customerNameInput.focus();
}

async function deleteCustomer(id) {
  if (!confirm("Delete this customer?")) return;
  try {
    const res = await fetch(`/api/customers/${encodeURIComponent(id)}`, {
      method: "DELETE",
      headers: { "Accept": "application/json" }
    });
    const data = await parseJsonSafe(res);
    if (!res.ok || !data.success) { alert(data.message || "Delete failed"); return; }
    await loadCustomers();
    applyFiltersAndSort();
  } catch (err) {
    alert(err.message || "Delete failed");
  }
}

customerForm.addEventListener("submit", async e => {
  e.preventDefault();
  const isEditing = !!editingCustomerId;
  try {
    saveCustomerBtn.disabled = true;
    saveCustomerBtn.textContent = isEditing ? "Updating..." : "Adding...";

    const formData = new FormData();
    formData.append("name", customerNameInput.value.trim());
    if (customerLogoInput.files[0]) formData.append("logo", customerLogoInput.files[0]);

    const url = isEditing ? `/api/customers/${editingCustomerId}` : "/api/customers";
    const method = isEditing ? "PUT" : "POST";

    const res = await fetch(url, { method, body: formData, headers: { "Accept": "application/json" } });
    const data = await parseJsonSafe(res);
    if (!res.ok || !data.success) throw new Error(data.message || "Save failed");

    resetCustomerForm();
    await loadCustomers();
    applyFiltersAndSort();
  } catch (err) {
    alert(err.message || "Something went wrong");
  } finally {
    saveCustomerBtn.disabled = false;
    saveCustomerBtn.textContent = editingCustomerId ? "Update Customer" : "Add Customer";
  }
});

cancelCustomerBtn.addEventListener("click", resetCustomerForm);

customerLogoInput.addEventListener("change", () => {
  const file = customerLogoInput.files[0];
  if (!file) { customerLogoPreview.innerHTML = ""; return; }
  const reader = new FileReader();
  reader.onload = e => {
    customerLogoPreview.innerHTML = `<img src="${e.target.result}" alt="preview" class="current-logo-preview" />`;
  };
  reader.readAsDataURL(file);
});

/* =========================
   DELETE PRODUCT
========================= */
async function deleteProduct(id) {
  if (!id) return;
  const ok = confirm("Are you sure you want to delete this product?");
  if (!ok) return;
  try {
    const res = await fetch(`/api/items/${encodeURIComponent(id)}`, {
      method: "DELETE",
      headers: { "Accept": "application/json" }
    });
    const text = await res.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; }
    catch (e) { data = { success: false, message: text || "Server returned invalid response" }; }
    if (!res.ok || !data.success) { alert(data.message || `Delete failed (${res.status})`); return; }
    alert("Product deleted successfully");
    if (currentModalItem && String(currentModalItem.id) === String(id)) closeDetail();
    await loadProducts();
  } catch (err) {
    console.error("Delete failed:", err);
    alert(err.message || "Delete failed");
  }
}

/* =========================
   CUSTOM SUGGESTIONS
========================= */
function hideSuggestions(box) {
  if (!box) return;
  box.innerHTML = "";
  box.classList.remove("show");
  box.style.display = "none";
}

function showSuggestions(input, box, list) {
  if (!box) return;
  const value = String(input.value || "").trim().toLowerCase();
  const filtered = uniqueCI(list).filter(v => String(v).toLowerCase().includes(value));
  box.innerHTML = "";
  if (!filtered.length) { hideSuggestions(box); return; }
  filtered.forEach(text => {
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.textContent = text;
    div.addEventListener("mousedown", e => {
      e.preventDefault();
      input.value = text;
      hideSuggestions(box);
      input.focus();
    });
    box.appendChild(div);
  });
  box.classList.add("show");
  box.style.display = "block";
}

function getCategorySuggestionList() {
  const typed = category.value.trim();
  return uniqueCI([typed, ...categoryHistory, ...getSavedCategories(allProducts)]);
}

function getClientSuggestionList() {
  const typed = client.value.trim();
  return uniqueCI([typed, ...clientHistory, ...getSavedClients(allProducts)]);
}

function wireAutocomplete(input, box, getList, saveHistory) {
  input.addEventListener("focus", () => { showSuggestions(input, box, getList()); });
  input.addEventListener("input", () => { showSuggestions(input, box, getList()); });
  input.addEventListener("blur", () => {
    const typed = input.value.trim();
    if (typed) saveHistory(typed);
    setTimeout(() => { hideSuggestions(box); }, 150);
  });
}

/* =========================
   FORM
========================= */
function resetForm() {
  productForm.reset();
  itemNo.value = "AUTO";
  photoPreview.innerHTML = "";
  capitalCurrency.value = "LAK";
  wholesaleCurrency.value = "LAK";
  capitalPrice.value = "";
  wholesalePrice.value = "";
  editingId = null;
  existingPhotos = [];
  currentModalItem = null;
  syncEditingIdInput();
  saveBtn.textContent = "Save Product";
  saveBtn.disabled = false;
  const titleEl2 = document.getElementById("productFormTitle");
  if (titleEl2) titleEl2.textContent = "Add Product";
  hideSuggestions(categorySuggestions);
  hideSuggestions(clientSuggestions);
}

function renderPhotoPreview(files) {
  photoPreview.innerHTML = "";
  const selected = Array.from(files || []).slice(0, 5);
  selected.forEach(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const div = document.createElement("div");
      div.className = "preview-item";
      div.innerHTML = `<img src="${e.target.result}" alt="${escapeHtml(file.name)}" />`;
      photoPreview.appendChild(div);
    };
    reader.readAsDataURL(file);
  });
}

function renderExistingPhotoPreview(photoUrls) {
  photoPreview.innerHTML = "";
  photoUrls.forEach((url, index) => {
    const div = document.createElement("div");
    div.className = "preview-item existing";
    div.innerHTML = `
      <img src="${escapeHtml(url)}" alt="photo ${index + 1}" />
      <button type="button" class="remove-existing-photo" data-index="${index}">×</button>
    `;
    photoPreview.appendChild(div);
  });
  photoPreview.querySelectorAll(".remove-existing-photo").forEach(btn => {
    btn.addEventListener("click", e => {
      const index = Number(e.currentTarget.dataset.index);
      existingPhotos.splice(index, 1);
      renderExistingPhotoPreview(existingPhotos);
    });
  });
}

/* =========================
   FILTER OPTIONS
========================= */
function populateFilterOptions(items) {
  const categories = getSavedCategories(items);
  const clients = getSavedClients(items);
  const currentCategory = categoryFilter.value;
  const currentClient = clientFilter.value;

  categoryFilter.innerHTML =
    `<option value="">All Categories</option>` +
    categories.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");

  clientFilter.innerHTML =
    `<option value="">All Clients</option>` +
    clients.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");

  if (categories.includes(currentCategory)) categoryFilter.value = currentCategory;
  if (clients.includes(currentClient)) clientFilter.value = currentClient;
}

/* =========================
   FILTER + SORT
========================= */
function filterProducts(items) {
  const keyword = searchInput.value.trim().toLowerCase();
  const selectedCategory = categoryFilter.value.trim().toLowerCase();
  const selectedClient = clientFilter.value.trim().toLowerCase();
  const selectedPrice = priceFilter.value;

  return items.filter(item => {
    const itemWholesale = getNumeric(item.wholesalePrice);
    const haystack = [item.itemNo, item.itemName, item.category, item.client, item.descEn, item.descLa, item.capitalCurrency, item.wholesaleCurrency].join(" ").toLowerCase();
    const matchKeyword = !keyword || haystack.includes(keyword);
    const matchCategory = !selectedCategory || String(item.category || "").toLowerCase() === selectedCategory;
    const matchClient = !selectedClient || String(item.client || "").toLowerCase() === selectedClient;
    let matchPrice = true;
    if (selectedPrice === "0-100") matchPrice = itemWholesale >= 0 && itemWholesale <= 100;
    if (selectedPrice === "101-500") matchPrice = itemWholesale >= 101 && itemWholesale <= 500;
    if (selectedPrice === "501-1000") matchPrice = itemWholesale >= 501 && itemWholesale <= 1000;
    if (selectedPrice === "1001-5000") matchPrice = itemWholesale >= 1001 && itemWholesale <= 5000;
    if (selectedPrice === "5001+") matchPrice = itemWholesale >= 5001;
    return matchKeyword && matchCategory && matchClient && matchPrice;
  });
}

function sortProducts(items) {
  const list = [...items];
  const mode = sortBy.value;
  list.sort((a, b) => {
    if (mode === "newest") return getItemNoNumber(b.itemNo) - getItemNoNumber(a.itemNo);
    if (mode === "oldest") return getItemNoNumber(a.itemNo) - getItemNoNumber(b.itemNo);
    if (mode === "name-asc") return String(a.itemName || "").localeCompare(String(b.itemName || ""));
    if (mode === "name-desc") return String(b.itemName || "").localeCompare(String(a.itemName || ""));
    if (mode === "price-asc") return getNumeric(a.wholesalePrice) - getNumeric(b.wholesalePrice);
    if (mode === "price-desc") return getNumeric(b.wholesalePrice) - getNumeric(a.wholesalePrice);
    if (mode === "itemno-asc") return getItemNoNumber(a.itemNo) - getItemNoNumber(b.itemNo);
    if (mode === "itemno-desc") return getItemNoNumber(b.itemNo) - getItemNoNumber(a.itemNo);
    return 0;
  });
  return list;
}

/* =========================
   DETAIL MODAL
========================= */
function openDetail(id) {
  const item = allProducts.find(x => String(x.id) === String(id));
  if (!item || !detailModal) return;
  currentModalItem = item;
  const photos = normalizePhotos(item.photos);

  modalItemNo.textContent = item.itemNo || "-";
  modalItemName.textContent = item.itemName || "-";
  modalCategory.textContent = item.category || "-";
  modalClient.textContent = item.client || "-";
  modalCapital.textContent = buildPriceText(item.capitalCurrency, item.capitalPrice);
  modalWholesale.textContent = buildPriceText(item.wholesaleCurrency, item.wholesalePrice);
  modalDescEn.textContent = item.descEn || "-";
  modalDescLa.textContent = item.descLa || "-";

  // Customer logo in modal
  const logoUrl = getCustomerLogo(item.client);
  if (logoUrl && modalCustomerLogo) {
    modalCustomerLogo.innerHTML = `<img src="${escapeHtml(logoUrl)}" alt="${escapeHtml(item.client || "")}" onerror="this.parentElement.classList.add('hidden')" />`;
    modalCustomerLogo.classList.remove("hidden");
  } else if (modalCustomerLogo) {
    modalCustomerLogo.innerHTML = "";
    modalCustomerLogo.classList.add("hidden");
  }

  modalGallery.innerHTML = "";
  if (photos.length) {
    modalGallery.innerHTML = photos.map(src => `<img src="${escapeHtml(src)}" alt="photo" onerror="this.style.display='none'">`).join("");
  } else {
    modalGallery.innerHTML = `<div class="no-photo large">No Photo</div>`;
  }

  detailModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
}

function closeDetail() {
  if (!detailModal) return;
  detailModal.classList.add("hidden");
  document.body.style.overflow = "";
  currentModalItem = null;
}

/* =========================
   EDIT
========================= */
function startEdit(id) {
  const item = allProducts.find(x => String(x.id) === String(id));
  if (!item) return;
  editingId = item.id;
  syncEditingIdInput();
  existingPhotos = normalizePhotos(item.photos);
  itemNo.value = item.itemNo || "";
  itemName.value = item.itemName || "";
  category.value = item.category || "";
  client.value = item.client || "";
  capitalCurrency.value = normalizeCurrencyCode(item.capitalCurrency || "LAK");
  capitalPrice.value = item.capitalPrice || "";
  refreshPriceWithCurrency(capitalPrice, capitalCurrency);
  wholesaleCurrency.value = normalizeCurrencyCode(item.wholesaleCurrency || "LAK");
  wholesalePrice.value = item.wholesalePrice || "";
  refreshPriceWithCurrency(wholesalePrice, wholesaleCurrency);
  descEn.value = item.descEn || "";
  descLa.value = item.descLa || "";
  photosInput.value = "";
  renderExistingPhotoPreview(existingPhotos);
  saveBtn.textContent = "Update Product";
  saveBtn.disabled = false;
  const titleEl = document.getElementById("productFormTitle");
  if (titleEl) titleEl.textContent = "Edit Product";
  closeDetail();
  openProductForm();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* =========================
   RENDER PRODUCTS
========================= */
function renderProducts(items) {
  if (!productList) return;
  if (!items || !items.length) {
    productList.innerHTML = `<div class="empty-state">No products found.</div>`;
    return;
  }

  productList.innerHTML = items.map(item => {
    const photos = normalizePhotos(item.photos);
    const photo = photos.length ? photos[0] : "";
    const wholesale = buildPriceText(item.wholesaleCurrency, item.wholesalePrice || 0);
    const safeId = escapeHtml(item.id || "");
    const logoUrl = getCustomerLogo(item.client);

    return `
      <div class="product-card clean-card" onclick="openDetail('${safeId}')">
        <div class="card-inner">
          <div class="card-photo-wrap">
            ${photo
              ? `<img class="card-photo" src="${photo}" alt="${escapeHtml(item.itemName || "")}" onerror="this.outerHTML='<div class=&quot;card-no-photo&quot;>No Photo</div>'">`
              : `<div class="card-no-photo">No Photo</div>`
            }
            ${logoUrl
              ? `<div class="card-customer-badge"><img src="${escapeHtml(logoUrl)}" alt="${escapeHtml(item.client || "")}" onerror="this.parentElement.style.display='none'" /></div>`
              : item.client
                ? `<div class="card-customer-badge card-customer-initial">${escapeHtml(item.client.charAt(0).toUpperCase())}</div>`
                : ""
            }
          </div>

          <div class="card-bottom-row">
            <h3 class="card-title">${escapeHtml(item.itemName || "-")}</h3>
            <div class="card-price-pill">${wholesale}</div>
          </div>

          <div class="card-itemno">${escapeHtml(item.itemNo || "-")}</div>
        </div>
      </div>
    `;
  }).join("");
}

function applyFiltersAndSort() {
  const filtered = filterProducts(allProducts);
  const sorted = sortProducts(filtered);
  renderProducts(sorted);
}

/* =========================
   LOAD PRODUCTS
========================= */
async function loadProducts() {
  try {
    const res = await fetch("/api/items", { headers: { "Accept": "application/json" } });
    const data = await parseJsonSafe(res);
    if (!res.ok || !data.success) {
      allProducts = [];
      populateFilterOptions([]);
      renderProducts([]);
      return;
    }
    allProducts = Array.isArray(data.items) ? data.items : [];
    populateFilterOptions(allProducts);
    applyFiltersAndSort();
  } catch (err) {
    console.error("Load products failed:", err);
    allProducts = [];
    populateFilterOptions([]);
    renderProducts([]);
  }
}

/* =========================
   EVENTS
========================= */
photosInput.addEventListener("change", () => {
  const files = Array.from(photosInput.files || []);
  if (files.length > 5) {
    alert("You can upload up to 5 images only.");
    photosInput.value = "";
    photoPreview.innerHTML = "";
    return;
  }
  if (files.length) renderPhotoPreview(files);
  else if (existingPhotos.length) renderExistingPhotoPreview(existingPhotos);
  else photoPreview.innerHTML = "";
});

capitalPrice.addEventListener("input", function () { formatDisplay(this, capitalCurrency); });
wholesalePrice.addEventListener("input", function () { formatDisplay(this, wholesaleCurrency); });
capitalCurrency.addEventListener("change", () => { refreshPriceWithCurrency(capitalPrice, capitalCurrency); });
wholesaleCurrency.addEventListener("change", () => { refreshPriceWithCurrency(wholesalePrice, wholesaleCurrency); });

productForm.addEventListener("submit", async e => {
  e.preventDefault();
  const isEditing = !!editingId;
  try {
    saveBtn.disabled = true;
    saveBtn.textContent = isEditing ? "Updating..." : "Saving...";
    const files = Array.from(photosInput.files || []);
    const totalPhotos = existingPhotos.length + files.length;
    if (totalPhotos > 5) {
      alert("You can keep/upload up to 5 images only.");
      saveBtn.disabled = false;
      saveBtn.textContent = isEditing ? "Update Product" : "Save Product";
      return;
    }
    const formData = new FormData();
    formData.append("itemName", itemName.value.trim());
    formData.append("category", category.value.trim());
    formData.append("client", client.value.trim());
    formData.append("capitalCurrency", normalizeCurrencyCode(capitalCurrency.value));
    formData.append("capitalPrice", cleanMoneyInput(capitalPrice.value));
    formData.append("wholesaleCurrency", normalizeCurrencyCode(wholesaleCurrency.value));
    formData.append("wholesalePrice", cleanMoneyInput(wholesalePrice.value));
    formData.append("descEn", descEn.value.trim());
    formData.append("descLa", descLa.value.trim());
    formData.append("existingPhotos", JSON.stringify(existingPhotos));
    files.forEach(file => { formData.append("photos", file); });

    const typedCategory = category.value.trim();
    const typedClient = client.value.trim();
    if (typedCategory) categoryHistory = uniqueCI([typedCategory, ...categoryHistory]);
    if (typedClient) clientHistory = uniqueCI([typedClient, ...clientHistory]);

    const url = isEditing ? `/api/items/${editingId}` : "/api/items";
    const method = isEditing ? "PUT" : "POST";
    const res = await fetch(url, { method, body: formData, headers: { "Accept": "application/json" } });
    const data = await parseJsonSafe(res);
    if (!res.ok || !data.success) throw new Error(data.message || "Save failed");

    alert(isEditing
      ? `Updated successfully. Item No: ${data.item?.itemNo || itemNo.value || ""}`
      : `Saved successfully. Item No: ${data.item?.itemNo || ""}`
    );
    resetForm();
    closeProductForm();
    await loadProducts();
  } catch (err) {
    console.error(err);
    alert(err.message || "Something went wrong");
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = editingId ? "Update Product" : "Save Product";
  }
});

clearBtn.addEventListener("click", () => { resetForm(); closeProductForm(); });
downloadExcelBtn.addEventListener("click", () => { window.location.href = "/api/download-excel"; });

let searchTimer;
searchInput.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { applyFiltersAndSort(); }, 250);
});

categoryFilter.addEventListener("change", applyFiltersAndSort);
clientFilter.addEventListener("change", applyFiltersAndSort);
priceFilter.addEventListener("change", applyFiltersAndSort);
sortBy.addEventListener("change", applyFiltersAndSort);

resetFilterBtn.addEventListener("click", () => {
  searchInput.value = "";
  categoryFilter.value = "";
  clientFilter.value = "";
  priceFilter.value = "";
  sortBy.value = "newest";
  applyFiltersAndSort();
});

document.addEventListener("click", e => {
  if (!e.target.closest(".autocomplete")) {
    hideSuggestions(categorySuggestions);
    hideSuggestions(clientSuggestions);
  }
});

if (modalClose) modalClose.addEventListener("click", closeDetail);
if (modalBackdrop) modalBackdrop.addEventListener("click", closeDetail);
if (modalEditBtn) modalEditBtn.addEventListener("click", () => { if (currentModalItem) startEdit(currentModalItem.id); });
if (modalDeleteBtn) modalDeleteBtn.addEventListener("click", async () => { if (currentModalItem) await deleteProduct(currentModalItem.id); });

window.addEventListener("keydown", e => {
  if (e.key === "Escape" && detailModal && !detailModal.classList.contains("hidden")) closeDetail();
});

window.openDetail = openDetail;
window.startEdit = startEdit;
window.closeDetail = closeDetail;
window.deleteProduct = deleteProduct;
window.startEditCustomer = startEditCustomer;
window.deleteCustomer = deleteCustomer;

/* =========================
   COLLAPSIBLE PRODUCT FORM
========================= */
function openProductForm() {
  const body = document.getElementById("productFormBody");
  const btn = document.getElementById("productFormCollapseBtn");
  if (!body) return;
  body.classList.remove("collapsed");
  if (btn) btn.classList.add("open");
}

function closeProductForm() {
  const body = document.getElementById("productFormBody");
  const btn = document.getElementById("productFormCollapseBtn");
  if (!body) return;
  body.classList.add("collapsed");
  if (btn) btn.classList.remove("open");
}

window.addEventListener("DOMContentLoaded", async () => {
  // Toggle button wiring
  const toggleBtn = document.getElementById("productFormCollapseBtn");
  const toggleHead = document.getElementById("productFormToggle");
  if (toggleBtn) toggleBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const body = document.getElementById("productFormBody");
    if (body && body.classList.contains("collapsed")) openProductForm();
    else closeProductForm();
  });
  if (toggleHead) toggleHead.addEventListener("click", (e) => {
    if (e.target.closest("button")) return;
    const body = document.getElementById("productFormBody");
    if (body && body.classList.contains("collapsed")) openProductForm();
    else closeProductForm();
  });
  resetForm();
  await loadCustomers();
  await loadProducts();

  wireAutocomplete(category, categorySuggestions, getCategorySuggestionList, typed => {
    categoryHistory = uniqueCI([typed, ...categoryHistory]);
  });
  wireAutocomplete(client, clientSuggestions, getClientSuggestionList, typed => {
    clientHistory = uniqueCI([typed, ...clientHistory]);
  });
});
