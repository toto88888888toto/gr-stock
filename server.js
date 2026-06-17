const express = require("express");
const cors = require("cors");
const multer = require("multer");
const { google } = require("googleapis");
const { Readable } = require("stream");
const ExcelJS = require("exceljs");
const { v4: uuidv4 } = require("uuid");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

// ── Google Sheets config (service account) ────────────────────
const CREDENTIALS = JSON.parse(process.env.GOOGLE_CREDENTIALS || process.env.GOOGLE_SERVICE_ACCOUNT);
const SHEET_ID = process.env.SHEET_ID || "1e7jXUY4kC0ecGldIEBSsewnhPJSvtiwvmTIm7K8uqoA";
const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID || "";
const SHEET_NAME = "Items";
const CUSTOMER_SHEET_NAME = "Clients";

const sheetsAuth = new google.auth.GoogleAuth({
  credentials: CREDENTIALS,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"]
});
const sheets = google.sheets({ version: "v4", auth: sheetsAuth });

// ── Google Drive config (OAuth2 - personal account) ───────────
let drive = null;
function initDrive() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken || !DRIVE_FOLDER_ID) {
    console.log("[Drive] OAuth credentials not set — Drive uploads disabled");
    console.log("[Drive] clientId:", !!clientId, "secret:", !!clientSecret, "token:", !!refreshToken, "folder:", !!DRIVE_FOLDER_ID);
    return;
  }
  const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, "http://localhost:3333/callback");
  oauth2Client.setCredentials({ refresh_token: refreshToken });
  drive = google.drive({ version: "v3", auth: oauth2Client });
  console.log("[Drive] Google Drive initialized (OAuth2) — folder:", DRIVE_FOLDER_ID);
}
initDrive();

// ── Get or create a named subfolder inside DRIVE_FOLDER_ID ─────
async function getOrCreateSubfolder(name) {
  if (!drive) throw new Error("Drive not initialized");
  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${DRIVE_FOLDER_ID}' in parents and trashed=false`,
    fields: "files(id, name)",
  });
  if (res.data.files.length > 0) return res.data.files[0].id;
  const folder = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [DRIVE_FOLDER_ID] },
    fields: "id",
  });
  console.log("[Drive] Created subfolder:", name, folder.data.id);
  return folder.data.id;
}

// ── Multer (memory storage — no local files) ───────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { files: 20, fileSize: 10 * 1024 * 1024 }
});

app.use(cors());
app.use((req, res, next) => { console.log(`[REQ] ${req.method} ${req.path}`); next(); });
app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

// ── Constants ──────────────────────────────────────────────────
const HEADERS = [
  "id", "itemNo", "itemName", "category", "client",
  "capitalCurrency", "capitalPrice", "wholesaleCurrency", "wholesalePrice",
  "descEn", "descLa", "photos", "createdAt"
];

const CUSTOMER_HEADERS = ["id", "name", "logoUrl"];

// ── Helpers ────────────────────────────────────────────────────
function normalizeMoney(value) {
  const cleaned = String(value || "").replace(/[^\d.-]/g, "").trim();
  if (!cleaned) return "";
  const num = Number(cleaned);
  return Number.isNaN(num) ? "" : String(num);
}

function normalizeCurrency(value) {
  return String(value || "").trim().toUpperCase();
}

function parsePhotos(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.filter(Boolean).map(v => String(v).trim()).filter(Boolean);
  const str = String(value).trim();
  if (!str) return [];
  try {
    const parsed = JSON.parse(str);
    if (Array.isArray(parsed)) return parsed.filter(Boolean).map(v => String(v).trim()).filter(Boolean);
  } catch (e) {}
  return str.split("|").map(v => String(v).trim()).filter(Boolean);
}

function photosToCellValue(photos) {
  return JSON.stringify((Array.isArray(photos) ? photos : []).filter(Boolean));
}

function uniqueArray(arr) {
  return [...new Set((arr || []).filter(Boolean).map(v => String(v).trim()).filter(Boolean))];
}

function generateNextItemNo(items) {
  let max = 0;
  for (const item of items) {
    const match = String(item.itemNo || "").match(/GR(\d+)/i);
    if (match) {
      const num = Number(match[1]);
      if (!Number.isNaN(num) && num > max) max = num;
    }
  }
  return `GR${String(max + 1).padStart(6, "0")}`;
}

// ── Google Sheets helpers ──────────────────────────────────────
async function ensureHeaders() {
  // Auto-create "Items" sheet if it doesn't exist
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = spreadsheet.data.sheets.some(s => s.properties.title === SHEET_NAME);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_NAME } } }] }
    });
  }
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A1:M1`
  });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [HEADERS] }
    });
  }
}

async function getAllItems() {
  await ensureHeaders();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:M`
  });
  const rows = res.data.values || [];
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .map(row => ({
      id: row[0] || "",
      itemNo: row[1] || "",
      itemName: row[2] || "",
      category: row[3] || "",
      client: row[4] || "",
      capitalCurrency: row[5] || "",
      capitalPrice: row[6] || "",
      wholesaleCurrency: row[7] || "",
      wholesalePrice: row[8] || "",
      descEn: row[9] || "",
      descLa: row[10] || "",
      photos: parsePhotos(row[11] || ""),
      createdAt: row[12] || ""
    }))
    .filter(item => item.id || item.itemName);
}

// ── Customer sheet helpers ─────────────────────────────────────
async function ensureCustomerSheet() {
  // Ensure the Customers sheet exists
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = spreadsheet.data.sheets.some(s => s.properties.title === CUSTOMER_SHEET_NAME);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: CUSTOMER_SHEET_NAME } } }] }
    });
  }
  // Ensure headers
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${CUSTOMER_SHEET_NAME}!A1:C1`
  });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${CUSTOMER_SHEET_NAME}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [CUSTOMER_HEADERS] }
    });
  }
}

async function getAllCustomers() {
  await ensureCustomerSheet();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${CUSTOMER_SHEET_NAME}!A:C`
  });
  const rows = res.data.values || [];
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .map(row => ({
      id: row[0] || "",
      name: row[1] || "",
      logoUrl: row[2] || ""
    }))
    .filter(c => c.id || c.name);
}

// ── Google Drive helpers (OAuth2) ─────────────────────────────
async function uploadToCloudinary(buffer, filename, folderId = DRIVE_FOLDER_ID) {
  if (!drive) throw new Error("Google Drive not initialized");
  const mimeType = filename.match(/\.png$/i) ? "image/png" : "image/jpeg";
  const file = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: "id"
  });
  await drive.permissions.create({
    fileId: file.data.id,
    requestBody: { role: "reader", type: "anyone" }
  });
  return `https://drive.google.com/uc?id=${file.data.id}`;
}

async function deleteFromCloudinary(url) {
  try {
    if (!drive) return;
    const match = String(url || "").match(/id=([^&]+)/);
    if (match) await drive.files.delete({ fileId: match[1] });
  } catch (e) {
    console.warn("Failed to delete Drive file:", e.message);
  }
}

// ── GET ITEMS ──────────────────────────────────────────────────
app.get("/api/items", async (req, res) => {
  try {
    const items = await getAllItems();
    items.sort((a, b) => {
      const aNum = Number(String(a.itemNo || "").replace(/\D/g, "")) || 0;
      const bNum = Number(String(b.itemNo || "").replace(/\D/g, "")) || 0;
      return bNum - aNum;
    });
    return res.json({ success: true, items });
  } catch (error) {
    console.error("GET /api/items error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── SAVE ITEM ──────────────────────────────────────────────────
app.post("/api/items", upload.array("photos", 20), async (req, res) => {
  try {
    const existingItems = await getAllItems();

    const photoPaths = [];
    for (const file of (req.files || [])) {
      const url = await uploadToCloudinary(file.buffer, file.originalname);
      photoPaths.push(url);
    }

    const item = {
      id: uuidv4(),
      itemNo: generateNextItemNo(existingItems),
      itemName: String(req.body.itemName || "").trim(),
      category: String(req.body.category || "").trim(),
      client: String(req.body.client || "").trim(),
      capitalCurrency: normalizeCurrency(req.body.capitalCurrency || "LAK"),
      capitalPrice: normalizeMoney(req.body.capitalPrice),
      wholesaleCurrency: normalizeCurrency(req.body.wholesaleCurrency || "LAK"),
      wholesalePrice: normalizeMoney(req.body.wholesalePrice),
      descEn: String(req.body.descEn || "").trim(),
      descLa: String(req.body.descLa || "").trim(),
      photos: uniqueArray(photoPaths),
      createdAt: new Date().toISOString()
    };

    if (!item.itemName) {
      return res.status(400).json({ success: false, message: "Item Name is required" });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:M`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          item.id, item.itemNo, item.itemName, item.category, item.client,
          item.capitalCurrency, item.capitalPrice, item.wholesaleCurrency, item.wholesalePrice,
          item.descEn, item.descLa, photosToCellValue(item.photos), item.createdAt
        ]]
      }
    });

    return res.json({ success: true, item });
  } catch (error) {
    console.error("POST /api/items error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── UPDATE ITEM ────────────────────────────────────────────────
app.put("/api/items/:id", upload.array("photos", 20), async (req, res) => {
  try {
    const targetId = String(req.params.id || "").trim();

    const sheetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:M`
    });

    const rows = sheetRes.data.values || [];
    let targetRowIndex = -1;
    let targetItem = null;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim() === targetId) {
        targetRowIndex = i + 1;
        targetItem = {
          id: rows[i][0] || "",
          itemNo: rows[i][1] || "",
          capitalCurrency: rows[i][5] || "LAK",
          wholesaleCurrency: rows[i][7] || "LAK",
          photos: parsePhotos(rows[i][11] || ""),
          createdAt: rows[i][12] || new Date().toISOString()
        };
        break;
      }
    }

    if (!targetItem) {
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    const keptPhotos = parsePhotos(req.body.existingPhotos);
    const newPhotoPaths = [];
    for (const file of (req.files || [])) {
      const url = await uploadToCloudinary(file.buffer, file.originalname);
      newPhotoPaths.push(url);
    }
    const finalPhotos = uniqueArray([...keptPhotos, ...newPhotoPaths]);
    const removedPhotos = targetItem.photos.filter(p => !finalPhotos.includes(p));

    const updatedItem = {
      id: targetItem.id,
      itemNo: targetItem.itemNo,
      itemName: String(req.body.itemName || "").trim(),
      category: String(req.body.category || "").trim(),
      client: String(req.body.client || "").trim(),
      capitalCurrency: normalizeCurrency(req.body.capitalCurrency || targetItem.capitalCurrency),
      capitalPrice: normalizeMoney(req.body.capitalPrice),
      wholesaleCurrency: normalizeCurrency(req.body.wholesaleCurrency || targetItem.wholesaleCurrency),
      wholesalePrice: normalizeMoney(req.body.wholesalePrice),
      descEn: String(req.body.descEn || "").trim(),
      descLa: String(req.body.descLa || "").trim(),
      photos: finalPhotos,
      createdAt: targetItem.createdAt
    };

    if (!updatedItem.itemName) {
      return res.status(400).json({ success: false, message: "Item Name is required" });
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A${targetRowIndex}:M${targetRowIndex}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          updatedItem.id, updatedItem.itemNo, updatedItem.itemName, updatedItem.category, updatedItem.client,
          updatedItem.capitalCurrency, updatedItem.capitalPrice, updatedItem.wholesaleCurrency, updatedItem.wholesalePrice,
          updatedItem.descEn, updatedItem.descLa, photosToCellValue(updatedItem.photos), updatedItem.createdAt
        ]]
      }
    });

    for (const photo of removedPhotos) {
      await deleteFromCloudinary(photo);
    }

    return res.json({ success: true, item: updatedItem });
  } catch (error) {
    console.error("PUT /api/items/:id error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── DELETE ITEM ────────────────────────────────────────────────
app.delete("/api/items/:id", async (req, res) => {
  try {
    const targetId = String(req.params.id || "").trim();

    const sheetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:M`
    });

    const rows = sheetRes.data.values || [];
    let targetRowIndex = -1;
    let targetPhotos = [];

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim() === targetId) {
        targetRowIndex = i + 1;
        targetPhotos = parsePhotos(rows[i][11] || "");
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheet = spreadsheet.data.sheets.find(s => s.properties.title === SHEET_NAME);
    const sheetId = sheet.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId,
              dimension: "ROWS",
              startIndex: targetRowIndex - 1,
              endIndex: targetRowIndex
            }
          }
        }]
      }
    });

    for (const photo of targetPhotos) {
      await deleteFromCloudinary(photo);
    }

    return res.json({ success: true, message: "Item deleted successfully" });
  } catch (error) {
    console.error("DELETE /api/items/:id error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── DOWNLOAD EXCEL ─────────────────────────────────────────────
app.get("/api/download-excel", async (req, res) => {
  try {
    const items = await getAllItems();
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Items");
    sheet.addRow(HEADERS);
    for (const item of items) {
      sheet.addRow([
        item.id, item.itemNo, item.itemName, item.category, item.client,
        item.capitalCurrency, item.capitalPrice, item.wholesaleCurrency, item.wholesalePrice,
        item.descEn, item.descLa, photosToCellValue(item.photos), item.createdAt
      ]);
    }
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=stock.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("GET /api/download-excel error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── GET CUSTOMERS ──────────────────────────────────────────────
app.get("/api/customers", async (req, res) => {
  try {
    const customers = await getAllCustomers();
    return res.json({ success: true, customers });
  } catch (error) {
    console.error("GET /api/customers error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── CREATE CUSTOMER ────────────────────────────────────────────
app.post("/api/customers", upload.single("logo"), async (req, res) => {
  try {
    const name = String(req.body.name || "").trim();
    if (!name) return res.status(400).json({ success: false, message: "Customer name is required" });

    let logoUrl = "";
    console.log("[Customer POST] name:", name, "hasFile:", !!req.file);
    if (req.file) {
      console.log("[Customer POST] uploading file:", req.file.originalname, req.file.size);
      const logoFolderId = await getOrCreateSubfolder("Customer logo");
      logoUrl = await uploadToCloudinary(req.file.buffer, req.file.originalname, logoFolderId);
      console.log("[Customer POST] uploaded URL:", logoUrl);
    }

    const customer = { id: uuidv4(), name, logoUrl };

    await ensureCustomerSheet();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${CUSTOMER_SHEET_NAME}!A:C`,
      valueInputOption: "RAW",
      requestBody: { values: [[customer.id, customer.name, customer.logoUrl]] }
    });

    return res.json({ success: true, customer });
  } catch (error) {
    console.error("POST /api/customers error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── UPDATE CUSTOMER ────────────────────────────────────────────
app.put("/api/customers/:id", upload.single("logo"), async (req, res) => {
  try {
    const targetId = String(req.params.id || "").trim();
    await ensureCustomerSheet();

    const sheetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${CUSTOMER_SHEET_NAME}!A:C`
    });

    const rows = sheetRes.data.values || [];
    let targetRowIndex = -1;
    let oldLogoUrl = "";

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim() === targetId) {
        targetRowIndex = i + 1;
        oldLogoUrl = rows[i][2] || "";
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ success: false, message: "Customer not found" });
    }

    const name = String(req.body.name || "").trim();
    if (!name) return res.status(400).json({ success: false, message: "Customer name is required" });

    let logoUrl = oldLogoUrl;
    if (req.file) {
      logoUrl = await uploadToCloudinary(req.file.buffer, req.file.originalname);
      if (oldLogoUrl) await deleteFromCloudinary(oldLogoUrl);
    } else if (req.body.removeLogo === "true") {
      if (oldLogoUrl) await deleteFromCloudinary(oldLogoUrl);
      logoUrl = "";
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${CUSTOMER_SHEET_NAME}!A${targetRowIndex}:C${targetRowIndex}`,
      valueInputOption: "RAW",
      requestBody: { values: [[targetId, name, logoUrl]] }
    });

    return res.json({ success: true, customer: { id: targetId, name, logoUrl } });
  } catch (error) {
    console.error("PUT /api/customers/:id error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ── DELETE CUSTOMER ────────────────────────────────────────────
app.delete("/api/customers/:id", async (req, res) => {
  try {
    const targetId = String(req.params.id || "").trim();
    await ensureCustomerSheet();

    const sheetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${CUSTOMER_SHEET_NAME}!A:C`
    });

    const rows = sheetRes.data.values || [];
    let targetRowIndex = -1;
    let logoUrl = "";

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim() === targetId) {
        targetRowIndex = i + 1;
        logoUrl = rows[i][2] || "";
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ success: false, message: "Customer not found" });
    }

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheet = spreadsheet.data.sheets.find(s => s.properties.title === CUSTOMER_SHEET_NAME);

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheet.properties.sheetId,
              dimension: "ROWS",
              startIndex: targetRowIndex - 1,
              endIndex: targetRowIndex
            }
          }
        }]
      }
    });

    if (logoUrl) await deleteFromCloudinary(logoUrl);

    return res.json({ success: true, message: "Customer deleted" });
  } catch (error) {
    console.error("DELETE /api/customers/:id error:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
});


// ── OAuth flow to generate refresh token ───────────────────────
const RAILWAY_REDIRECT = "https://gr-stock-production-83a5.up.railway.app/oauth-callback";

app.get("/oauth-start", (req, res) => {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  if (!clientId || !clientSecret) return res.send("Missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET");
  const oauth2 = new google.auth.OAuth2(clientId, clientSecret, RAILWAY_REDIRECT);
  const url = oauth2.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    login_hint: "outamavongsa.toto@gmail.com",
    scope: ["https://www.googleapis.com/auth/drive"],
  });
  res.redirect(url);
});

app.get("/oauth-callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send("No code received");
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const oauth2 = new google.auth.OAuth2(clientId, clientSecret, RAILWAY_REDIRECT);
  try {
    const { tokens } = await oauth2.getToken(code);
    res.send(`<h2>✅ Token generated!</h2>
<p>Copy this into Railway as <strong>GOOGLE_REFRESH_TOKEN</strong>:</p>
<textarea rows="4" cols="80" onclick="this.select()">${tokens.refresh_token}</textarea>
<p>Then redeploy on Railway and test again.</p>`);
  } catch (e) {
    res.send("Error: " + e.message);
  }
});

// ── DIAGNOSTIC ─────────────────────────────────────────────────
app.get("/api/drive-test", async (req, res) => {
  try {
    const clientId = process.env.GOOGLE_CLIENT_ID;
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
    const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
    const oauth2 = new google.auth.OAuth2(clientId, clientSecret, RAILWAY_REDIRECT);
    oauth2.setCredentials({ refresh_token: refreshToken });
    const tokenResponse = await oauth2.getAccessToken();
    res.json({ ok: true, tokenOk: !!tokenResponse.token });
  } catch (err) {
    res.json({ ok: false, error: err.message, details: err.response && err.response.data });
  }
});

// ── START ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`GR Stock server running on http://localhost:${PORT}`);
});
