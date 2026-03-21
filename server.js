const express = require("express");
const cors = require("cors");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");

const app = express();
const PORT = 3000;

const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const DATA_DIR = path.join(ROOT_DIR, "data");
const UPLOAD_DIR = path.join(ROOT_DIR, "uploads");
const EXCEL_FILE = path.join(DATA_DIR, "stock.xlsx");
const SHEET_NAME = "Items";

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended: true }));
app.use("/uploads", express.static(UPLOAD_DIR));
app.use(express.static(PUBLIC_DIR));

/* =========================
   HELPERS
========================= */
function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}

function normalizeMoney(value) {
  const cleaned = String(value || "")
    .replace(/[^\d.-]/g, "")
    .trim();

  if (!cleaned) return "";

  const num = Number(cleaned);
  if (Number.isNaN(num)) return "";

  return String(num);
}

function normalizeCurrency(value) {
  return String(value || "").trim().toUpperCase();
}

function safeCategoryFolderName(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Uncategorized";

  return raw.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}

function safeFileBaseName(value) {
  return String(value || "")
    .trim()
    .replace(/[^a-zA-Z0-9-_ ]/g, "_")
    .replace(/\s+/g, "_");
}

function parsePhotos(value) {
  if (!value) return [];

  if (Array.isArray(value)) {
    return value.filter(Boolean).map(v => String(v).trim()).filter(Boolean);
  }

  const str = String(value).trim();
  if (!str) return [];

  try {
    const parsed = JSON.parse(str);
    if (Array.isArray(parsed)) {
      return parsed.filter(Boolean).map(v => String(v).trim()).filter(Boolean);
    }
  } catch (err) {
    // fallback below
  }

  return str
    .split("|")
    .map(v => String(v).trim())
    .filter(Boolean);
}

function photosToCellValue(photos) {
  return JSON.stringify(
    (Array.isArray(photos) ? photos : [])
      .filter(Boolean)
      .map(v => String(v).trim())
      .filter(Boolean)
  );
}

function getRelativeUploadPath(filePath) {
  const relative = path.relative(ROOT_DIR, filePath).replace(/\\/g, "/");
  return `/${relative}`;
}

function uniqueArray(arr) {
  return [...new Set((arr || []).filter(Boolean).map(v => String(v).trim()).filter(Boolean))];
}

function safeUnlink(filePath) {
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (err) {
    console.warn("Failed to delete file:", filePath, err.message);
  }
}

function toAbsoluteFromRelative(relativePath) {
  const cleaned = String(relativePath || "").replace(/^\/+/, "");
  return path.join(ROOT_DIR, cleaned);
}

/* =========================
   MULTER
========================= */
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const category = safeCategoryFolderName(req.body.category);
    const categoryDir = path.join(UPLOAD_DIR, category);

    if (!fs.existsSync(categoryDir)) {
      fs.mkdirSync(categoryDir, { recursive: true });
    }

    cb(null, categoryDir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname || "");
    const category = safeFileBaseName(safeCategoryFolderName(req.body.category || "Item"));
    const base = safeFileBaseName(path.basename(file.originalname || "photo", ext)) || "photo";
    cb(null, `${category}-${Date.now()}-${base}${ext}`);
  }
});

const upload = multer({
  storage,
  limits: {
    files: 20,
    fileSize: 10 * 1024 * 1024
  }
});

/* =========================
   EXCEL HEADERS
========================= */
const HEADERS = [
  "id",
  "itemNo",
  "itemName",
  "category",
  "client",
  "capitalCurrency",
  "capitalPrice",
  "wholesaleCurrency",
  "wholesalePrice",
  "descEn",
  "descLa",
  "photos",
  "createdAt"
];

async function getWorkbookAndSheet() {
  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }

  let sheet = workbook.getWorksheet(SHEET_NAME);

  if (!sheet) {
    sheet = workbook.addWorksheet(SHEET_NAME);
    sheet.addRow(HEADERS);
  }

  if (sheet.rowCount === 0) {
    sheet.addRow(HEADERS);
  }

  const headerRow = sheet.getRow(1);
  const currentHeaders = [];

  for (let i = 1; i <= Math.max(headerRow.cellCount, HEADERS.length); i++) {
    currentHeaders.push(String(headerRow.getCell(i).value || "").trim());
  }

  const isExactHeader =
    currentHeaders.length >= HEADERS.length &&
    HEADERS.every((h, i) => normalizeHeader(currentHeaders[i]) === normalizeHeader(h));

  if (!isExactHeader) {
    for (let i = 0; i < HEADERS.length; i++) {
      headerRow.getCell(i + 1).value = HEADERS[i];
    }
    headerRow.commit();
  }

  return { workbook, sheet };
}

async function saveWorkbook(workbook) {
  await workbook.xlsx.writeFile(EXCEL_FILE);
}

function getHeaderMap(sheet) {
  const headerRow = sheet.getRow(1);
  const map = {};

  for (let i = 1; i <= Math.max(headerRow.cellCount, HEADERS.length); i++) {
    const key = normalizeHeader(headerRow.getCell(i).value);
    if (key) map[key] = i;
  }

  return map;
}

function rowToItem(row, headerMap) {
  const getVal = name => {
    const idx = headerMap[normalizeHeader(name)];
    if (!idx) return "";

    const cell = row.getCell(idx).value;

    if (cell == null) return "";
    if (typeof cell === "object" && cell.text) return String(cell.text).trim();
    if (typeof cell === "object" && cell.richText) {
      return cell.richText.map(x => x.text).join("").trim();
    }

    return String(cell).trim();
  };

  return {
    id: getVal("id"),
    itemNo: getVal("itemNo"),
    itemName: getVal("itemName"),
    category: getVal("category"),
    client: getVal("client"),
    capitalCurrency: getVal("capitalCurrency"),
    capitalPrice: getVal("capitalPrice"),
    wholesaleCurrency: getVal("wholesaleCurrency"),
    wholesalePrice: getVal("wholesalePrice"),
    descEn: getVal("descEn"),
    descLa: getVal("descLa"),
    photos: parsePhotos(getVal("photos")),
    createdAt: getVal("createdAt")
  };
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

function getAllItemsFromSheet(sheet, headerMap) {
  const items = [];

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);

    const values = Array.isArray(row.values) ? row.values : [];
    const isEmpty = values.every(v => v == null || String(v).trim() === "");
    if (isEmpty) continue;

    const item = rowToItem(row, headerMap);
    if (!item.id && !item.itemNo && !item.itemName) continue;

    items.push(item);
  }

  return items;
}

/* =========================
   GET ITEMS
========================= */
app.get("/api/items", async (req, res) => {
  try {
    const { sheet } = await getWorkbookAndSheet();
    const headerMap = getHeaderMap(sheet);
    const items = getAllItemsFromSheet(sheet, headerMap);

    items.sort((a, b) => {
      const aNum = Number(String(a.itemNo || "").replace(/\D/g, "")) || 0;
      const bNum = Number(String(b.itemNo || "").replace(/\D/g, "")) || 0;
      return bNum - aNum;
    });

    return res.json({
      success: true,
      items
    });
  } catch (error) {
    console.error("GET /api/items error:", error);
    return res.status(500).json({
      success: false,
      message: error.message || "Failed to load items"
    });
  }
});

/* =========================
   SAVE ITEM
========================= */
app.post("/api/items", upload.array("photos", 20), async (req, res) => {
  try {
    const { workbook, sheet } = await getWorkbookAndSheet();
    const headerMap = getHeaderMap(sheet);
    const existingItems = getAllItemsFromSheet(sheet, headerMap);

    const photoPaths = (req.files || []).map(file => getRelativeUploadPath(file.path));

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
      return res.status(400).json({
        success: false,
        message: "Item Name is required"
      });
    }

    sheet.addRow([
      item.id,
      item.itemNo,
      item.itemName,
      item.category,
      item.client,
      item.capitalCurrency,
      item.capitalPrice,
      item.wholesaleCurrency,
      item.wholesalePrice,
      item.descEn,
      item.descLa,
      photosToCellValue(item.photos),
      item.createdAt
    ]);

    await saveWorkbook(workbook);

    return res.json({
      success: true,
      item
    });
  } catch (error) {
    console.error("POST /api/items error:", error);
    return res.status(500).json({
      success: false,
      message: error.message || "Failed to save item"
    });
  }
});

/* =========================
   UPDATE ITEM
========================= */
app.put("/api/items/:id", upload.array("photos", 20), async (req, res) => {
  try {
    const targetId = String(req.params.id || "").trim();

    if (!targetId) {
      return res.status(400).json({
        success: false,
        message: "Item ID is required"
      });
    }

    const { workbook, sheet } = await getWorkbookAndSheet();
    const headerMap = getHeaderMap(sheet);

    let targetRow = null;
    let targetItem = null;

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);

      const values = Array.isArray(row.values) ? row.values : [];
      const isEmpty = values.every(v => v == null || String(v).trim() === "");
      if (isEmpty) continue;

      const item = rowToItem(row, headerMap);

      if (String(item.id || "").trim() === targetId) {
        targetRow = row;
        targetItem = item;
        break;
      }
    }

    if (!targetRow || !targetItem) {
      return res.status(404).json({
        success: false,
        message: "Item not found"
      });
    }

    const keptPhotos = parsePhotos(req.body.existingPhotos);
    const newPhotos = (req.files || []).map(file => getRelativeUploadPath(file.path));
    const finalPhotos = uniqueArray([...keptPhotos, ...newPhotos]);

    if (finalPhotos.length > 20) {
      return res.status(400).json({
        success: false,
        message: "Maximum 20 photos allowed"
      });
    }

    const oldPhotos = parsePhotos(targetItem.photos);
    const removedPhotos = oldPhotos.filter(photo => !finalPhotos.includes(photo));

    const updatedItem = {
      id: targetItem.id,
      itemNo: targetItem.itemNo,
      itemName: String(req.body.itemName || "").trim(),
      category: String(req.body.category || "").trim(),
      client: String(req.body.client || "").trim(),
      capitalCurrency: normalizeCurrency(req.body.capitalCurrency || targetItem.capitalCurrency || "LAK"),
      capitalPrice: normalizeMoney(req.body.capitalPrice),
      wholesaleCurrency: normalizeCurrency(req.body.wholesaleCurrency || targetItem.wholesaleCurrency || "LAK"),
      wholesalePrice: normalizeMoney(req.body.wholesalePrice),
      descEn: String(req.body.descEn || "").trim(),
      descLa: String(req.body.descLa || "").trim(),
      photos: finalPhotos,
      createdAt: targetItem.createdAt || new Date().toISOString()
    };

    if (!updatedItem.itemName) {
      return res.status(400).json({
        success: false,
        message: "Item Name is required"
      });
    }

    targetRow.getCell(headerMap[normalizeHeader("id")]).value = updatedItem.id;
    targetRow.getCell(headerMap[normalizeHeader("itemNo")]).value = updatedItem.itemNo;
    targetRow.getCell(headerMap[normalizeHeader("itemName")]).value = updatedItem.itemName;
    targetRow.getCell(headerMap[normalizeHeader("category")]).value = updatedItem.category;
    targetRow.getCell(headerMap[normalizeHeader("client")]).value = updatedItem.client;
    targetRow.getCell(headerMap[normalizeHeader("capitalCurrency")]).value = updatedItem.capitalCurrency;
    targetRow.getCell(headerMap[normalizeHeader("capitalPrice")]).value = updatedItem.capitalPrice;
    targetRow.getCell(headerMap[normalizeHeader("wholesaleCurrency")]).value = updatedItem.wholesaleCurrency;
    targetRow.getCell(headerMap[normalizeHeader("wholesalePrice")]).value = updatedItem.wholesalePrice;
    targetRow.getCell(headerMap[normalizeHeader("descEn")]).value = updatedItem.descEn;
    targetRow.getCell(headerMap[normalizeHeader("descLa")]).value = updatedItem.descLa;
    targetRow.getCell(headerMap[normalizeHeader("photos")]).value = photosToCellValue(updatedItem.photos);
    targetRow.getCell(headerMap[normalizeHeader("createdAt")]).value = updatedItem.createdAt;

    targetRow.commit();
    await saveWorkbook(workbook);

    removedPhotos.forEach(photo => {
      const abs = toAbsoluteFromRelative(photo);
      safeUnlink(abs);
    });

    return res.json({
      success: true,
      item: updatedItem
    });
  } catch (error) {
    console.error("PUT /api/items/:id error:", error);
    return res.status(500).json({
      success: false,
      message: error.message || "Failed to update item"
    });
  }
});

/* =========================
   DOWNLOAD EXCEL
========================= */
app.get("/api/download-excel", async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      const { workbook } = await getWorkbookAndSheet();
      await saveWorkbook(workbook);
    }

    return res.download(EXCEL_FILE, "stock.xlsx");
  } catch (error) {
    console.error("GET /api/download-excel error:", error);
    return res.status(500).json({
      success: false,
      message: error.message || "Failed to download Excel"
    });
  }
});

/* =========================
   START SERVER
========================= */
app.listen(PORT, () => {
  console.log(`GR Stock server running on http://localhost:${PORT}`);
});

/* =========================
   DELETE
========================= */
app.delete("/api/items/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const workbook = new ExcelJS.Workbook();

    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        message: "Excel file not found"
      });
    }

    await workbook.xlsx.readFile(EXCEL_FILE);
    const sheet = workbook.getWorksheet(SHEET_NAME);

    if (!sheet) {
      return res.status(404).json({
        success: false,
        message: "Sheet not found"
      });
    }

    let rowToDelete = -1;
    let itemRow = null;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header

      const rowId = String(row.getCell(1).value || "").trim(); // column A = id
      if (rowId === String(id).trim()) {
        rowToDelete = rowNumber;
        itemRow = row;
      }
    });

    if (rowToDelete === -1 || !itemRow) {
      return res.status(404).json({
        success: false,
        message: `Item not found for id: ${id}`
      });
    }

    // photos column
    // current GR Stock structure:
    // 1 id
    // 2 itemNo
    // 3 sku
    // 4 itemName
    // 5 category
    // 6 client
    // 7 capitalPrice
    // 8 wholesalePrice
    // 9 retailPrice
    // 10 descEn
    // 11 descLa
    // 12 photos
    // 13 createdAt
    const photoCell = itemRow.getCell(12).value;

    let photoList = [];

    if (photoCell) {
      try {
        if (typeof photoCell === "string") {
          const parsed = JSON.parse(photoCell);
          photoList = Array.isArray(parsed) ? parsed : [];
        }
      } catch (e) {
        photoList = String(photoCell)
          .split(",")
          .map(v => v.trim())
          .filter(Boolean);
      }
    }

    photoList.forEach(photoPath => {
      try {
        if (!photoPath) return;

        const filename = path.basename(photoPath);
        const fullPath = path.join(UPLOAD_DIR, filename);

        if (fs.existsSync(fullPath)) {
          fs.unlinkSync(fullPath);
        }
      } catch (err) {
        console.error("Failed to delete photo:", err);
      }
    });

    sheet.spliceRows(rowToDelete, 1);
    await workbook.xlsx.writeFile(EXCEL_FILE);

    return res.json({
      success: true,
      message: "Item deleted successfully"
    });
  } catch (err) {
    console.error("DELETE ITEM ERROR:", err);
    return res.status(500).json({
      success: false,
      message: err.message || "Server error while deleting item"
    });
  }
});