/****************************************************
 * API
 ****************************************************/

const Api = {
  baseUrl: "https://gimlety-persuadably-archie.ngrok-free.dev",

  getUrl(path) {
    if (!path.startsWith("/")) path = "/" + path;
    return this.baseUrl + path;
  },

  fetchJson(path) {
    const res = UrlFetchApp.fetch(this.getUrl(path), {
      headers: { "ngrok-skip-browser-warning": "true" },
    });

    const text = res.getContentText();

    try {
      return JSON.parse(text);
    } catch (e) {
      throw new Error("Invalid JSON:\n" + text);
    }
  },
};

/****************************************************
 * Base Sheet
 ****************************************************/

const BaseSheet = {
  getSheet() {
    const sheet = SpreadsheetApp.getActive().getSheetByName(this.name);

    if (!sheet) {
      throw new Error(`Sheet not found: ${this.name}`);
    }

    return sheet;
  },

  getLastCol() {
    return this.getSheet().getLastColumn();
  },

  getLastRow() {
    return this.getSheet().getLastRow();
  },

  getHeaders() {
    return this.getSheet()
      .getRange(this.layout.headerRow, 1, 1, this.getLastCol())
      .getValues()[0];
  },

  getDataBlock() {
    const numRows = this.getLastRow() - this.layout.headerRow + 1;

    if (numRows <= 0) return [];

    return this.getSheet()
      .getRange(this.layout.headerRow, 1, numRows, this.getLastCol())
      .getValues();
  },
};

/****************************************************
 * Read / Write Sheets
 ****************************************************/

const ReadOnlySheet = Object.create(BaseSheet);

const WritableSheet = Object.create(BaseSheet);

WritableSheet.clear = function () {
  this.getSheet().clearContents();
};

WritableSheet.write = function (headers, rows, startRow = 1) {
  const sheet = this.getSheet();

  sheet.clearContents();

  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet
      .getRange(startRow + 1, 1, rows.length, headers.length)
      .setValues(rows);
  }
};

WritableSheet.formatNumbers = function (
  rowCount,
  colCount,
  startRow = 2,
  startCol = 2,
) {
  if (rowCount <= 0 || colCount <= 0) return;

  this.getSheet()
    .getRange(startRow, startCol, rowCount, colCount)
    .setNumberFormat("#,##0.####");
};

/****************************************************
 * Factory
 ****************************************************/

function createSheet(name, layout = {}, type = "read") {
  const base = type === "write" ? WritableSheet : ReadOnlySheet;

  const obj = Object.create(base);
  obj.name = name;
  obj.layout = layout;

  return obj;
}

/****************************************************
 * Sheet Definitions
 ****************************************************/

const LinnunData = createSheet(
  "LinnunData",
  {
    statusCell: "A1",
    weightsRow: 3,
    headerRow: 4,
    attrStart: "FP",
    attrEnd: "Items/Fragments",
  },
  "read",
);

LinnunData.getWeights = function () {
  return this.getSheet()
    .getRange(this.layout.weightsRow, 1, 1, this.getLastCol())
    .getValues()[0];
};

LinnunData.getAttributeOrder = function () {
  const headers = this.getHeaders();

  const startIdx = headers.indexOf(this.layout.attrStart);
  const endIdx = headers.indexOf(this.layout.attrEnd);

  if (startIdx === -1 || endIdx === -1) {
    throw new Error("Attribute bounds not found");
  }

  return headers.slice(startIdx, endIdx);
};

LinnunData.setStatus = function (msg) {
  const ts = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );
  this.getSheet()
    .getRange(this.layout.statusCell)
    .setValue(msg + " at " + ts);
  console.log("Status:", msg);
};

/****************************************************
 * Data Push
 ****************************************************/

function pushLinnunDataToDuckDB() {
  const values = LinnunData.getDataBlock();

  if (!values.length) {
    return { ok: false, body: "No data", rows: 0 };
  }

  const res = UrlFetchApp.fetch(Api.getUrl("/ingest_linnun"), {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ rows: values }),
    muteHttpExceptions: true,
  });

  return {
    ok: res.getResponseCode() === 200,
    code: res.getResponseCode(),
    body: res.getContentText(),
    rows: values.length,
  };
}

/****************************************************
 * Weights Push
 ****************************************************/

function pushLinnunWeightsToDuckDB() {
  const headers = LinnunData.getHeaders();
  const weights = LinnunData.getWeights();

  const startIdx = headers.indexOf("FP");
  const endIdx = headers.indexOf("Items/Fragments");

  if (startIdx === -1 || endIdx === -1) {
    LinnunData.setStatus("❌ Could not find attribute boundaries");
    return { ok: false };
  }

  const rows = [];

  for (let i = startIdx; i < endIdx; i++) {
    const attr = String(headers[i]).trim();
    const raw = weights[i];

    if (!attr) continue;
    if (raw === "" || raw === null) continue;

    const value = Number(String(raw).replace(/,/g, ""));
    if (isNaN(value) || value === 0) continue;

    rows.push(["Linnun", "attributes", attr, value]);
  }

  const res = UrlFetchApp.fetch(Api.getUrl("/ingest_weights"), {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ rows }),
    muteHttpExceptions: true,
  });

  return {
    ok: res.getResponseCode() === 200,
    code: res.getResponseCode(),
    body: res.getContentText(),
    pushed: rows.length,
  };
}

/****************************************************
 * Refresh
 ****************************************************/

function refreshLinnunDuckDB() {
  LinnunData.setStatus("⏳ Refresh started " + new Date());
  SpreadsheetApp.flush();

  const data = pushLinnunDataToDuckDB();
  if (!data.ok) {
    LinnunData.setStatus("❌ Data push to DuckDB failed");
    return;
  }

  const weights = pushLinnunWeightsToDuckDB();
  if (!weights.ok) {
    LinnunData.setStatus("⚠️ Weights failed");
    return;
  }

  LinnunData.setStatus("✅ Refresh DuckDB OK");
}

const EfficiencyView = createSheet(
  "EfficiencyView",
  {
    profileCell: "C2",
    rankModeCell: "E2",
    headerRow: 3,
    dataStartRow: 4,
  },
  "write",
);

EfficiencyView.columns = [
  { key: "Building", label: "Building" },
  { key: "linnun_rank", label: "Linnun" },
  { key: "efficiency_rank", label: "Efficiency" },
  { key: "combined_rank", label: "Combined" },
  { key: "ln_efficiency", label: "Ln Eff" },
  { key: "efficiency", label: "Efficiency" },
  { key: "total_weight", label: "Weight" },
];

EfficiencyView.rankMap = {
  Combined: "combined_rank",
  Efficiency: "efficiency_rank",
  Linnun: "linnun_rank",
};

EfficiencyView.getProfile = function () {
  return this.getSheet().getRange(this.layout.profileCell).getValue();
};

EfficiencyView.getRankMode = function () {
  return this.getSheet().getRange(this.layout.rankModeCell).getValue();
};

/****************************************************
 * Efficiency Load
 ****************************************************/

function loadEfficiency() {
  const profile = EfficiencyView.getProfile() || "TedMilitary";
  const mode = EfficiencyView.getRankMode() || "Combined";

  if (!profile) {
    throw new Error("No profile selected");
  }

  const json = Api.fetchJson(
    `/efficiency?profile=${encodeURIComponent(profile)}`,
  );

  if (json.status !== "ok") {
    throw new Error("API error");
  }

  const { columns, rows } = json;

  Logger.log(columns);

  // -----------------------------
  // Build column index map
  // -----------------------------
  const indexMap = {};
  columns.forEach((c, i) => {
    indexMap[c] = i;
  });

  // -----------------------------
  // Reorder rows using view schema
  // -----------------------------
  const orderedRows = rows.map((row) =>
    EfficiencyView.columns.map((c) => row[indexMap[c.key]]),
  );

  Logger.log(indexMap);

  // -----------------------------
  // Determine sort column
  // -----------------------------
  const sortKey = EfficiencyView.rankMap[mode] || "combined_rank";

  const sortIndex = EfficiencyView.columns.findIndex((c) => c.key === sortKey);

  // Defensive (should never happen, but safe)
  if (sortIndex === -1) {
    throw new Error(`Sort column not found: ${sortKey}`);
  }

  // -----------------------------
  // Sort rows
  // -----------------------------
  orderedRows.sort((a, b) => a[sortIndex] - b[sortIndex]);

  // -----------------------------
  // Headers
  // -----------------------------
  const headers = EfficiencyView.columns.map((c) => c.label);

  // -----------------------------
  // Write to sheet (preserve selectors)
  // -----------------------------
  const sheet = EfficiencyView.getSheet();

  const headerRow = EfficiencyView.layout.headerRow;
  const dataStart = EfficiencyView.layout.dataStartRow;

  const lastRow = sheet.getLastRow();

  // Clear only output area
  if (lastRow >= headerRow) {
    sheet
      .getRange(headerRow, 1, lastRow - headerRow + 1, sheet.getLastColumn())
      .clearContent();
  }

  // Write headers
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);

  // Write data
  if (orderedRows.length > 0) {
    sheet
      .getRange(dataStart, 1, orderedRows.length, headers.length)
      .setValues(orderedRows);
  }

  console.log(
    `Efficiency loaded: ${orderedRows.length} rows | profile=${profile} | mode=${mode}`,
  );
}

const ConfigWeights = createSheet("ConfigWeights", {}, "write");

/****************************************************
 * Config Weights Load
 ****************************************************/

function loadConfigWeights() {
  const json = Api.fetchJson("/config_weights");

  const rows = json.rows;

  const profiles = [...new Set(rows.map((r) => r[0]))];

  const lookup = {};
  rows.forEach(([profile, mode, attr, weight]) => {
    lookup[`${profile}|${attr}`] = weight;
  });

  const ATTRIBUTE_ORDER = LinnunData.getAttributeOrder();
  const attributeSet = new Set(ATTRIBUTE_ORDER);

  const itemNames = [
    ...new Set(rows.filter((r) => r[1] === "items").map((r) => r[2])),
  ]
    .filter((n) => !attributeSet.has(n))
    .sort();

  const output = [];

  ATTRIBUTE_ORDER.forEach((attr) => {
    const row = [attr];
    profiles.forEach((p) => row.push(lookup[`${p}|${attr}`] || ""));
    output.push(row);
  });

  itemNames.forEach((attr) => {
    const row = [attr];
    profiles.forEach((p) => row.push(lookup[`${p}|${attr}`] || ""));
    output.push(row);
  });

  const headers = ["Attribute", ...profiles];

  ConfigWeights.write(headers, output);
  ConfigWeights.formatNumbers(output.length, profiles.length);
}
