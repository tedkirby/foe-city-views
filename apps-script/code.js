const BASE_URL = "https://gimlety-persuadably-archie.ngrok-free.dev";

const SHEET_NAME = "LinnunData";

// layout
const STATUS_CELL = "A1";

const DATA_START_ROW = 2; // IMPORTRANGE starts here
const WEIGHTS_ROW = 3;
const HEADER_ROW = 4;
const FIRST_DATA_ROW = 5;

/****************************************************
 * Helpers
 ****************************************************/

function getSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function setLCStatus_(msg) {
  const sheet = getSheet_();
  if (!sheet) return;

  sheet.getRange(STATUS_CELL).setValue(msg);
  console.log("setLCStatus_ - " + msg);
}

/****************************************************
 * Data push
 ****************************************************/

function pushLinnunDataToDuckDB() {
  const sheet = getSheet_();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const numRows = lastRow - HEADER_ROW + 1;

  if (numRows <= 0) {
    return { ok: false, code: 0, body: "No data", rows: 0 };
  }

  const values = sheet.getRange(HEADER_ROW, 1, numRows, lastCol).getValues();

  const url = BASE_URL + "/ingest_linnun";

  const res = UrlFetchApp.fetch(url, {
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
 * Weights push
 ****************************************************/

function pushLinnunWeightsToDuckDB() {
  const sheet = getSheet_();

  const lastCol = sheet.getLastColumn();

  const weights = sheet.getRange(WEIGHTS_ROW, 1, 1, lastCol).getValues()[0];

  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const profile = "Linnun";
  const mode = "attributes";

  const startIdx = headers.indexOf("FP");
  const endIdx = headers.indexOf("Items/Fragments");

  if (startIdx === -1 || endIdx === -1) {
    setLCStatus_("❌ Could not find FP or Items/Fragments");
    return { ok: false };
  }

  const rows = [];
  let total = 0;

  for (let i = startIdx; i < endIdx; i++) {
    const attr = String(headers[i]).trim();
    const raw = weights[i];

    if (!attr) continue;

    total++;

    if (raw === "" || raw === null) continue;

    const value = Number(String(raw).replace(/,/g, ""));

    if (isNaN(value) || value === 0) continue;

    rows.push([profile, mode, attr, value]);
  }

  const url = BASE_URL + "/ingest_weights";

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ rows }),
    muteHttpExceptions: true,
  });

  const ok = res.getResponseCode() === 200;

  return {
    ok,
    code: res.getResponseCode(),
    body: res.getContentText(),
    pushed: rows.length,
    total,
  };
}

/****************************************************
 * Full refresh
 ****************************************************/

function refreshLinnunDuckDB() {
  setLCStatus_("⏳ Full refresh started " + new Date());
  SpreadsheetApp.flush();

  const data = pushLinnunDataToDuckDB();

  if (!data.ok) {
    setLCStatus_(`❌ Data push failed (${data.code}): ${data.body}`);
    return;
  }

  setLCStatus_(`⏳ Data OK (${data.rows} rows), pushing weights...`);

  const weights = pushLinnunWeightsToDuckDB();

  if (!weights.ok) {
    setLCStatus_(
      `⚠️ Data OK (${data.rows}), weights failed (${weights.code}): ${weights.body}`,
    );
    return;
  }

  setLCStatus_(
    `✅ Refresh DuckDB OK: ${data.rows} rows, ${weights.pushed}/${weights.total} weights`,
  );
}

/****************************************************
 * Weights only
 ****************************************************/

function pushWeightsOnly() {
  setLCStatus_("⏳ Pushing weights only...");
  SpreadsheetApp.flush();

  const weights = pushLinnunWeightsToDuckDB();

  if (!weights.ok) {
    setLCStatus_(`❌ Weights failed (${weights.code}): ${weights.body}`);
    return;
  }

  setLCStatus_(`✅ Weights OK: ${weights.pushed}/${weights.total}`);
}

/****************************************************
 * Efficiency load
 ****************************************************/

function loadEfficiency() {
  const url = BASE_URL + "/efficiency?profile=TedMilitary";

  const response = UrlFetchApp.fetch(url, {
    headers: { "ngrok-skip-browser-warning": "true" },
  });

  const json = JSON.parse(response.getContentText());

  if (json.status !== "ok") {
    throw new Error("API error");
  }

  const { columns, rows } = json;

  console.log(columns);

  if (!rows.length) return;

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EfficiencyView");

  sheet.clearContents();

  // headers from backend (ordered)
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);

  // rows already aligned
  sheet.getRange(2, 1, rows.length, columns.length).setValues(rows);
}

function getAttributeOrder_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("LinnunData");

  if (!sheet) {
    throw new Error("LinnunData sheet not found");
  }

  const lastCol = sheet.getLastColumn();

  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const startIdx = headers.indexOf("FP");
  const endIdx = headers.indexOf("Items/Fragments");

  if (startIdx === -1 || endIdx === -1) {
    throw new Error(
      "Could not find FP or Items/Fragments in LinnunData headers",
    );
  }

  return headers.slice(startIdx, endIdx);
}

function loadConfigWeights() {
  const url = BASE_URL + "/config_weights";

  const res = UrlFetchApp.fetch(url, {
    headers: { "ngrok-skip-browser-warning": "true" },
  });

  const text = res.getContentText();

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error("Invalid JSON:\n" + text);
  }

  const rows = json.rows; // [profile, mode, attribute, weight]

  const sheet = SpreadsheetApp.getActive().getSheetByName("ConfigWeights");

  if (!sheet) {
    throw new Error("ConfigWeights sheet not found");
  }

  // -----------------------------
  // Build profile list (columns)
  // -----------------------------
  const profiles = [...new Set(rows.map((r) => r[0]))];

  // -----------------------------
  // Build lookup map
  // -----------------------------
  const lookup = {};
  rows.forEach(([profile, mode, attr, weight]) => {
    lookup[`${profile}|${attr}`] = weight;
  });

  // -----------------------------
  // Get canonical attribute order
  // -----------------------------
  const ATTRIBUTE_ORDER = getAttributeOrder_();

  const attributeSet = new Set(ATTRIBUTE_ORDER);

  // -----------------------------
  // Separate items
  // -----------------------------
  const itemNames = [
    ...new Set(rows.filter((r) => r[1] === "items").map((r) => r[2])),
  ]
    .filter((name) => !attributeSet.has(name))
    .sort();

  // -----------------------------
  // Build output matrix
  // -----------------------------
  const output = [];

  // Attributes first (ordered)
  ATTRIBUTE_ORDER.forEach((attr) => {
    const row = [attr];

    profiles.forEach((profile) => {
      row.push(lookup[`${profile}|${attr}`] || "");
    });

    output.push(row);
  });

  // Then items
  itemNames.forEach((attr) => {
    const row = [attr];

    profiles.forEach((profile) => {
      row.push(lookup[`${profile}|${attr}`] || "");
    });

    output.push(row);
  });

  // -----------------------------
  // Headers
  // -----------------------------
  const headers = ["Attribute", ...profiles];

  // -----------------------------
  // Write to sheet
  // -----------------------------
  sheet.clearContents();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(2, 1, output.length, headers.length).setValues(output);

  // -----------------------------
  // Formatting (commas!)
  // -----------------------------
  if (profiles.length > 0 && output.length > 0) {
    sheet
      .getRange(2, 2, output.length, profiles.length)
      .setNumberFormat("#,##0.####");
  }

  console.log(
    `ConfigWeights loaded: ${output.length} rows, ${profiles.length} profiles`,
  );
}
