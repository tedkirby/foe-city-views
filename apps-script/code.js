const BASE_URL = "https://gimlety-persuadably-archie.ngrok-free.dev";

function setLCStatus_(msg) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("LinnunData");

  if (!sheet) return;

  sheet.getRange("A2").setValue(msg); // status value
  console.log("setLCStatus_ - " + msg);
}

function pushLinnunDataToDuckDB() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("LinnunData");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

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

function pushLinnunWeightsToDuckDB() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("LinnunData");

  const lastCol = sheet.getLastColumn();

  const weights = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const headers = sheet.getRange(3, 1, 1, lastCol).getValues()[0];

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

    if (!attr) {
      continue;
    }

    total++;

    if (raw === "" || raw === null) {
      continue;
    }

    const value = Number(String(raw).replace(/,/g, ""));

    if (isNaN(value) || value === 0) {
      continue; // 👈 skip zero weights
    }

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
    `✅ Refresh OK: ${data.rows} rows, ${weights.pushed}/${weights.total} weights`,
  );
}

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

