// ===== Plant Care App — Backend =====
// Version: 1.2.0
// Last updated: 2026-09-26
// Change log (most recent first):
//   1.2.0 (2026-09-26) — Added getDashboardData, returning Plants + Readings in one
//                        request instead of two, to cut Dashboard load time (each
//                        separate request pays Apps Script's own per-request
//                        overhead, so two requests run roughly 2x as slow as one,
//                        not the same). getPlants/getReadings are kept as-is for
//                        anything else still using them individually.
//   1.1.3 (2026-09-25) — logReading's fallback timestamp (used only if the frontend
//                        doesn't send one) now uses the Sheet's own timezone via
//                        Utilities.formatDate(..., getSpreadsheetTimeZone(), ...)
//                        instead of new Date().toISOString(), which was always UTC.
//                        Matches the frontend's v1.2.0 local-timestamp fix.
//   1.1.2 (2026-09-25) — doGet() is now wrapped in try/catch too (doPost() already was),
//                        so a refresh/getPlants/getReadings failure always returns JSON
//                        instead of Apps Script's own HTML error page — the same "Unexpected
//                        token '<' ... <!DOCTYPE" symptom, but triggered by a GET this time.
//   1.1.1 (2026-09-25) — Removed the leftover "SensorName" column from logReading's
//                        appendRow. It was left over from an abandoned sensor-naming
//                        feature: the frontend stopped sending sensorName a while ago,
//                        but this backend kept writing a blank slot for it, which
//                        shifted every value one column to the right of its header
//                        if the sheet's header row didn't also have a SensorName
//                        column. See note below on fixing an already-shifted sheet.
//   1.1.0 (2026-09-25) — doGet()'s default response (no action param) now reports
//                        BACKEND_VERSION, so the app's Setup tab can show which
//                        backend version is actually deployed.
//   1.0.0 (2026-09-25) — Baseline version marker added; doPost() now wrapped in try/catch
//                        so failures always return JSON instead of an HTML error page.
var BACKEND_VERSION = "1.2.0";

// ===== CONFIG =====
// Fill these in after you register at https://open.plantbook.io/apikey/show/
var PLANTBOOK_CLIENT_ID = "YOUR_CLIENT_ID";
var PLANTBOOK_SECRET = "YOUR_SECRET";
// A shared secret only your app knows, so random people can't write to your sheet
// even though the deployment is set to "Anyone". Change this to something unique.
var APP_SHARED_SECRET = "change-me-to-something-random";

var READINGS_SHEET = "Readings";
var PLANTS_SHEET = "Plants";

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === READINGS_SHEET) {
      sheet.appendRow(["Timestamp", "PlantID", "PlantName", "Moisture", "Light", "Temperature", "Conductivity"]);
    } else if (name === PLANTS_SHEET) {
      sheet.appendRow(["PlantID", "PlantName", "Species", "MoistureMin", "MoistureMax", "LightMin", "LightMax",
                        "TempMin", "TempMax", "ConductivityMin", "ConductivityMax", "Notes", "NFC Tag"]);
    }
  }
  return sheet;
}

function getPlantbookToken_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("plantbook_token");
  if (cached) return cached;

  var resp = UrlFetchApp.fetch("https://open.plantbook.io/api/v1/token/", {
    method: "post",
    payload: {
      grant_type: "client_credentials",
      client_id: PLANTBOOK_CLIENT_ID,
      client_secret: PLANTBOOK_SECRET
    },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (data.access_token) {
    cache.put("plantbook_token", data.access_token, data.expires_in - 60);
    return data.access_token;
  }
  throw new Error("Could not get OpenPlantbook token: " + resp.getContentText());
}

function searchSpecies_(query) {
  var token = getPlantbookToken_();
  var resp = UrlFetchApp.fetch(
    "https://open.plantbook.io/api/v1/plant/search?alias=" + encodeURIComponent(query),
    { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true }
  );
  return JSON.parse(resp.getContentText());
}

function getSpeciesDetail_(pid) {
  var token = getPlantbookToken_();
  var resp = UrlFetchApp.fetch(
    "https://open.plantbook.io/api/v1/plant/detail/" + encodeURIComponent(pid) + "/",
    { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true }
  );
  return JSON.parse(resp.getContentText());
}

function ensureNfcTagColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf("NFC Tag");
  if (idx === -1) {
    sheet.getRange(1, lastCol + 1).setValue("NFC Tag");
    idx = lastCol; // 0-based index of the new column
  }
  return idx;
}

function markNfcTagWritten_(plantId) {
  var sheet = getSheet_(PLANTS_SHEET);
  var colIdx = ensureNfcTagColumn_(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === plantId) {
      sheet.getRange(i + 1, colIdx + 1).setValue("Y");
      return true;
    }
  }
  return false;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.secret !== APP_SHARED_SECRET) {
      return jsonOut_({ status: "error", message: "unauthorized" });
    }

    if (data.action === "logReading") {
      var rs = getSheet_(READINGS_SHEET);
      // The frontend always sends its own local-time timestamp, so this fallback only
      // fires for a direct/manual API call without one. It uses the Sheet's own
      // timezone (Brussels, if you've set that under File > Settings) rather than UTC,
      // to match what the frontend sends.
      var fallbackTimestamp = Utilities.formatDate(
        new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"
      );
      rs.appendRow([
        data.timestamp || fallbackTimestamp,
        data.plantId || "",
        data.plantName || "",
        data.moisture, data.light, data.temperature, data.conductivity
      ]);
      return jsonOut_({ status: "ok" });
    }

    if (data.action === "searchSpecies") {
      return jsonOut_({ status: "ok", results: searchSpecies_(data.query) });
    }

    if (data.action === "getSpeciesDetail") {
      return jsonOut_({ status: "ok", detail: getSpeciesDetail_(data.pid) });
    }

    if (data.action === "savePlant") {
      var ps = getSheet_(PLANTS_SHEET);
      ensureNfcTagColumn_(ps);
      var t = data.thresholds || {};
      ps.appendRow([
        data.plantId, data.plantName, data.species || "",
        t.moistureMin, t.moistureMax, t.lightMin, t.lightMax,
        t.tempMin, t.tempMax, t.conductivityMin, t.conductivityMax,
        data.notes || "", "N"
      ]);
      return jsonOut_({ status: "ok" });
    }

    if (data.action === "markTagWritten") {
      var found = markNfcTagWritten_(data.plantId);
      return jsonOut_(found ? { status: "ok" } : { status: "error", message: "plant not found" });
    }

    return jsonOut_({ status: "error", message: "unknown action" });
  } catch (err) {
    // Always return JSON, even on an unexpected failure, so the app never
    // has to parse an HTML error page as if it were JSON.
    return jsonOut_({ status: "error", message: "Server error: " + err.message });
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;

    if (action === "getPlants") {
      var ps = getSheet_(PLANTS_SHEET);
      var values = ps.getDataRange().getValues();
      var headers = values.shift();
      var plants = values.map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i]; });
        return obj;
      });
      return jsonOut_({ status: "ok", plants: plants });
    }

    if (action === "getReadings") {
      var rs = getSheet_(READINGS_SHEET);
      var values = rs.getDataRange().getValues();
      var headers = values.shift();
      var readings = values.map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i]; });
        return obj;
      });
      return jsonOut_({ status: "ok", readings: readings });
    }

    // Combines getPlants + getReadings into a single request. The Dashboard used to
    // fire both as separate fetches (even run in parallel via Promise.all on the
    // frontend); each one pays Apps Script's own per-request overhead, so two
    // requests cost roughly double one request's latency, not the same. This halves
    // that cost by returning both sheets from one round trip.
    if (action === "getDashboardData") {
      var plantsSheet = getSheet_(PLANTS_SHEET);
      var plantsValues = plantsSheet.getDataRange().getValues();
      var plantsHeaders = plantsValues.shift();
      var plants = plantsValues.map(function(row) {
        var obj = {};
        plantsHeaders.forEach(function(h, i) { obj[h] = row[i]; });
        return obj;
      });

      var readingsSheet = getSheet_(READINGS_SHEET);
      var readingsValues = readingsSheet.getDataRange().getValues();
      var readingsHeaders = readingsValues.shift();
      var readings = readingsValues.map(function(row) {
        var obj = {};
        readingsHeaders.forEach(function(h, i) { obj[h] = row[i]; });
        return obj;
      });

      return jsonOut_({ status: "ok", plants: plants, readings: readings });
    }

    return jsonOut_({ status: "ok", message: "Plant sensor endpoint is running", version: BACKEND_VERSION });
  } catch (err) {
    // Always return JSON, even on an unexpected failure, so the app never
    // has to parse an HTML error page as if it were JSON.
    return jsonOut_({ status: "error", message: "Server error: " + err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
