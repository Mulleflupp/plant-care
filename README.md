# Plant Care App — Quick Start

This app reads a Flower Care / Mi Flora Bluetooth plant sensor with your phone, logs the data
to your own Google Sheet, and shows you a dashboard with watering/light/fertilizing
recommendations per plant.

**What you need:**
- An Android phone with Chrome
- A Google account
- One or more Flower Care / Mi Flora Bluetooth sensors
- (Optional) NFC stickers, one per plant, if you want tap-to-select

You'll set up your **own** Google Sheet and backend — your data stays separate from anyone
else's, even though you're all using the same app page.

---

## Step 1: Create your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank sheet.
   Name it whatever you like, e.g. "Plant Sensor Log".
2. That's it for this step — the app will create the right tabs and columns automatically
   the first time it talks to the sheet.

## Step 2: Add the backend script

1. In your new Sheet, go to **Extensions > Apps Script**.
2. Delete any placeholder code in the editor, and paste in the script below.
3. Near the top of the script, change these two lines to your own values:
   - `PLANTBOOK_CLIENT_ID` and `PLANTBOOK_SECRET` — see Step 3 below for where to get these.
   - `APP_SHARED_SECRET` — make up any random string. This acts like a password so random
     people on the internet can't write to your sheet. Keep it private, but you'll need to
     remember it for Step 5.

```javascript
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
```

## Step 3: Get a free OpenPlantbook API key

This is what auto-fills recommended moisture/light/temperature ranges per plant species.

1. Go to [open.plantbook.io](https://open.plantbook.io/) and sign up for a free account.
2. Visit [open.plantbook.io/apikey/show/](https://open.plantbook.io/apikey/show/) to find your
   **client_id** and **secret**.
3. Go back to the Apps Script editor and paste them into `PLANTBOOK_CLIENT_ID` and
   `PLANTBOOK_SECRET` at the top of the script.

## Step 4: Deploy the script as a Web App

1. In the Apps Script editor, click **Deploy > New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set "Execute as" to **Me**, and "Who has access" to **Anyone**.
4. Click **Deploy**, and authorize the permissions it asks for (this is normal — it's your own
   script accessing your own sheet).
5. Copy the **Web app URL** it gives you (ends in `/exec`). You'll need this next.

## Step 5: Set up the app on your phone

1. Open this link in Chrome on your phone: ** [https://mulleflupp.github.io/plant-care/] **
2. Go to the **Setup** tab.
3. Paste your Web app URL (from Step 4) into "Apps Script Web App URL".
4. Paste your shared secret (the one you made up in Step 2) into "Shared secret".
5. Tap **Save config**.

## Step 6: Add your plants

1. Still in Setup, under "Add a plant": type a name for your plant (e.g. "Monstera - living room").
2. Search for its species (e.g. "Monstera deliciosa") — pick the best match from the results.
   The app will show you the recommended moisture/light/temp/conductivity ranges it found.
3. Tap **Save plant**. Repeat for each plant.

## Step 7 (optional): Write NFC tags

If you have NFC stickers, stick one on each pot, then in Setup > "Write NFC tag", select the
plant and tap "Write tag" while holding your phone against the sticker. This lets you tap a
tag on the **Log** tab to instantly select that plant instead of choosing manually.

## Step 8: Take a reading

1. Go to the **Log** tab.
2. Select a plant (via NFC tap, or the dropdown).
3. Insert the sensor's probe into that plant's soil.
4. Tap **Connect & read sensor**, and pick your sensor from the Bluetooth list that pops up.
5. Check the values, then tap **Save reading to sheet**.
6. Check the **Dashboard** tab to see plant status and recommendations.

---

**A few notes:**
- Everything is your own data, in your own Google Sheet — nothing is shared between users of
  this app unless you explicitly share your Sheet or your Web app URL/secret with someone.
- Light readings taken after dark are automatically excluded from recommendations (a sensor
  reading light at night will always read low, which isn't meaningful).
- If you ever change the Apps Script code, remember to redeploy it (Deploy > Manage deployments
  > edit > "New version" > Deploy) — just saving the code doesn't push it live.
