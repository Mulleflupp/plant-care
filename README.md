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
      sheet.appendRow(["Timestamp", "PlantID", "PlantName", "SensorName", "Moisture", "Light", "Temperature", "Conductivity"]);
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
    idx = lastCol;
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
  var data = JSON.parse(e.postData.contents);

  if (data.secret !== APP_SHARED_SECRET) {
    return jsonOut_({ status: "error", message: "unauthorized" });
  }

  if (data.action === "logReading") {
    var rs = getSheet_(READINGS_SHEET);
    rs.appendRow([
      data.timestamp || new Date().toISOString(),
      data.plantId || "",
      data.plantName || "",
      data.sensorName || "",
      data.moisture, data.light, data.temperature, data.conductivity
    ]);
    return jsonOut_({ status: "ok" });
  }

  if (data.action === "searchSpecies") {
    try {
      return jsonOut_({ status: "ok", results: searchSpecies_(data.query) });
    } catch (err) {
      return jsonOut_({ status: "error", message: err.message });
    }
  }

  if (data.action === "getSpeciesDetail") {
    try {
      return jsonOut_({ status: "ok", detail: getSpeciesDetail_(data.pid) });
    } catch (err) {
      return jsonOut_({ status: "error", message: err.message });
    }
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
}

function doGet(e) {
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

  return jsonOut_({ status: "ok", message: "Plant sensor endpoint is running" });
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
