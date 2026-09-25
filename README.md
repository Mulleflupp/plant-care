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

1. Open [apps-script-code.gs in this repo](https://github.com/mulleflupp/plant-care/blob/main/apps-script-code.gs).
2. Click the **copy icon** in the top-right of the file view (or click "Raw" and select-all/copy
   the page) to copy the whole script.
3. In your new Sheet, go to **Extensions > Apps Script**.
4. Delete any placeholder code in the editor, and paste in the script you just copied.
5. Near the top of the script, change these two lines to your own values:
   - `PLANTBOOK_CLIENT_ID` and `PLANTBOOK_SECRET` — see Step 3 below for where to get these.
   - `APP_SHARED_SECRET` — make up any random string. This acts like a password so random
     people on the internet can't write to your sheet. Keep it private, but you'll need to
     remember it for Step 5.

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

1. Open this link in Chrome on your phone: **[APP URL GOES HERE]**
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
- The backend script is versioned (see the comment block and `BACKEND_VERSION` at the top of
  `apps-script-code.gs`), same as `index.html`'s `APP_VERSION` — the app's Setup tab shows both
  automatically, so you can tell at a glance if what's deployed matches the latest in the repo.
  If it's fallen behind, re-copy the file from Step 2 into your Apps Script editor and redeploy.
