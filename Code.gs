// ============================================================
//  GMC Supplemental Feed — AI Description & Category Generator
//  Google Apps Script (Apps Script Editor -> Code.gs)
//
//  SHEET COLUMN LAYOUT (Row 1 = Headers):
//    A  : id                                    (from primary feed / TSV import)
//    B  : title                                 (product title — used for AI prompts)
//    C  : structured_description(content)       (AI-generated description)
//    D  : google_product_category               (AI-generated category)
//    E  : structured_description(digital_source_type)  (always "trained_algorithmic_media")
//    F  : Last Updated                          (Amsterdam timestamp)
//    G  : Char Count                            (description character count)
//
//  REQUIRED SERVICES (Apps Script > Services > + Add):
//    • Merchant API  →  Identifier: MerchantApiProducts  →  Version: products_v1
//    • Merchant API  →  Identifier: MerchantApiAccounts  →  Version: accounts_v1
//
//  HOW TO USE:
//    1. Set GEMINI_API_KEY and MERCHANT_ID below.
//    2. Ensure your sheet has headers in Row 1 (A–G as above).
//    3. Paste IDs in Column A, Titles in Column B.
//    4. Run  startProcessing()  once — then close the tab.
//       The script self-relays via triggers until all rows are done.
// ============================================================

const GEMINI_API_KEY = 'PASTE_YOUR_GEMINI_API_KEY_HERE';
const MERCHANT_ID    = '288388400';

// Models tried in order — falls back if one is busy or unavailable
const MODEL_PRIORITY = [
  'gemini-2.5-flash',
  'gemini-2.0-flash-001',
  'gemini-2.0-flash-lite-001',
];


// ── ENTRY POINTS ─────────────────────────────────────────────

/**
 * Run this ONCE to kick off the full AI generation process.
 * It clears any stale triggers and begins processing from the
 * first row that is missing a description.
 */
function startProcessing() {
  clearAllTriggers_();
  setSheetHeaders_();
  processNextBatch();
}

/**
 * (Optional) Run this ONCE to pull all product IDs + titles
 * directly from your Merchant Center account into the sheet.
 * Skip if you have already pasted IDs/titles from a TSV export.
 */
function syncIdsFromGMC() {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const parent = 'accounts/' + MERCHANT_ID;

  setSheetHeaders_();

  const lastRow    = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String)
      : []
  );

  let pageToken;
  let added = 0;

  try {
    do {
      const res = MerchantApiProducts.Accounts.Products.list(parent, {
        pageToken,
        pageSize: 250,
      });

      if (res && res.products) {
        res.products.forEach(p => {
          const id    = String(p.offerId);
          const title = p.title || '';

          if (!existingIds.has(id)) {
            sheet.appendRow([id, title, '', '', 'trained_algorithmic_media', '', '']);
            existingIds.add(id);
            added++;
          }
        });
      }
      pageToken = res.nextPageToken;
    } while (pageToken);

    SpreadsheetApp.flush();
    console.log(`Sync complete — ${added} new IDs added.`);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert(`Sync complete! Added ${added} new products.`);
    }
  } catch (e) {
    console.error('syncIdsFromGMC error: ' + e.message);
  }
}

/**
 * (Optional) One-time GCP project registration with Merchant Center.
 * Only needs to be run once per project. Safe to run again —
 * it will tell you if already registered.
 */
function registerGCPProject() {
  const parent      = 'accounts/' + MERCHANT_ID + '/developerRegistration';
  const requestBody = { developerEmail: 'dannyhill2302@gmail.com' };

  try {
    const res = MerchantApiAccounts.Accounts.DeveloperRegistration.registerGcp(requestBody, parent);
    console.log('Registration response: ' + JSON.stringify(res));
  } catch (e) {
    // "Already registered" is expected and fine
    console.log('Registration note: ' + e.message);
  }
}


// ── CORE BATCH PROCESSOR ─────────────────────────────────────

/**
 * Processes as many rows as possible within 5 minutes, then
 * schedules itself to resume via a time-based trigger.
 * Called by startProcessing() and directly by triggers.
 * NOTE: must NOT have a trailing underscore — Apps Script
 * cannot trigger private (underscore-suffixed) functions.
 */
function processNextBatch() {
  const sheet      = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const totalRows  = sheet.getLastRow();
  const startTime  = Date.now();
  const TIME_LIMIT = 4 * 60 * 1000; // 4 min — leaves 2 min buffer before the 6-min hard kill

  // In-session cache — avoids duplicate AI calls for size variants
  const descCache = {};

  for (let row = 2; row <= totalRows; row++) {

    // ── Time check FIRST — before any API call ───────────────
    // This guarantees scheduleResume_() always has time to run.
    if (Date.now() - startTime > TIME_LIMIT) {
      console.log(`4-min limit reached at row ${row}. Scheduling resume…`);
      SpreadsheetApp.flush();
      scheduleResume_();
      return;
    }

    // Read ID, Title, existing Description
    const rowData   = sheet.getRange(row, 1, 1, 3).getValues()[0];
    const id        = String(rowData[0]).trim();
    const fullTitle = String(rowData[1]).trim();
    const existDesc = String(rowData[2]).trim();

    // Skip blank IDs
    if (!id) continue;

    // Skip rows that already have a valid description (but retry PENDING)
    if (existDesc && existDesc !== 'PENDING' && !existDesc.startsWith('Error')) continue;

    const baseTitle = getBaseTitle_(fullTitle || id);

    // ── AI Generation ───────────────────────────────────────
    if (!descCache[baseTitle]) {
      console.log(`Row ${row}: generating for "${baseTitle}"`);

      Utilities.sleep(3000);
      const desc = callGemini_(baseTitle, 'description');

      if (!desc || desc.startsWith('Error') || desc.includes('busy')) {
        console.warn(`Row ${row}: model busy — marking PENDING for retry.`);
        sheet.getRange(row, 3).setValue('PENDING');
        continue;
      }

      descCache[baseTitle] = desc;
    }

    // ── Write to Sheet ───────────────────────────────────────
    // Column D (google_product_category) is formula-driven via =SHOPPING_CATEGORY(B2)
    const desc      = descCache[baseTitle];
    const timestamp = new Date().toLocaleString('en-GB', { timeZone: 'Europe/Amsterdam' });

    sheet.getRange(row, 3).setValue(desc);
    sheet.getRange(row, 5).setValue('trained_algorithmic_media');
    sheet.getRange(row, 6).setValue(timestamp);
    sheet.getRange(row, 7).setValue(desc.length);
  }

  // ── Final pass: any PENDING rows left? ──────────────────────
  const pendingRows = findPendingRows_(sheet, totalRows);
  if (pendingRows > 0) {
    console.log(`${pendingRows} PENDING rows remain. Scheduling another pass…`);
    SpreadsheetApp.flush();
    scheduleResume_();
  } else {
    console.log('All rows complete!');
    SpreadsheetApp.flush();
  }
}


// ── AI CONTENT GENERATION ────────────────────────────────────

/**
 * Calls the Gemini API. Tries each model in MODEL_PRIORITY order.
 * Returns the generated text, or an error string.
 */
function callGemini_(title, type) {
  const prompt = type === 'category'
    ? buildCategoryPrompt_(title)
    : buildDescriptionPrompt_(title);

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: type === 'description' ? 8192 : 256,
    },
  };

  const options = {
    method      : 'post',
    contentType : 'application/json',
    payload     : JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  for (const model of MODEL_PRIORITY) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
    try {
      const res  = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const json = JSON.parse(res.getContentText());

      if (code === 200 && json.candidates && json.candidates[0].content) {
        return json.candidates[0].content.parts[0].text.trim();
      }

      if (code === 503 || code === 429) {
        console.warn(`${model} busy (${code}). Trying next…`);
        Utilities.sleep(2000);
        continue;
      }

      console.error(`${model} error ${code}: ${res.getContentText()}`);
    } catch (e) {
      console.warn(`${model} threw: ${e.message}`);
    }
  }

  return 'Error: all models busy';
}

function buildCategoryPrompt_(title) {
  return `Return ONLY the official Google Product Category string for this product: "${title}".
Use the full path with > separators, e.g. "Apparel & Accessories > Clothing > Shirts & Tops".
Do not add any explanation or punctuation outside the category string itself.`;
}

function buildDescriptionPrompt_(title) {
  return `You are a Senior E-commerce SEO Copywriter specialising in the Dutch / EU market.

Write a product description for: "${title}"

REQUIREMENTS:
1. Length: 4,500–5,000 characters (this is the Google Merchant Center 5,000-char description field).
2. FRONTLOAD: The very first sentence (≤150 chars) must state the product name, primary material, specific colour, and one key technical feature.
3. Structure (use these exact headings):
   Overview
   Key Features & Benefits  (use bullet points)
   Design & Aesthetic
   Care & Maintenance  (step-by-step)
   About the Brand / Heritage
4. Keywords: use high-intent, specific terms (e.g. "navy blue" not "blue"; "heavyweight 320 gsm cotton" not "cotton").
5. Audience: Adults / EU market. Mention EU compliance where relevant.
6. Do NOT include: pricing, shipping info, "best seller", "free", ALL CAPS phrases, or competitor names.
7. Do NOT mention specific sizes (this description applies to all size variants).
8. Format: clear line breaks between sections; bullet points for features.`;
}


// ── TITLE CLEANER ─────────────────────────────────────────────

/**
 * Strips size / shoe-size variants from the end of a Shopify title.
 *
 * Examples:
 *   "Obey Bernard Zip Up Sweater Black / M"  →  "Obey Bernard Zip Up Sweater Black"
 *   "Nike Air Force 1 White / US 11"         →  "Nike Air Force 1 White"
 *   "Levi 501 Jeans Blue / W32 L32"          →  "Levi 501 Jeans Blue"
 */
function getBaseTitle_(title) {
  if (!title) return 'Unknown Product';

  const parts = title.split(' / ');
  if (parts.length < 2) return title.trim();

  const last = parts[parts.length - 1].trim().toUpperCase();

  // Clothing sizes, shoe size prefixes, numeric-only (e.g. "32"), W32L32 patterns
  const sizeMarkers = ['US', 'UK', 'EU', 'SIZE', 'XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL', 'W', 'ONE SIZE'];
  const isSize =
    sizeMarkers.some(m => last === m || last.startsWith(m + ' ') || last.startsWith(m + '/')) ||
    /^\d+(\s+\d+)?$/.test(last) ||           // pure numbers: "11", "32 34"
    /^W\d+/.test(last) ||                     // W32, W32L32
    /^EU\s?\d+/.test(last) ||                 // EU 42
    /^US\s?\d+/.test(last) ||                 // US 11
    /^UK\s?\d+/.test(last);                   // UK 9

  if (isSize) {
    parts.pop();
    return parts.join(' / ').trim();
  }

  return title.trim();
}


// ── HELPERS ──────────────────────────────────────────────────

function setSheetHeaders_() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = [
    'id',
    'title',
    'structured_description(content)',
    'google_product_category',
    'structured_description(digital_source_type)',
    'Last Updated',
    'Char Count',
  ];

  // Only write headers if Row 1 Col A is blank
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function clearAllTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

function scheduleResume_() {
  clearAllTriggers_();
  // 'processNextBatch' has no trailing underscore so triggers can call it
  ScriptApp.newTrigger('processNextBatch')
    .timeBased()
    .after(2 * 60 * 1000)   // resume in 2 minutes
    .create();
}

function findPendingRows_(sheet, totalRows) {
  if (totalRows < 2) return 0;
  const descs = sheet.getRange(2, 3, totalRows - 1, 1).getValues().flat();
  return descs.filter(v => String(v).trim() === 'PENDING').length;
}

function isUiAvailable_() {
  try { SpreadsheetApp.getUi(); return true; } catch (e) { return false; }
}


// ── CATEGORY LOOKUP ──────────────────────────────────────────

/**
 * Returns the Google Shopping category for a product title by looking it up
 * in the 'Categories' sheet tab. Size/variant suffixes are stripped before
 * matching (e.g. "Nike Air Force 1 White / US 11" → "Nike Air Force 1 White").
 *
 * Usage in a cell: =SHOPPING_CATEGORY(B2)
 *
 * @param {string} title The product title to look up.
 * @return {string} The Google Shopping category, or "" if not yet mapped.
 * @customfunction
 */
function SHOPPING_CATEGORY(title) {
  if (!title || !String(title).trim()) return '';

  const base      = getBaseTitle_(String(title).trim()).toLowerCase();
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const catSheet  = ss.getSheetByName('Categories');
  if (!catSheet) return '';

  const lastRow = catSheet.getLastRow();
  if (lastRow < 2) return '';

  const data = catSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === base) {
      return String(data[i][1]).trim();
    }
  }
  return '';
}

/**
 * Creates the 'Categories' sheet tab if it doesn't already exist,
 * and writes the column headers in row 1.
 */
function ensureCategoriesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let catSheet = ss.getSheetByName('Categories');
  if (!catSheet) {
    catSheet = ss.insertSheet('Categories');
  }
  // Write headers only if row 1 col A is blank
  if (!catSheet.getRange(1, 1).getValue()) {
    catSheet.getRange(1, 1, 1, 2).setValues([['base_title', 'google_product_category']]);
    catSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return catSheet;
}


// ── DAILY SYNC ───────────────────────────────────────────────

/**
 * Pulls all products from Merchant Center and:
 *   1. Adds any new product rows (id + title) to the main sheet.
 *   2. Adds any new unique base titles to the 'Categories' sheet
 *      with a blank category column, ready for manual entry.
 *
 * Run setupDailySync() once to schedule this automatically every day.
 */
function dailySyncFromGMC() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheet      = ss.getActiveSheet();
  const catSheet   = ensureCategoriesSheet_();
  const parent     = 'accounts/' + MERCHANT_ID;

  setSheetHeaders_();

  // Build sets of existing product IDs and existing base titles
  const lastMainRow = sheet.getLastRow();
  const existingIds = new Set(
    lastMainRow > 1
      ? sheet.getRange(2, 1, lastMainRow - 1, 1).getValues().flat().map(String)
      : []
  );

  const lastCatRow = catSheet.getLastRow();
  const existingBaseTitles = new Set(
    lastCatRow > 1
      ? catSheet.getRange(2, 1, lastCatRow - 1, 1).getValues().flat()
          .map(v => String(v).trim().toLowerCase())
      : []
  );

  let pageToken;
  let addedProducts  = 0;
  let addedBaseTitles = 0;

  try {
    do {
      const res = MerchantApiProducts.Accounts.Products.list(parent, {
        pageToken,
        pageSize: 250,
      });

      if (res && res.products) {
        res.products.forEach(p => {
          const id        = String(p.offerId);
          const title     = p.title || '';
          const baseTitle = getBaseTitle_(title);
          const baseKey   = baseTitle.toLowerCase();

          // Add new product row to main sheet
          if (!existingIds.has(id)) {
            sheet.appendRow([id, title, '', '', 'trained_algorithmic_media', '', '']);
            existingIds.add(id);
            addedProducts++;
          }

          // Add new base title to Categories sheet (blank category)
          if (baseTitle && !existingBaseTitles.has(baseKey)) {
            catSheet.appendRow([baseTitle, '']);
            existingBaseTitles.add(baseKey);
            addedBaseTitles++;
          }
        });
      }
      pageToken = res.nextPageToken;
    } while (pageToken);

    SpreadsheetApp.flush();
    const msg = `Daily sync complete — ${addedProducts} new products, ${addedBaseTitles} new base titles added to Categories sheet.`;
    console.log(msg);
    if (isUiAvailable_()) SpreadsheetApp.getUi().alert(msg);

  } catch (e) {
    console.error('dailySyncFromGMC error: ' + e.message);
  }
}

/**
 * One-time setup: creates the Categories sheet and schedules
 * dailySyncFromGMC() to run automatically every day at ~02:00.
 * Run this once from the Apps Script editor.
 */
function setupDailySync() {
  ensureCategoriesSheet_();

  // Remove any existing dailySyncFromGMC triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailySyncFromGMC')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailySyncFromGMC')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  const msg = 'Daily sync scheduled. dailySyncFromGMC() will run every day at ~02:00.';
  console.log(msg);
  if (isUiAvailable_()) SpreadsheetApp.getUi().alert(msg);
}
