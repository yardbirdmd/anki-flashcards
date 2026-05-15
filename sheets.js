/**
 * sheets.js — Google Sheets API layer
 *
 * ─── SETUP ────────────────────────────────────────────────────────────────
 * 1. Go to https://console.cloud.google.com/ and create a project.
 * 2. Enable the "Google Sheets API" and "Google Drive API".
 * 3. Create an OAuth 2.0 Web Client ID:
 *      Credentials → Create Credentials → OAuth client ID → Web application
 *      Add your app's origin to "Authorised JavaScript origins"
 *      (e.g. https://yourusername.github.io)
 * 4. Create an API Key:
 *      Credentials → Create Credentials → API Key
 *      Restrict it to the Sheets API.
 * 5. Create a Google Sheet. Its URL looks like:
 *      https://docs.google.com/spreadsheets/d/SHEET_ID_IS_HERE/edit
 * 6. The sheet needs two tabs named exactly: "decks" and "cards"
 *      decks tab columns:  name | icon
 *      cards tab columns:  id | deck | front | back | reps | interval | ef | due
 *    Row 1 of each tab must be these exact headers.
 * 7. Fill in the three CONFIG values below.
 * ──────────────────────────────────────────────────────────────────────────
 */

export const CONFIG = {
  // Paste your Sheet ID from the URL:
  SHEET_ID: 'YOUR_SHEET_ID_HERE',

  // Your OAuth 2.0 Web Client ID:
  CLIENT_ID: 'YOUR_CLIENT_ID_HERE.apps.googleusercontent.com',

  // Your API Key (used for the GIS token client):
  API_KEY: 'YOUR_API_KEY_HERE',

  // OAuth scopes needed
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets',

  // Sheet tab names (change if you rename them)
  DECKS_TAB: 'decks',
  CARDS_TAB: 'cards',
};

/* ── State ── */
let tokenClient = null;
let accessToken = null;
let gapiReady = false;
let gisReady = false;

const BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

/* ══════════════════════════════════════════════════════════
   INIT — call once on app start
══════════════════════════════════════════════════════════ */
export function initSheets(onReady) {
  // Load GAPI
  const gapiScript = document.createElement('script');
  gapiScript.src = 'https://apis.google.com/js/api.js';
  gapiScript.onload = () => {
    window.gapi.load('client', async () => {
      await window.gapi.client.init({
        apiKey: CONFIG.API_KEY,
        discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
      });
      gapiReady = true;
      if (gisReady) onReady();
    });
  };
  document.head.appendChild(gapiScript);

  // Load GIS
  const gisScript = document.createElement('script');
  gisScript.src = 'https://accounts.google.com/gsi/client';
  gisScript.onload = () => {
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.CLIENT_ID,
      scope: CONFIG.SCOPES,
      callback: (resp) => {
        if (resp.error) throw resp;
        accessToken = resp.access_token;
        window.gapi.client.setToken({ access_token: accessToken });
      },
    });
    gisReady = true;
    if (gapiReady) onReady();
  };
  document.head.appendChild(gisScript);
}

/* ══════════════════════════════════════════════════════════
   AUTH
══════════════════════════════════════════════════════════ */
export function isSignedIn() { return !!accessToken; }

export function signIn() {
  return new Promise((resolve, reject) => {
    tokenClient.callback = (resp) => {
      if (resp.error) { reject(resp); return; }
      accessToken = resp.access_token;
      window.gapi.client.setToken({ access_token: accessToken });
      resolve();
    };
    tokenClient.requestAccessToken({ prompt: '' });
  });
}

export function signOut() {
  if (accessToken) {
    window.google.accounts.oauth2.revoke(accessToken);
    accessToken = null;
    window.gapi.client.setToken(null);
  }
}

/* ══════════════════════════════════════════════════════════
   LOW-LEVEL SHEETS HELPERS
══════════════════════════════════════════════════════════ */
async function getRange(range) {
  const resp = await window.gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SHEET_ID,
    range,
  });
  return resp.result.values || [];
}

async function appendRows(tab, rows) {
  await window.gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId: CONFIG.SHEET_ID,
    range: `${tab}!A1`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    resource: { values: rows },
  });
}

async function updateRange(range, values) {
  await window.gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: CONFIG.SHEET_ID,
    range,
    valueInputOption: 'RAW',
    resource: { values },
  });
}

async function batchUpdate(data) {
  await window.gapi.client.sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: CONFIG.SHEET_ID,
    resource: {
      valueInputOption: 'RAW',
      data,
    },
  });
}

async function clearAndRewrite(tab, rows) {
  // Clear everything after the header, then write fresh
  await window.gapi.client.sheets.spreadsheets.values.clear({
    spreadsheetId: CONFIG.SHEET_ID,
    range: `${tab}!A2:Z`,
  });
  if (rows.length) {
    await appendRows(tab, rows);
  }
}

/* ══════════════════════════════════════════════════════════
   DECK OPERATIONS
══════════════════════════════════════════════════════════ */

/**
 * Fetch all decks from the sheet.
 * Returns: [{ name, icon }, ...]
 */
export async function fetchDecks() {
  const rows = await getRange(`${CONFIG.DECKS_TAB}!A2:B`);
  return rows
    .filter(r => r[0]?.trim())
    .map(r => ({ name: r[0], icon: r[1] || '📖' }));
}

/**
 * Add a new deck row.
 */
export async function addDeck(name, icon = '📖') {
  await appendRows(CONFIG.DECKS_TAB, [[name, icon]]);
}

/**
 * Delete a deck by name (removes its row from decks tab + all its cards).
 */
export async function deleteDeck(name) {
  // Rewrite decks tab without this deck
  const rows = await getRange(`${CONFIG.DECKS_TAB}!A2:B`);
  const remaining = rows.filter(r => r[0] !== name);
  await clearAndRewrite(CONFIG.DECKS_TAB, remaining);

  // Rewrite cards tab without this deck's cards
  const cardRows = await getRange(`${CONFIG.CARDS_TAB}!A2:H`);
  const keptCards = cardRows.filter(r => r[1] !== name);
  await clearAndRewrite(CONFIG.CARDS_TAB, keptCards);
}

/* ══════════════════════════════════════════════════════════
   CARD OPERATIONS
══════════════════════════════════════════════════════════ */

/**
 * Fetch all cards, optionally filtered by deck.
 * Returns: [{ id, deck, front, back, reps, interval, ef, due }, ...]
 */
export async function fetchCards(deckName = null) {
  const rows = await getRange(`${CONFIG.CARDS_TAB}!A2:H`);
  return rows
    .filter(r => r[0] && (!deckName || r[1] === deckName))
    .map(rowToCard);
}

function rowToCard(r) {
  return {
    id:       r[0] || '',
    deck:     r[1] || '',
    front:    r[2] || '',
    back:     r[3] || '',
    reps:     parseInt(r[4]) || 0,
    interval: parseInt(r[5]) || 1,
    ef:       parseFloat(r[6]) || 2.5,
    due:      r[7] || today(),
  };
}

function cardToRow(c) {
  return [c.id, c.deck, c.front, c.back, c.reps, c.interval, c.ef, c.due];
}

function today() { return new Date().toISOString().slice(0, 10); }

/**
 * Append new cards to the sheet.
 */
export async function addCards(cards) {
  const rows = cards.map(cardToRow);
  await appendRows(CONFIG.CARDS_TAB, rows);
}

/**
 * Batch-update SM-2 fields for a list of cards after a study session.
 * Only writes the columns that change: reps, interval, ef, due (cols E–H).
 */
export async function syncStudyResults(updatedCards) {
  if (!updatedCards.length) return;

  // Fetch all current rows to find line numbers
  const allRows = await getRange(`${CONFIG.CARDS_TAB}!A2:H`);
  const idToRowIndex = {};
  allRows.forEach((r, i) => { if (r[0]) idToRowIndex[r[0]] = i + 2; }); // +2: 1-indexed + header

  const updates = updatedCards
    .filter(c => idToRowIndex[c.id])
    .map(c => ({
      range: `${CONFIG.CARDS_TAB}!E${idToRowIndex[c.id]}:H${idToRowIndex[c.id]}`,
      values: [[c.reps, c.interval, c.ef, c.due]],
    }));

  if (updates.length) await batchUpdate(updates);
}

/**
 * Delete a single card by ID.
 */
export async function deleteCard(cardId) {
  const allRows = await getRange(`${CONFIG.CARDS_TAB}!A2:H`);
  const remaining = allRows.filter(r => r[0] !== String(cardId));
  await clearAndRewrite(CONFIG.CARDS_TAB, remaining);
}

/* ══════════════════════════════════════════════════════════
   FULL SYNC — load everything at once
══════════════════════════════════════════════════════════ */

/**
 * Load all decks and cards in two parallel requests.
 * Returns the full app state object: { decks: { name: { icon, cards[] } } }
 */
export async function loadAll() {
  const [deckRows, cardRows] = await Promise.all([
    getRange(`${CONFIG.DECKS_TAB}!A2:B`),
    getRange(`${CONFIG.CARDS_TAB}!A2:H`),
  ]);

  const decks = {};
  deckRows
    .filter(r => r[0]?.trim())
    .forEach(r => { decks[r[0]] = { icon: r[1] || '📖', cards: [] }; });

  cardRows
    .filter(r => r[0] && r[1] && decks[r[1]])
    .forEach(r => decks[r[1]].cards.push(rowToCard(r)));

  return { decks };
}

/**
 * Ensure the sheet has the correct header rows.
 * Safe to call on first run — won't overwrite existing headers.
 */
export async function ensureHeaders() {
  const decksHeader = await getRange(`${CONFIG.DECKS_TAB}!A1:B1`);
  if (!decksHeader.length || decksHeader[0][0] !== 'name') {
    await updateRange(`${CONFIG.DECKS_TAB}!A1:B1`, [['name', 'icon']]);
  }
  const cardsHeader = await getRange(`${CONFIG.CARDS_TAB}!A1:H1`);
  if (!cardsHeader.length || cardsHeader[0][0] !== 'id') {
    await updateRange(`${CONFIG.CARDS_TAB}!A1:H1`, [['id', 'deck', 'front', 'back', 'reps', 'interval', 'ef', 'due']]);
  }
}
