/**
 * graph.js — Microsoft Graph API layer (replaces sheets.js)
 *
 * ─── SETUP (one-time, ~10 minutes) ────────────────────────────────────────
 *
 * 1. Go to https://portal.azure.com → Azure Active Directory → App registrations
 *    → New registration
 *      Name:     Flashcards PWA
 *      Account type: "Accounts in any organizational directory and personal Microsoft accounts"
 *      Redirect URI: Single-page application (SPA) → https://yourusername.github.io/flashcards
 *    → Register
 *
 * 2. Copy the "Application (client) ID" and paste it into CONFIG.CLIENT_ID below.
 *    That's the only value you need. No secrets, no API keys.
 *
 * 3. Under "API permissions" → Add a permission → Microsoft Graph → Delegated:
 *      Files.ReadWrite   (read/write the Excel file)
 *      User.Read         (show the signed-in user's name)
 *    → Grant admin consent if prompted (only needed for org accounts)
 *
 * 4. Create an Excel file in your OneDrive:
 *    - Name it exactly:  flashcards.xlsx
 *    - Add two sheets (tabs) named exactly:  decks   and   cards
 *    - The app will write headers automatically on first run.
 *    - Default path assumed: /Documents/flashcards.xlsx
 *      Change CONFIG.FILE_PATH below if you put it somewhere else.
 *
 * 5. Deploy the app to the redirect URI you registered in step 1.
 *    Open it in a browser, tap "Sign in with Microsoft", and you're live.
 *
 * ─── NO API KEY NEEDED ────────────────────────────────────────────────────
 *  MSAL (Microsoft Authentication Library) uses OAuth 2.0 PKCE — the client
 *  ID is a public identifier, not a secret. It's safe in source code.
 * ──────────────────────────────────────────────────────────────────────────
 */

export const CONFIG = {
  // Paste your Azure App Registration "Application (client) ID" here:
  CLIENT_ID: 'cfc61d4b-e2a7-46cd-9911-681b89733c5d',

  // OAuth authority — works for both personal (outlook.com) and work/school accounts:
  AUTHORITY: 'https://login.microsoftonline.com/common',

  // Scopes requested during sign-in:
  SCOPES: ['Files.ReadWrite', 'User.Read'],

  // Path to your Excel file in OneDrive (relative to the user's OneDrive root).
  // Default: OneDrive root → Documents → flashcards.xlsx
  FILE_PATH: '/Documents/flashcards.xlsx',

  // Sheet (tab) names inside the Excel file:
  DECKS_SHEET: 'decks',
  CARDS_SHEET: 'cards',
};

// ── Microsoft Graph base URL ──────────────────────────────────────────────
const GRAPH = 'https://graph.microsoft.com/v1.0';

// ── MSAL instance (initialised in initGraph) ──────────────────────────────
let msalInstance = null;
let currentAccount = null;

/* ══════════════════════════════════════════════════════════════════════════
   INIT — call once on app start
══════════════════════════════════════════════════════════════════════════ */
export async function initGraph() {
  await loadMsal();

  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:   CONFIG.CLIENT_ID,
      authority:  CONFIG.AUTHORITY,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: 'sessionStorage', // keeps token in session; clears on tab close
      storeAuthStateInCookie: false,
    },
  });

  await msalInstance.initialize();

  // Handle redirect response (called after the OAuth redirect returns)
  const response = await msalInstance.handleRedirectPromise();
  if (response) {
    currentAccount = response.account;
    msalInstance.setActiveAccount(currentAccount);
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length) {
      currentAccount = accounts[0];
      msalInstance.setActiveAccount(currentAccount);
    }
  }
}

function loadMsal() {
  return new Promise((resolve, reject) => {
    if (window.msal) { resolve(); return; }
    const s = document.createElement('script');
    s.src = './msal-browser.min.js';
    s.onload = resolve;
    s.onerror = () => reject(new Error('Failed to load MSAL'));
    document.head.appendChild(s);
  });
}

/* ══════════════════════════════════════════════════════════════════════════
   AUTH
══════════════════════════════════════════════════════════════════════════ */
export function isSignedIn() {
  return !!currentAccount;
}

export function getUser() {
  return currentAccount ? {
    name:  currentAccount.name,
    email: currentAccount.username,
  } : null;
}

export async function signIn() {
  // Try silent first (token already cached), fall back to popup
  try {
    const result = await msalInstance.ssoSilent({ scopes: CONFIG.SCOPES });
    currentAccount = result.account;
    msalInstance.setActiveAccount(currentAccount);
  } catch {
    const result = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    currentAccount = result.account;
    msalInstance.setActiveAccount(currentAccount);
  }
}

export async function signOut() {
  await msalInstance.logoutPopup({ account: currentAccount });
  currentAccount = null;
}

/* ── Get a fresh access token (silently refreshes when expired) ── */
async function getToken() {
  if (!currentAccount) throw new Error('Not signed in');
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.SCOPES,
      account: currentAccount,
    });
    return result.accessToken;
  } catch {
    // Silent refresh failed — prompt user
    const result = await msalInstance.acquireTokenPopup({ scopes: CONFIG.SCOPES });
    return result.accessToken;
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   LOW-LEVEL GRAPH HELPERS
══════════════════════════════════════════════════════════════════════════ */
async function graphFetch(path, options = {}) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`Graph API error ${res.status}: ${err?.error?.message || res.statusText}`);
  }
  if (res.status === 204) return null; // No content
  return res.json();
}

// Drive item path for the Excel file
const filePath = () =>
  `/me/drive/root:${CONFIG.FILE_PATH}:/workbook`;

/* ── Excel table/range helpers ── */

async function getUsedRange(sheet) {
  const data = await graphFetch(
    `${filePath()}/worksheets/${sheet}/usedRange`
  );
  return data?.values || [];
}

async function updateRange(sheet, address, values) {
  await graphFetch(
    `${filePath()}/worksheets/${sheet}/range(address='${address}')`,
    {
      method: 'PATCH',
      body: JSON.stringify({ values }),
    }
  );
}

async function appendRows(sheet, rows) {
  // Read current used range to find next empty row
  const existing = await getUsedRange(sheet);
  const nextRow = existing.length + 1; // 1-indexed, header is row 1
  const startCell = `A${nextRow}`;
  const endCol = String.fromCharCode(64 + rows[0].length); // A=65
  const endCell = `${endCol}${nextRow + rows.length - 1}`;
  await updateRange(sheet, `${startCell}:${endCell}`, rows);
}

async function clearDataRows(sheet) {
  // Clear everything from row 2 downward (keep headers)
  const used = await getUsedRange(sheet);
  if (used.length <= 1) return; // Only header or empty
  const lastRow = used.length;
  const lastCol = String.fromCharCode(64 + used[0].length);
  await updateRange(sheet, `A2:${lastCol}${lastRow}`,
    Array(lastRow - 1).fill(Array(used[0].length).fill('')));
}

/* ══════════════════════════════════════════════════════════════════════════
   ENSURE FILE & HEADERS EXIST
══════════════════════════════════════════════════════════════════════════ */
export async function ensureWorkbook() {
  // Check if file exists; if not, create it via Graph
  try {
    await graphFetch(`/me/drive/root:${CONFIG.FILE_PATH}:`);
  } catch {
    // File doesn't exist — create empty xlsx
    const token = await getToken();
    await fetch(`${GRAPH}/me/drive/root:${CONFIG.FILE_PATH}:/content`, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      // Minimal valid xlsx bytes (empty workbook)
      body: new Blob([]),
    });
  }

  // Ensure sheets exist
  try {
    await graphFetch(`${filePath()}/worksheets/${CONFIG.DECKS_SHEET}`);
  } catch {
    await graphFetch(`${filePath()}/worksheets`, {
      method: 'POST',
      body: JSON.stringify({ name: CONFIG.DECKS_SHEET }),
    });
  }
  try {
    await graphFetch(`${filePath()}/worksheets/${CONFIG.CARDS_SHEET}`);
  } catch {
    await graphFetch(`${filePath()}/worksheets`, {
      method: 'POST',
      body: JSON.stringify({ name: CONFIG.CARDS_SHEET }),
    });
  }

  // Ensure headers
  const decksRows = await getUsedRange(CONFIG.DECKS_SHEET);
  if (!decksRows.length || decksRows[0][0] !== 'name') {
    await updateRange(CONFIG.DECKS_SHEET, 'A1:B1', [['name', 'icon']]);
  }
  const cardsRows = await getUsedRange(CONFIG.CARDS_SHEET);
  if (!cardsRows.length || cardsRows[0][0] !== 'id') {
    await updateRange(CONFIG.CARDS_SHEET, 'A1:H1',
      [['id', 'deck', 'front', 'back', 'reps', 'interval', 'ef', 'due']]);
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   FULL LOAD
══════════════════════════════════════════════════════════════════════════ */

/**
 * Load all decks and cards in two parallel requests.
 * Returns: { decks: { name: { icon, cards[] } } }
 */
export async function loadAll() {
  const [deckRows, cardRows] = await Promise.all([
    getUsedRange(CONFIG.DECKS_SHEET),
    getUsedRange(CONFIG.CARDS_SHEET),
  ]);

  const decks = {};

  // Skip header row (index 0)
  deckRows.slice(1)
    .filter(r => r[0]?.toString().trim())
    .forEach(r => {
      decks[r[0]] = { icon: r[1] || '📖', cards: [] };
    });

  cardRows.slice(1)
    .filter(r => r[0] && r[1] && decks[r[1]])
    .forEach(r => decks[r[1]].cards.push(rowToCard(r)));

  return { decks };
}

/* ══════════════════════════════════════════════════════════════════════════
   DECK OPERATIONS
══════════════════════════════════════════════════════════════════════════ */
export async function addDeck(name, icon = '📖') {
  await appendRows(CONFIG.DECKS_SHEET, [[name, icon]]);
}

export async function deleteDeck(name) {
  const [deckRows, cardRows] = await Promise.all([
    getUsedRange(CONFIG.DECKS_SHEET),
    getUsedRange(CONFIG.CARDS_SHEET),
  ]);

  const newDecks = deckRows.slice(1).filter(r => r[0] !== name);
  const newCards = cardRows.slice(1).filter(r => r[1] !== name);

  await rewriteSheet(CONFIG.DECKS_SHEET, deckRows[0], newDecks);
  await rewriteSheet(CONFIG.CARDS_SHEET, cardRows[0], newCards);
}

/* ══════════════════════════════════════════════════════════════════════════
   CARD OPERATIONS
══════════════════════════════════════════════════════════════════════════ */
export async function addCards(cards) {
  const rows = cards.map(cardToRow);
  await appendRows(CONFIG.CARDS_SHEET, rows);
}

/**
 * Batch-write updated SM-2 fields for cards reviewed in a session.
 * Rewrites only the changed rows (reps, interval, ef, due = cols E–H).
 */
export async function syncStudyResults(updatedCards) {
  if (!updatedCards.length) return;

  const allRows = await getUsedRange(CONFIG.CARDS_SHEET);
  const header = allRows[0];
  const dataRows = allRows.slice(1);

  const updatedById = Object.fromEntries(updatedCards.map(c => [c.id, c]));
  const newRows = dataRows.map(r => {
    const updated = updatedById[r[0]];
    if (!updated) return r;
    const row = [...r];
    row[4] = updated.reps;
    row[5] = updated.interval;
    row[6] = updated.ef;
    row[7] = updated.due;
    return row;
  });

  await rewriteSheet(CONFIG.CARDS_SHEET, header, newRows);
}

export async function deleteCard(cardId) {
  const allRows = await getUsedRange(CONFIG.CARDS_SHEET);
  const header = allRows[0];
  const remaining = allRows.slice(1).filter(r => String(r[0]) !== String(cardId));
  await rewriteSheet(CONFIG.CARDS_SHEET, header, remaining);
}

/* ══════════════════════════════════════════════════════════════════════════
   INTERNAL HELPERS
══════════════════════════════════════════════════════════════════════════ */
async function rewriteSheet(sheet, header, dataRows) {
  // Clear all data rows
  await clearDataRows(sheet);
  // Re-write header (in case clear wiped it)
  await updateRange(sheet, `A1:${col(header.length)}1`, [header]);
  // Write data rows if any
  if (dataRows.length) {
    const startRow = 2;
    const endRow = startRow + dataRows.length - 1;
    const endColLetter = col(header.length);
    // Pad short rows to full width
    const padded = dataRows.map(r => {
      const row = [...r];
      while (row.length < header.length) row.push('');
      return row;
    });
    await updateRange(sheet, `A${startRow}:${endColLetter}${endRow}`, padded);
  }
}

// Convert column count to letter (1→A, 2→B, … 8→H)
function col(n) { return String.fromCharCode(64 + n); }

function today() { return new Date().toISOString().slice(0, 10); }

function rowToCard(r) {
  return {
    id:       String(r[0] || ''),
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
