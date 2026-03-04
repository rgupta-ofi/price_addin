/**
 * Excel Custom Functions — Live-streaming price formulas.
 *
 * Usage:
 *   =INFINITY.LIVEPRICE("cc1", "MID")
 *   =INFINITY.LIVEPRICE("usd-jpy", "BID")
 *   =INFINITY.LIVEPRICE("USDJPY", "ASK")     ← also matches by Security ID
 *
 * The WebSocket connects automatically when the first formula is entered
 * and disconnects when the last formula is removed. No sign-in required.
 */

import { liveDataService, LiveDataSnapshot, TickerRecord } from "../helpers/live-data-service";

// ─── State ──────────────────────────────────────────────────────────────────

interface CellHandler {
  setResult: (v: string | number) => void;
  ticker: string;
  field: string;
}

/** All active streaming cells, keyed by a unique random ID (for cancellation) */
const cells = new Map<string, CellHandler>();

/** Cells grouped by their resolved canonical ticker key (e.g. "cc1") */
const cellsByKey = new Map<string, Set<CellHandler>>();

/** Cells that haven't been resolved to a canonical key yet */
const pendingCells = new Set<CellHandler>();

let connected = false;
let listenerAttached = false;

/** Accumulated latest values for every ticker we've ever seen */
const latestData: Record<string, TickerRecord> = {};

/** Reverse lookup: Security ID (e.g. "USDJPY") → ticker key (e.g. "usd-jpy") */
const secIdToTicker = new Map<string, string>();

/** Case-insensitive lookup: "CC1" → "cc1" */
const canonicalKeys = new Map<string, string>();


// ─── Optimised Lookup & Update ──────────────────────────────────────────────

/** 
 * Try to resolve a user-provided ticker string to a canonical key.
 * Returns undefined if not found in our known data.
 */
function resolveTickerKey(input: string): string | undefined {
  const upper = input.toUpperCase();
  // 1. Try case-insensitive match against known keys
  if (canonicalKeys.has(upper)) return canonicalKeys.get(upper);
  // 2. Try Security ID match
  if (secIdToTicker.has(upper)) return secIdToTicker.get(upper);
  return undefined;
}

/** 
 * Register a cell. Tries to resolve its key immediately. 
 * If successful, adds to cellsByKey. If not, adds to pendingCells.
 */
function registerCell(handler: CellHandler) {
  const key = resolveTickerKey(handler.ticker);
  if (key) {
    let set = cellsByKey.get(key);
    if (!set) {
      set = new Set();
      cellsByKey.set(key, set);
    }
    set.add(handler);
    // Push immediate data if available
    pushValueToHandler(handler, key);
  } else {
    pendingCells.add(handler);
    handler.setResult("Waiting...");
  }
}

/** Update a specific handler with data for a known key */
function pushValueToHandler(handler: CellHandler, key: string) {
  const rec = latestData[key];
  if (!rec) return;

  if (handler.field === "_ALL_KEYS_") {
    if (rec.result) {
      // Sort keys for consistent display
      const keys = Object.keys(rec.result).sort().join(", ");
      handler.setResult(keys);
    }
    return;
  }

  const fieldUpper = handler.field.toUpperCase();
  // Try exact casing first, then uppercase
  const val = (rec.result?.[handler.field] ?? rec.result?.[fieldUpper]) as string | number | undefined;

  if (val !== undefined && val !== null) {
    handler.setResult(val);
  }
}

// ─── Snapshot Listener ──────────────────────────────────────────────────────

function onSnapshot(snapshot: LiveDataSnapshot): void {
  // 1. Process new data structure (discover new keys/IDs)
  for (const [key, rec] of Object.entries(snapshot)) {
    latestData[key] = rec;
    canonicalKeys.set(key.toUpperCase(), key);
    
    // Check for ID_BB_SEC_NUMBER_DESCRIPTION_RT
    const secId = rec.result?.ID_BB_SEC_NUMBER_DESCRIPTION_RT;
    if (secId && typeof secId === 'string') {
      secIdToTicker.set(secId.toUpperCase(), key);
    }
  }

  // 2. Process pending cells (maybe we can resolve them now?)
  if (pendingCells.size > 0) {
    for (const handler of pendingCells) {
      const key = resolveTickerKey(handler.ticker);
      if (key) {
        pendingCells.delete(handler);
        let set = cellsByKey.get(key);
        if (!set) {
          set = new Set();
          cellsByKey.set(key, set);
        }
        set.add(handler);
        // We'll update it in step 3 if it's in the snapshot, 
        // OR we should update it now from latestData just in case
        pushValueToHandler(handler, key);
      }
    }
  }

  // 3. Efficient Update: Only update cells for keys that are IN THIS SNAPSHOT
  for (const key of Object.keys(snapshot)) {
    const set = cellsByKey.get(key);
    if (set) {
      for (const handler of set) {
        pushValueToHandler(handler, key);
      }
    }
  }
}

// ─── Service lifecycle ──────────────────────────────────────────────────────

function ensureService(): void {
  if (!listenerAttached) {
    liveDataService.subscribe(onSnapshot);
    listenerAttached = true;
  }
  if (!connected) {
    liveDataService.acquire();
    connected = true;
  }
}

// ─── Custom Functions ───────────────────────────────────────────────────────

function startStreaming(
  ticker: string,
  field: string,
  invocation: CustomFunctions.StreamingInvocation<string | number>
): void {
  const id = Math.random().toString(36).slice(2, 11);

  ensureService();

  const handler: CellHandler = { setResult: invocation.setResult, ticker, field };
  cells.set(id, handler);

  // Try to resolve and register immediately
  const key = resolveTickerKey(ticker);
  
  if (key) {
    let set = cellsByKey.get(key);
    if (!set) {
      set = new Set();
      cellsByKey.set(key, set);
    }
    set.add(handler);
    
    // Provide immediate value if available
    const rec = latestData[key];
    if (rec) {
      pushValueToHandler(handler, key);
    } else {
       invocation.setResult("Waiting...");
    }
  } else {
    // Key not known yet, add to pending
    pendingCells.add(handler);
    invocation.setResult("Waiting...");
  }

  invocation.onCanceled = () => {
    cells.delete(id);
    pendingCells.delete(handler);
    
    // Remove from canonical map (scan all keys)
    for (const set of cellsByKey.values()) {
      if (set.delete(handler)) break;
    }

    if (cells.size === 0 && connected) {
      liveDataService.release();
      connected = false;
    }
  };
}

// ─── Custom Functions ───────────────────────────────────────────────────────

/**
 * @customfunction LIVEPRICE
 * @streaming
 * @description Returns a live-streaming price field for the given ticker.
 * @param {string} ticker The ticker key (e.g. "cc1", "usd-jpy") or security ID (e.g. "CCH6", "USDJPY").
 * @param {string} field The field to return: MID, BID, ASK, LAST_PRICE, OPEN, HIGH, LOW, VOLUME.
 * @param {CustomFunctions.StreamingInvocation<string | number>} invocation
 */
function livePrice(
  ticker: string,
  field: string,
  invocation: CustomFunctions.StreamingInvocation<string | number>
): void {
  startStreaming(ticker, field, invocation);
}

/**
 * @customfunction FIELDS
 * @streaming
 * @description Returns a list of all available data fields for a ticker.
 * @param {string} ticker The ticker key (e.g. "cc1", "usd-jpy") or security ID (e.g. "CCH6", "USDJPY").
 * @param {CustomFunctions.StreamingInvocation<string>} invocation
 */
function getFields(
  ticker: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  // Use a special internal field name to signal "all keys" request
  startStreaming(ticker, "_ALL_KEYS_", invocation);
}

// Register
CustomFunctions.associate("LIVEPRICE", livePrice);
CustomFunctions.associate("FIELDS", getFields);

// Initialize Office.js
Office.onReady(() => {
  console.log("Infinity Custom Functions loaded");
});
