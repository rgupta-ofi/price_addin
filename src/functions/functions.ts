/**
 * Excel Custom Functions — Live market signal formulas.
 *
 * Usage:
 *   =INFINITY.LIVEPRICE("spot", "RATE")
 *   =INFINITY.LIVEPRICE("1m", "RATE")
 *   =INFINITY.FIELDS("spot")
 *
 * Polling starts automatically when the first formula is entered
 * and stops when the last formula is removed.
 */

import { liveDataService, LiveDataSnapshot, TickerRecord } from "../helpers/live-data-service";
import { Config } from "../helpers/config";

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
let latestStatus = "Disconnected";

/** Accumulated latest values for every ticker we've ever seen */
const latestData: Record<string, TickerRecord> = {};

/** Reverse lookup: Security ID (e.g. "USDJPY") → ticker key (e.g. "usd-jpy") */
const secIdToTicker = new Map<string, string>();

/** Case-insensitive lookup: "CC1" → "cc1" */
const canonicalKeys = new Map<string, string>();
// Seed the special key so resolveTickerKey works
canonicalKeys.set("_ALL_TICKERS_", "_ALL_TICKERS_");


// ─── Optimised Lookup & Update ──────────────────────────────────────────────

/** 
 * Try to resolve a user-provided ticker string to a canonical key.
 * Returns undefined if not found in our known data.
 */
function resolveTickerKey(input: string): string | undefined {
  if (input === "_ALL_TICKERS_") return "_ALL_TICKERS_";

  const upper = input.toUpperCase();
  // 1. Try case-insensitive match against known keys
  // Note: We intentionally skip the special _ALL_TICKERS_ key here if user typed it manually
  if (canonicalKeys.has(upper)) {
     const k = canonicalKeys.get(upper);
     if (k !== "_ALL_TICKERS_") return k;
  }
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

  // 4. Update any special "TICKERS" function calls
  const tickerWatchers = cellsByKey.get("_ALL_TICKERS_");
  if (tickerWatchers) {
    // Only rebuild string if we have new keys in snapshot (implies possible new tickers)
    // Or just update anyway since it's infrequent relative to price ticks
    const allKeys = Array.from(canonicalKeys.values())
        .filter(k => k !== "_ALL_TICKERS_")
        .sort().join(", ");
    for (const handler of tickerWatchers) {
      handler.setResult(allKeys);
    }
  }
}

// ─── Service lifecycle ──────────────────────────────────────────────────────

function ensureService(): void {
  if (!listenerAttached) {
    liveDataService.subscribe(onSnapshot);
    liveDataService.addStatusListener((status) => {
        latestStatus = status;
        // Broadcast status to all pending cells
        for (const handler of pendingCells) {
            handler.setResult(`Waiting... (${status})`);
        }

        const tickerWatchers = cellsByKey.get("_ALL_TICKERS_");
        if (tickerWatchers && !hasLiveTickers()) {
          for (const handler of tickerWatchers) {
            handler.setResult(`Waiting... (${status})`);
          }
        }
    });

    listenerAttached = true;
  }
  if (!connected) {
    liveDataService.acquire();
    connected = true;
  }
}

function hasLiveTickers(): boolean {
  return Object.keys(latestData).length > 0;
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
    if (key === "_ALL_TICKERS_") {
        const allKeys = Array.from(canonicalKeys.values())
             .filter(k => k !== "_ALL_TICKERS_")
             .sort().join(", ");
      handler.setResult(allKeys || `Waiting... (${latestStatus})`);
    } else {
        const rec = latestData[key];
        if (rec) {
          pushValueToHandler(handler, key);
        } else {
           // We found the key but have no data yet (? should not happen if key is resolved, unless structure is partial)
           invocation.setResult("Waiting... (Data pending)");
        }
    }
  } else {
    // Key not known yet, add to pending and show current WS info
    pendingCells.add(handler);
    invocation.setResult(`Waiting... (${liveDataService.getStatus()})`);
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
 * @description Returns a live FX market signal field for the given tenor.
 * @param {string} ticker The FX tenor (e.g. "spot", "1m", "2m", "3m").
 * @param {string} field The field to return: RATE, MID, METHOD, AS_OF.
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
 * @param {string} ticker The FX tenor (e.g. "spot", "1m", "2m", "3m").
 * @param {CustomFunctions.StreamingInvocation<string>} invocation
 */
function getFields(
  ticker: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  // Use a special internal field name to signal "all keys" request
  // Cast invocation to any to allow passing to shared startStreaming which handles string|number
  startStreaming(ticker, "_ALL_KEYS_", invocation as any);
}

/**
 * @customfunction TICKERS
 * @streaming
 * @description Returns a list of all live FX tenors currently returned by Infinity.
 * @param {CustomFunctions.StreamingInvocation<string>} invocation
 */
function getTickers(
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  // Use a special invalid ticker name and field to signal "all tickers" request
  // This will hook into the pushValueToHandler special check we will add
  // We use "_ALL_TICKERS_" as the ticker name, registered in canonicalKeys
  startStreaming("_ALL_TICKERS_", "_ALL_TICKERS_", invocation as any);
}

function cleanToken(token: string): string {
  return token
    .trim()
    .replace(/^["']|["']$/g, "")
    .replace(/^Bearer\s+/i, "")
    .replace(/\s+/g, "");
}

function describeToken(token: string): string {
  const jwtSegmentCount = token.split(".").length;
  const invalidJwtChars = token.match(/[^A-Za-z0-9._-]/g)?.length || 0;
  return `Token chars: ${token.length}. JWT segments: ${jwtSegmentCount}. Invalid JWT chars: ${invalidJwtChars}.`;
}

async function fetchStatus(url: string, token?: string): Promise<string> {
  try {
    const response = await fetch(url, {
      method: "GET",
      headers: token ? { Authorization: `Bearer ${token}` } : undefined,
      credentials: "omit",
      cache: "no-store",
      mode: "cors",
    });
    const text = await response.text();
    return `HTTP ${response.status}: ${text.slice(0, 120)}`;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return `failed: ${message}`;
  }
}

/**
 * @customfunction SETTOKEN
 * @description Saves the Infinity API bearer token in this Excel add-in runtime.
 * @param {string} token The bearer token value without the leading "Bearer ".
 * @returns A confirmation message.
 */
function setToken(token: string): string {
  const cleaned = cleanToken(token);
  if (!cleaned) return "Token was empty";
  localStorage.setItem(Config.authTokenStorageKey, cleaned);
  return `Infinity token saved (${cleaned.length} chars). Recalculate LIVEPRICE formulas.`;
}

/**
 * @customfunction DIAGNOSTICS
 * @description Checks the saved token and Infinity live FX API from this Excel runtime.
 * @returns A diagnostic status message.
 */
async function diagnostics(): Promise<string> {
  const rawToken = typeof localStorage === "undefined"
    ? ""
    : localStorage.getItem(Config.authTokenStorageKey) || "";

  const token = cleanToken(rawToken);
  if (!token) return "No token saved. Run INFINITY.SETTOKEN first.";

  const tokenDetails = describeToken(token);
  const url = `${Config.serverUrl}${Config.liveFxPath}`;
  const origin = typeof location === "undefined" ? "unknown origin" : location.origin;

  const anonymousStatus = await fetchStatus(url);
  const authorizedStatus = await fetchStatus(url, token);
  const runtime = typeof navigator === "undefined" ? "unknown runtime" : navigator.userAgent.slice(0, 100);

  return `${tokenDetails} Origin: ${origin}. No-auth: ${anonymousStatus}. Auth: ${authorizedStatus}. Runtime: ${runtime}`;
}

// Register
CustomFunctions.associate("LIVEPRICE", livePrice);
CustomFunctions.associate("FIELDS", getFields);
CustomFunctions.associate("TICKERS", getTickers);
CustomFunctions.associate("SETTOKEN", setToken);
CustomFunctions.associate("DIAGNOSTICS", diagnostics);

// Initialize Office.js
Office.onReady(() => {
  console.log("Infinity Custom Functions loaded - v4");
});
