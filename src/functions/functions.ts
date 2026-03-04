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

/** All active streaming cells, keyed by a unique random ID */
const cells = new Map<string, CellHandler>();

let connected = false;
let listenerAttached = false;

/** Accumulated latest values for every ticker we've ever seen, so new formulas
 *  can immediately display the most recent value instead of waiting for the next WS push. */
const latestData: Record<string, TickerRecord> = {};

/** Reverse lookup: Security ID (e.g. "USDJPY") → ticker key (e.g. "usd-jpy") */
const secIdToTicker: Record<string, string> = {};

// ─── Snapshot listener ──────────────────────────────────────────────────────

function onSnapshot(snapshot: LiveDataSnapshot): void {
  // Accumulate latest data and build reverse lookup
  for (const [key, rec] of Object.entries(snapshot)) {
    latestData[key] = rec;
    const secId = rec.result?.ID_BB_SEC_NUMBER_DESCRIPTION_RT;
    if (secId) secIdToTicker[secId.toUpperCase()] = key;
  }

  // Push values to each active cell
  cells.forEach((handler) => {
    pushValue(handler);
  });
}

/** Resolve a user-provided ticker to its internal key */
function resolveTickerKey(input: string): string | undefined {
  const upper = input.toUpperCase();
  // Direct match by WebSocket key (e.g. "cc1", "usd-jpy")
  for (const key of Object.keys(latestData)) {
    if (key.toUpperCase() === upper) return key;
  }
  // Match by Security ID (e.g. "USDJPY", "CCH6")
  return secIdToTicker[upper];
}

/** Push a value to a single cell handler */
function pushValue(handler: CellHandler): void {
  const key = resolveTickerKey(handler.ticker);
  if (!key) return;
  const rec = latestData[key];
  if (!rec) return;

  const fieldUpper = handler.field.toUpperCase();
  let value: string | number | undefined;

  // Try exact casing first, then uppercase
  value = (rec.result?.[handler.field] ?? rec.result?.[fieldUpper]) as string | number | undefined;

  if (value !== undefined && value !== null) {
    handler.setResult(value);
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
  const id = Math.random().toString(36).slice(2, 11);

  ensureService();

  const handler: CellHandler = { setResult: invocation.setResult, ticker, field };
  cells.set(id, handler);

  // If we already have data for this ticker, show it immediately
  const existing = resolveTickerKey(ticker);
  if (existing && latestData[existing]) {
    pushValue(handler);
  } else {
    invocation.setResult("Waiting…");
  }

  invocation.onCanceled = () => {
    cells.delete(id);
    if (cells.size === 0 && connected) {
      liveDataService.release();
      connected = false;
    }
  };
}

// Register
CustomFunctions.associate("LIVEPRICE", livePrice);

// Initialize Office.js
Office.onReady(() => {
  console.log("Infinity Custom Functions loaded");
});
