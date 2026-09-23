/**
 * LiveDataService — Singleton client for streaming live price data.
 *
 * Polls the Infinity live FX endpoint, coalesces rapid updates into
 * batched snapshots, and notifies all subscribed listeners.
 *
 * Features:
 *   - Reference-counted acquire/release for clean lifecycle
 *   - Coalescing buffer to avoid flooding Excel with per-tick updates
 *   - Auth-aware HTTP polling
 *   - Network-aware (online/offline)
 */

import { Config } from "./config";

// ─── Types ───────────────────────────────────────────────────────────────────

export interface TickerResult {
  ID_BB_SEC_NUMBER_DESCRIPTION_RT?: string;
  MID?: number;
  BID?: number;
  ASK?: number;
  LAST_PRICE?: number;
  VOLUME?: number;
  OPEN?: number;
  HIGH?: number;
  LOW?: number;
  [key: string]: unknown;
}

export interface TickerRecord {
  result: TickerResult;
}

export type LiveDataSnapshot = Record<string, TickerRecord>;
export type LiveDataListener = (snapshot: LiveDataSnapshot) => void;
export type StatusListener = (status: string) => void;

interface LiveFxRow {
  tenor?: string | null;
  ticker?: string | null;
  value?: string | null;
  name?: string | null;
  pair?: string | null;
  currency_pair?: string | null;
  ccy_pair?: string | null;
  rate?: number | string | null;
  mid?: number | string | null;
  prevMid?: number | string | null;
  bid?: number | string | null;
  ask?: number | string | null;
  last?: number | string | null;
  lastPrice?: number | string | null;
  last_price?: number | string | null;
  method?: string | null;
  as_of?: string | null;
  asOf?: string | null;
  time?: string | null;
  lastChangedAt?: string | number | null;
  [key: string]: unknown;
}

type LiveFxRows = LiveFxRow[] | Record<string, LiveFxRow>;

interface LiveFxResponse {
  data?: LiveFxRows;
  [key: string]: unknown;
}

// ─── LiveDataService ────────────────────────────────────────────────────────

class LiveDataService {
  private listeners = new Set<LiveDataListener>();
  private statusListeners = new Set<StatusListener>();
  private currentStatus = "Disconnected";
  private refCount = 0;
  private pollTimer: ReturnType<typeof setTimeout> | null = null;
  private closeTimer: ReturnType<typeof setTimeout> | null = null;
  private buffer: LiveDataSnapshot = {};
  private bufferDirty = false;
  private flushTimer: ReturnType<typeof setTimeout> | null = null;
  private activeRequest: AbortController | null = null;

  constructor() {
    if (typeof window !== "undefined") {
      window.addEventListener("online", () => this.onOnline());
      window.addEventListener("offline", () => this.onOffline());
    }
  }

  // ─── Public API ──────────────────────────────────────────────────────────

  getStatus(): string {
    return this.currentStatus;
  }

  addStatusListener(fn: StatusListener): void {
    this.statusListeners.add(fn);
    fn(this.currentStatus);
  }

  removeStatusListener(fn: StatusListener): void {
    this.statusListeners.delete(fn);
  }

  private setStatus(status: string): void {
    if (this.currentStatus === status) return;
    this.currentStatus = status;
    this.statusListeners.forEach(fn => fn(status));
  }

  acquire(): void {
    this.clearCloseTimer();
    this.refCount += 1;
    this.ensureConnected();
  }

  release(): void {
    this.refCount = Math.max(0, this.refCount - 1);
    if (this.refCount === 0) this.scheduleClose();
  }

  subscribe(fn: LiveDataListener): void {
    this.listeners.add(fn);
  }

  unsubscribe(fn: LiveDataListener): void {
    this.listeners.delete(fn);
  }

  // ─── Connection ──────────────────────────────────────────────────────────

  private ensureConnected(): void {
    if (this.refCount <= 0 || this.pollTimer || this.activeRequest) return;
    if (typeof navigator !== "undefined" && !navigator.onLine) return;
    this.pollLiveFx();
  }

  private pollLiveFx(): void {
    if (this.refCount <= 0) return;

    const token = this.getAuthToken();
    if (!token) {
      this.setStatus("Authentication required: set INFINITY_API_TOKEN");
      this.scheduleNextPoll();
      return;
    }

    const url = `${Config.serverUrl}${Config.liveFxPath}`;
    const controller = new AbortController();
    this.activeRequest = controller;
    this.setStatus("Fetching live FX data");

    this.fetchLiveFxJson(url, token, controller.signal)
      .then(payload => this.handleLiveFx(payload))
      .catch(error => {
        if (controller.signal.aborted) return;
        const message = error instanceof Error ? error.message : String(error);
        this.setStatus(
          message === "AUTH_REJECTED"
            ? "Authentication expired or rejected"
            : `Live FX fetch failed: ${message}`
        );
      })
      .finally(() => {
        if (this.activeRequest === controller) this.activeRequest = null;
        this.scheduleNextPoll();
      });
  }

  private async fetchLiveFxJson(url: string, token: string, signal: AbortSignal): Promise<LiveFxResponse> {
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` },
        signal,
      });
      return await this.parseLiveFxResponse(response);
    } catch (error) {
      if (signal.aborted) throw error;
      return this.fetchLiveFxJsonWithXhr(url, token, signal);
    }
  }

  private async parseLiveFxResponse(response: Response): Promise<LiveFxResponse> {
    if (response.status === 401 || response.status === 403) {
      throw new Error("AUTH_REJECTED");
    }
    if (!response.ok) {
      throw new Error(`HTTP_${response.status}`);
    }
    return response.json() as Promise<LiveFxResponse>;
  }

  private fetchLiveFxJsonWithXhr(url: string, token: string, signal: AbortSignal): Promise<LiveFxResponse> {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      const abort = () => {
        xhr.abort();
        reject(new Error("ABORTED"));
      };

      if (signal.aborted) {
        reject(new Error("ABORTED"));
        return;
      }

      signal.addEventListener("abort", abort, { once: true });
      xhr.open("GET", url, true);
      xhr.setRequestHeader("Authorization", `Bearer ${token}`);
      xhr.onload = () => {
        signal.removeEventListener("abort", abort);
        if (xhr.status === 401 || xhr.status === 403) {
          reject(new Error("AUTH_REJECTED"));
          return;
        }
        if (xhr.status < 200 || xhr.status >= 300) {
          reject(new Error(`HTTP_${xhr.status}`));
          return;
        }
        try {
          resolve(JSON.parse(xhr.responseText) as LiveFxResponse);
        } catch {
          reject(new Error("INVALID_JSON"));
        }
      };
      xhr.onerror = () => {
        signal.removeEventListener("abort", abort);
        reject(new Error("XHR_NETWORK_ERROR"));
      };
      xhr.ontimeout = () => {
        signal.removeEventListener("abort", abort);
        reject(new Error("XHR_TIMEOUT"));
      };
      xhr.timeout = 15_000;
      xhr.send();
    });
  }

  private getAuthToken(): string | null {
    if (typeof localStorage === "undefined") return null;
    const token = localStorage.getItem(Config.authTokenStorageKey);
    const cleaned = token
      ?.trim()
      .replace(/^["']|["']$/g, "")
      .replace(/^Bearer\s+/i, "")
      .replace(/\s+/g, "");
    return cleaned || null;
  }

  private handleLiveFx(payload: LiveFxResponse): void {
    const rows = this.getLiveFxRows(payload);
    const snapshot: LiveDataSnapshot = {};

    for (const row of rows) {
      const ticker = this.getTickerKey(row);
      if (!ticker) continue;

      const rate = this.toNumber(row.rate ?? row.mid);
      const mid = this.toNumber(row.mid ?? row.rate);
      const bid = this.toNumber(row.bid);
      const ask = this.toNumber(row.ask);
      const last = this.toNumber(row.last ?? row.lastPrice ?? row.last_price);

      snapshot[ticker.toLowerCase()] = {
        result: {
          ID_BB_SEC_NUMBER_DESCRIPTION_RT: ticker,
          RATE: row.rate,
          MID: mid ?? undefined,
          BID: bid ?? undefined,
          ASK: ask ?? undefined,
          LAST_PRICE: last ?? undefined,
          PREV_MID: this.toNumber(row.prevMid) ?? undefined,
          METHOD: row.method ?? undefined,
          AS_OF: row.as_of ?? row.asOf ?? row.time ?? row.lastChangedAt ?? undefined,
        },
      };

      if (rate !== null) snapshot[ticker.toLowerCase()].result.RATE = rate;
    }

    if (!Object.keys(snapshot).length) {
      this.setStatus("No recognized live FX data returned");
      return;
    }

    this.setStatus("Receiving Data");
    for (const [key, record] of Object.entries(snapshot)) this.buffer[key] = record;
    this.bufferDirty = true;
    this.scheduleFlush();
  }

  private getLiveFxRows(payload: LiveFxResponse | LiveFxRow[] | Record<string, LiveFxRow>): LiveFxRow[] {
    const source = Array.isArray(payload) ? payload : payload.data ?? payload;
    if (Array.isArray(source)) return source;
    if (!source || typeof source !== "object") return [];

    return Object.entries(source).map(([key, value]) => ({ ticker: key, ...(value || {}) }));
  }

  private getTickerKey(row: LiveFxRow): string | null {
    const raw =
      row.tenor ??
      row.ticker ??
      row.value ??
      row.pair ??
      row.currency_pair ??
      row.ccy_pair ??
      row.name;

    if (typeof raw !== "string") return null;
    const normalized = raw.trim().replace("/", "-").toLowerCase();
    return normalized || null;
  }

  private toNumber(value: unknown): number | null {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value !== "string") return null;

    const parsed = Number(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }

  // ─── Buffer ──────────────────────────────────────────────────────────────

  private scheduleFlush(): void {
    if (this.flushTimer) return;
    this.flushTimer = setTimeout(() => {
      this.flushTimer = null;
      if (!this.bufferDirty) return;
      this.bufferDirty = false;
      const snapshot = this.buffer;
      this.buffer = {};
      this.listeners.forEach((fn) => { try { fn(snapshot); } catch (e) { console.error("[LiveData] listener error", e); } });
    }, Config.bufferFlushMs);
  }

  // ─── Polling ─────────────────────────────────────────────────────────────

  private scheduleNextPoll(): void {
    if (this.refCount <= 0 || this.pollTimer) return;
    this.pollTimer = setTimeout(() => {
      this.pollTimer = null;
      this.ensureConnected();
    }, Config.pollIntervalMs);
  }

  // ─── Graceful close ──────────────────────────────────────────────────────

  private scheduleClose(): void {
    if (this.closeTimer) return;
    this.closeTimer = setTimeout(() => {
      this.closeTimer = null;
      if (this.refCount > 0) return;
      if (this.pollTimer) { clearTimeout(this.pollTimer); this.pollTimer = null; }
      if (this.activeRequest) { this.activeRequest.abort(); this.activeRequest = null; }
      if (this.flushTimer) { clearTimeout(this.flushTimer); this.flushTimer = null; }
      this.buffer = {};
      this.bufferDirty = false;
      console.log("[LiveData] Closed (no consumers)");
    }, Config.closeDelayMs);
  }

  private clearCloseTimer(): void {
    if (this.closeTimer) { clearTimeout(this.closeTimer); this.closeTimer = null; }
  }

  // ─── Network ─────────────────────────────────────────────────────────────

  private onOnline(): void {
    if (this.refCount <= 0) return;
    if (this.pollTimer) { clearTimeout(this.pollTimer); this.pollTimer = null; }
    this.ensureConnected();
  }

  private onOffline(): void {
    if (this.pollTimer) { clearTimeout(this.pollTimer); this.pollTimer = null; }
    if (this.closeTimer) { clearTimeout(this.closeTimer); this.closeTimer = null; }
    if (this.activeRequest) { this.activeRequest.abort(); this.activeRequest = null; }
    this.setStatus("Offline");
  }
}

/** Singleton */
export const liveDataService = new LiveDataService();
