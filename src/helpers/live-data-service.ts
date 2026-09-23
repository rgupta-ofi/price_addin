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
  tenor: string;
  rate: number | null;
  method?: string | null;
  as_of?: string | null;
}

interface LiveFxResponse {
  data?: LiveFxRow[];
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

    fetch(url, {
      method: "GET",
      headers: { Authorization: `Bearer ${token}` },
      signal: controller.signal,
    })
      .then(async response => {
        if (response.status === 401 || response.status === 403) {
          throw new Error("AUTH_REJECTED");
        }
        if (!response.ok) {
          throw new Error(`HTTP_${response.status}`);
        }
        return response.json() as Promise<LiveFxResponse>;
      })
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

  private getAuthToken(): string | null {
    if (typeof localStorage === "undefined") return null;
    const token = localStorage.getItem(Config.authTokenStorageKey);
    return token?.trim() || null;
  }

  private handleLiveFx(payload: LiveFxResponse): void {
    const rows = Array.isArray(payload.data) ? payload.data : [];
    const snapshot: LiveDataSnapshot = {};

    for (const row of rows) {
      if (!row.tenor) continue;
      const tenor = row.tenor.toLowerCase();
      snapshot[tenor] = {
        result: {
          ID_BB_SEC_NUMBER_DESCRIPTION_RT: row.tenor,
          RATE: row.rate,
          MID: row.rate ?? undefined,
          METHOD: row.method ?? undefined,
          AS_OF: row.as_of ?? undefined,
        },
      };
    }

    if (!Object.keys(snapshot).length) {
      this.setStatus("No live FX data returned");
      return;
    }

    this.setStatus("Receiving Data");
    for (const [key, record] of Object.entries(snapshot)) this.buffer[key] = record;
    this.bufferDirty = true;
    this.scheduleFlush();
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
