/**
 * LiveDataService — Singleton WebSocket client for streaming live price data.
 *
 * Connects to the server's WebSocket endpoint, coalesces rapid updates into
 * batched snapshots, and notifies all subscribed listeners.
 *
 * Features:
 *   - Reference-counted acquire/release for clean lifecycle
 *   - Coalescing buffer to avoid flooding Excel with per-tick updates
 *   - Exponential back-off reconnection
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

// ─── LiveDataService ────────────────────────────────────────────────────────

class LiveDataService {
  private ws: WebSocket | null = null;
  private listeners = new Set<LiveDataListener>();
  private statusListeners = new Set<StatusListener>();
  private currentStatus = "Disconnected";
  private refCount = 0;
  private retryCount = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private closeTimer: ReturnType<typeof setTimeout> | null = null;
  private buffer: LiveDataSnapshot = {};
  private bufferDirty = false;
  private flushTimer: ReturnType<typeof setTimeout> | null = null;

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
    if (this.refCount <= 0 || this.ws || this.reconnectTimer) return;
    if (typeof navigator !== "undefined" && !navigator.onLine) return;
    this.connect();
  }

  private connect(): void {
    if (this.refCount <= 0) return;

    const base = Config.serverUrl.replace(/^http/, "ws");
    const url = `${base}${Config.wsPath}`;
    console.log(`[WS] Connecting to ${url}`);
    this.setStatus(`Connecting to ${url}`);

    const ws = new WebSocket(url);
    this.ws = ws;

    ws.onopen = () => {
      if (this.ws !== ws) return;
      this.retryCount = 0;
      console.log("[WS] Connected");
      this.setStatus("Connected");
    };

    ws.onmessage = (ev: MessageEvent) => {
      if (this.ws !== ws) return;
      let data: Record<string, unknown>;
      try { data = JSON.parse(ev.data); } catch { return; }
      if (!data || typeof data !== "object" || Array.isArray(data) || "type" in data) return;
      
      this.setStatus("Receiving Data");
      
      const entries = Object.entries(data).filter(
        ([, v]) => v && typeof v === "object" && !Array.isArray(v) && "result" in (v as object)
      ) as [string, TickerRecord][];

      if (entries.length) {
        for (const [key, rec] of entries) this.buffer[key] = rec;
        this.bufferDirty = true;
        this.scheduleFlush();
      }
    };

    const onFail = () => {
      this.setStatus("Error / Reconnecting");
      if (this.ws !== ws) { ws.close(); return; }
      this.ws = null;
      ws.onopen = ws.onmessage = ws.onerror = ws.onclose = null;
      ws.close();
      this.scheduleReconnect();
    };
    ws.onerror = onFail;
    ws.onclose = onFail;
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
      this.listeners.forEach((fn) => { try { fn(snapshot); } catch (e) { console.error("[WS] listener error", e); } });
    }, Config.bufferFlushMs);
  }

  // ─── Reconnect ───────────────────────────────────────────────────────────

  private scheduleReconnect(): void {
    if (this.refCount <= 0 || this.reconnectTimer) return;
    const delay = Math.min(Config.reconnect.initialDelayMs * 2 ** this.retryCount, Config.reconnect.maxDelayMs);
    this.retryCount += 1;
    console.log(`[WS] Reconnecting in ${delay}ms (attempt ${this.retryCount})`);
    this.setStatus(`Reconnecting in ${delay/1000}s`);

    this.reconnectTimer = setTimeout(() => { this.reconnectTimer = null; this.ensureConnected(); }, delay);
  }

  // ─── Graceful close ──────────────────────────────────────────────────────

  private scheduleClose(): void {
    if (this.closeTimer) return;
    this.closeTimer = setTimeout(() => {
      this.closeTimer = null;
      if (this.refCount > 0) return;
      if (this.reconnectTimer) { clearTimeout(this.reconnectTimer); this.reconnectTimer = null; }
      this.retryCount = 0;
      if (this.flushTimer) { clearTimeout(this.flushTimer); this.flushTimer = null; }
      const ws = this.ws; this.ws = null;
      if (ws && ws.readyState !== WebSocket.CLOSED) ws.close();
      this.buffer = {};
      this.bufferDirty = false;
      console.log("[WS] Closed (no consumers)");
    }, Config.closeDelayMs);
  }

  private clearCloseTimer(): void {
    if (this.closeTimer) { clearTimeout(this.closeTimer); this.closeTimer = null; }
  }

  // ─── Network ─────────────────────────────────────────────────────────────

  private onOnline(): void {
    if (this.refCount <= 0) return;
    if (this.ws?.readyState === WebSocket.OPEN) return;
    if (this.ws) { this.ws.close(); this.ws = null; }
    this.retryCount = 0;
    if (this.reconnectTimer) { clearTimeout(this.reconnectTimer); this.reconnectTimer = null; }
    this.ensureConnected();
  }

  private onOffline(): void {
    if (this.reconnectTimer) { clearTimeout(this.reconnectTimer); this.reconnectTimer = null; }
    if (this.closeTimer) { clearTimeout(this.closeTimer); this.closeTimer = null; }
    if (this.ws) { this.ws.close(); this.ws = null; }
    this.setStatus("Offline");
  }
}

/** Singleton */
export const liveDataService = new LiveDataService();
