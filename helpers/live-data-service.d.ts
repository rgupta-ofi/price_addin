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
declare class LiveDataService {
    private ws;
    private listeners;
    private statusListeners;
    private currentStatus;
    private refCount;
    private retryCount;
    private reconnectTimer;
    private closeTimer;
    private buffer;
    private bufferDirty;
    private flushTimer;
    constructor();
    getStatus(): string;
    addStatusListener(fn: StatusListener): void;
    removeStatusListener(fn: StatusListener): void;
    private setStatus;
    acquire(): void;
    release(): void;
    subscribe(fn: LiveDataListener): void;
    unsubscribe(fn: LiveDataListener): void;
    private ensureConnected;
    private connect;
    private scheduleFlush;
    private scheduleReconnect;
    private scheduleClose;
    private clearCloseTimer;
    private onOnline;
    private onOffline;
}
/** Singleton */
export declare const liveDataService: LiveDataService;
export {};
//# sourceMappingURL=live-data-service.d.ts.map