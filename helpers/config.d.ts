/**
 * Centralized configuration for the Infinity Live Prices Excel Add-in.
 * Values are injected from .env at build time via dotenv-webpack.
 */
export declare const Config: {
    /** Base URL of the Infinity API server */
    serverUrl: string;
    /** Legacy WebSocket path for live data */
    wsPath: string;
    /** HTTPS endpoint used by the Infinity portal for live FX rates */
    liveFxPath: string;
    /** Local storage key for an Infinity API bearer token */
    authTokenStorageKey: string;
    /** Polling interval for live FX rates */
    pollIntervalMs: number;
    /** Reconnect timing */
    reconnect: {
        initialDelayMs: number;
        maxDelayMs: number;
    };
    /** Buffer / coalescing — how often (ms) to flush batched updates to Excel cells */
    bufferFlushMs: number;
    /** Delay (ms) before closing WebSocket after last formula is removed */
    closeDelayMs: number;
};
//# sourceMappingURL=config.d.ts.map