/**
 * Centralized configuration for the Infinity Live Prices Excel Add-in.
 * Values are injected from .env at build time via dotenv-webpack.
 */
export declare const Config: {
    /** Base URL of the WebSocket / API server */
    serverUrl: string;
    /** WebSocket path for live data */
    wsPath: string;
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