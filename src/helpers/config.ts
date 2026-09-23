/**
 * Centralized configuration for the Infinity Live Prices Excel Add-in.
 * Values are injected from .env at build time via dotenv-webpack.
 */
export const Config = {
  /** Base URL of the Infinity API server */
  serverUrl: process.env.SERVER_URL || "https://infinity.ofi.ai",

  /** Legacy WebSocket path for live data */
  wsPath: "/api/realtime/live-data/all",

  /** HTTPS endpoint used by the Infinity portal for live FX rates */
  liveFxPath: "/api/v1/tprm_ps_ims/market-signals/live-fx",

  /** Local storage key for an Infinity API bearer token */
  authTokenStorageKey: "INFINITY_API_TOKEN",

  /** Polling interval for live FX rates */
  pollIntervalMs: 5_000,

  /** Reconnect timing */
  reconnect: {
    initialDelayMs: 1_000,
    maxDelayMs: 30_000,
  },

  /** Buffer / coalescing — how often (ms) to flush batched updates to Excel cells */
  bufferFlushMs: 300,

  /** Delay (ms) before closing WebSocket after last formula is removed */
  closeDelayMs: 10_000,
};
