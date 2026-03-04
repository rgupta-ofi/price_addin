/**
 * Centralized configuration for the Infinity Live Prices Excel Add-in.
 * Values are injected from .env at build time via dotenv-webpack.
 */
export const Config = {
  /** Base URL of the WebSocket / API server */
  serverUrl: process.env.SERVER_URL || "https://infinity-qa.ofi.ai",

  /** WebSocket path for live data */
  wsPath: "/api/realtime/live-data/all",

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
