# Olam Live Prices — Excel Add-in

A WebSocket-based Excel Add-in that streams live commodity prices directly into your workbook, authenticated via Microsoft Single Sign-On (SSO).

## Architecture

```
┌─────────────────────────────────────────────────┐
│                  Excel Workbook                  │
│                                                  │
│  ┌──────────────┐   ┌────────────────────────┐  │
│  │   Taskpane    │   │    Custom Functions     │  │
│  │  (UI Panel)   │   │  =OLAM.LIVEPRICE(...)  │  │
│  └──────┬───────┘   └───────────┬────────────┘  │
│         │                       │                │
│         └───────┬───────────────┘                │
│                 ▼                                 │
│       ┌─────────────────┐                        │
│       │  LiveDataService │ (WebSocket Client)    │
│       │  + SSO Auth      │                        │
│       └────────┬────────┘                        │
└────────────────┼────────────────────────────────┘
                 │ wss://
                 ▼
         ┌───────────────┐
         │  Your Server   │
         │  /api/realtime │
         │  /live-data/all│
         └───────────────┘
```

## Features

- **Microsoft SSO** — Users sign in seamlessly via Office.auth (with dialog fallback)
- **WebSocket streaming** — Real-time price updates with coalescing buffer
- **Taskpane UI** — Start/stop streaming, view stats, activity log
- **Custom Functions** — Use `=OLAM.LIVEPRICE("ticker", "field")` in any cell for live-streaming prices
- **Auto-reconnect** — Exponential backoff (1s → 30s), network-aware
- **Historical enrichment** — 5-day price deltas computed automatically

## Prerequisites

- **Node.js** >= 18
- **Excel** (Microsoft 365 desktop or Excel on the web)
- **Azure AD App Registration** (for SSO)

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Configure environment

Edit `.env` with your values:

```env
SERVER_URL=https://your-server.example.com
AZURE_CLIENT_ID=your-azure-ad-app-client-id
AZURE_TENANT_ID=common
```

### 3. Configure Azure AD App Registration

1. Go to [Azure Portal → App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Create a new registration (or use existing)
3. Set **Redirect URI**: `https://localhost:3000/taskpane.html` (type: SPA)
4. Under **Expose an API**:
   - Set **Application ID URI** to: `api://localhost:3000/YOUR_CLIENT_ID`
   - Add a scope: `access_as_user`
   - Add authorized client applications:
     - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Office on the web)
     - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office desktop)
     - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office desktop)
     - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)
     - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook desktop)
5. Under **API Permissions**, add: `openid`, `profile`, `User.Read`
6. Copy the **Application (client) ID** into your `.env` and `manifest.xml`

### 4. Update manifest.xml

Replace `YOUR_AZURE_AD_APP_CLIENT_ID_HERE` in `manifest.xml` (WebApplicationInfo section) with your actual Azure AD App Client ID.

### 5. Install dev certificates

```bash
npx office-addin-dev-certs install
```

### 6. Run the dev server

```bash
npm run dev-server
```

### 7. Sideload into Excel

**Desktop (Windows):**
```bash
npm start
```

**Excel on the Web:**
1. Open Excel Online
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Upload `manifest.xml`

## Usage

### Taskpane

1. Click the **Live Prices** button on the Home ribbon tab
2. Click **Sign In (SSO)** to authenticate
3. Click **▶ Start Streaming** to begin receiving live prices
4. Prices auto-populate in the active sheet as a formatted table
5. Use **↻ Reload 5d Prices** to refresh historical delta data
6. Click **■ Stop** to disconnect

### Custom Functions (in-cell streaming)

Type into any cell:

| Formula | Description |
|---|---|
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "MID")` | Mid price for CLZ4 |
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "BID")` | Bid price |
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "ASK")` | Ask price |
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "LAST_PRICE")` | Last trade price |
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "VOLUME")` | Volume |
| `=OLAM.LIVEPRICE("CLZ4 Comdty", "price_delta_days_ago_5")` | 5-day price change |

These cells update in real-time as new data arrives via WebSocket.

## Project Structure

```
price_addin/
├── manifest.xml                     # Office Add-in manifest (SSO config)
├── package.json
├── tsconfig.json
├── webpack.config.js
├── .env                             # Server URL & Azure AD config
├── src/
│   ├── helpers/
│   │   ├── config.ts                # Centralized configuration
│   │   ├── sso-auth.ts              # SSO + fallback dialog authentication
│   │   ├── live-data-service.ts     # WebSocket client (singleton)
│   │   └── excel-writer.ts          # Writes snapshots into Excel tables
│   ├── taskpane/
│   │   ├── taskpane.html            # Taskpane UI
│   │   ├── taskpane.ts              # Taskpane logic
│   │   └── taskpane.css             # Styles
│   ├── functions/
│   │   ├── functions.ts             # Custom streaming functions
│   │   ├── functions.json           # Function metadata
│   │   └── functions.html           # Functions runtime host
│   ├── commands/
│   │   ├── commands.ts              # Ribbon commands
│   │   └── commands.html            # Commands runtime host
│   └── assets/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-80.png
└── useLiveData-OagUQOyI.js          # Original reference client
```

## Server Requirements

Your server needs to expose:

| Endpoint | Method | Description |
|---|---|---|
| `/api/auth/token` | POST | Exchange Office SSO bootstrap token for app token (OBO flow) |
| `/api/auth/dialog-login` | GET | Fallback login page for dialog-based auth |
| `/api/realtime/live-data/all` | WebSocket | Live price stream (accepts `?access_token=...`) |
| `/api/prices_days_ago_5` | GET | Historical prices (5 days ago) |

## Build for Production

```bash
npm run build
```

Output is in `dist/`. Deploy the contents to your hosting server and update the URLs in `manifest.xml` accordingly.

## Troubleshooting

| Issue | Solution |
|---|---|
| SSO fails with error 13001 | Ensure Azure AD app is configured correctly |
| SSO fails with error 13003 | User needs to consent — try `allowConsentPrompt: true` |
| WebSocket won't connect | Check `SERVER_URL` in `.env`, ensure server accepts `access_token` query param |
| Custom functions not showing | Ensure `functions.json` is served correctly; re-sideload the add-in |
| Certificate errors | Run `npx office-addin-dev-certs install` again |
