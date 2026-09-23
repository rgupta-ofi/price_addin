# Infinity Live Prices Excel Add-in Guide

Welcome to the Infinity Live Prices Add-in! This tool allows you to stream live Infinity market signals directly into your Excel cells using a simple custom formula.

The current Infinity platform exposes live FX rates through an authenticated HTTPS API. The add-in reads a bearer token from browser storage key `INFINITY_API_TOKEN`; if no token is present, formulas show `Waiting... (Authentication required: set INFINITY_API_TOKEN)`.

## Part 1: How to Install the Add-in

Because this is an internal custom application, you will need to upload a small configuration file (the "Manifest") into your Excel to activate the tool.

### 1. Download the Manifest File
1. Go to the project repository: https://github.com/rgupta-ofi/price_addin (or wherever the manifest.xml is shared internally).
2. Download the file named **manifest.xml** and save it to your computer (e.g., to your Downloads folder).

### 2. Upload it to Excel Online (or Desktop)
1. Open Excel Online in your web browser (or open desktop Excel).
2. Create a new "Blank Workbook".
3. On the ribbon at the top, click the **Insert** tab.
4. Click **Add-ins** (or "Get Add-ins" / "My Add-ins").
5. In the window that pops up, click **Upload My Add-in** (usually near the top right).
6. Click **Browse...** and select the manifest.xml file you downloaded.
7. Click **Upload**.

*Note: Depending on your company's network policies, if Excel Desktop prevents local manifest uploads, simply use Excel Online in your browser  the formulas will perfectly calculate there.*

---

## Part 2: How to Use the Add-in

Once the Add-in is loaded, it operates completely invisibly in the background. You interact with it by typing a custom formula directly into any spreadsheet cell.

### The Missing Formula
The magic formula is:
**`=INFINITY.LIVEPRICE(ticker, field)`**

### Understanding the Inputs
1. **Ticker:** The live FX tenor returned by Infinity. You can use:
  - `"spot"`, `"1m"`, `"2m"`, `"3m"`, and any other tenor returned by the API
  - The display tenor value returned by `=INFINITY.TICKERS()`
   - *Note: Tickers are not case-sensitive.*
2. **Field:** The specific data point you want to stream. Supported fields include:
  - RATE or MID
  - METHOD
  - AS_OF

### Examples
Pick any blank cell and type:

* To get the live spot FX rate:
  `=INFINITY.LIVEPRICE("spot", "RATE")`

* To get the live 1-month FX rate:
  `=INFINITY.LIVEPRICE("1m", "RATE")`

### What to Expect
1. When you hit Enter, the cell might momentarily display #BUSY! or Waiting... as it connects to the live server.
2. Within a second, it will display the live number.
3. As long as your file remains open, the cell will automatically update instantly whenever trading prices tick up or down. You do not need to refresh.

---

## Troubleshooting

- **#NAME? error:** This means Excel hasn't loaded the add-in. Go back to Insert > Add-ins and make sure "Infinity Live Prices" is in your list.
- **Waiting... (Authentication required: set INFINITY_API_TOKEN):** The current Infinity API requires a bearer token. Sign in through the approved flow or set the token directly in the add-in browser context for testing only. Do not store tokens in the repository.
- **Waiting... never resolves:** Ensure you have internet access and that the ticker matches a tenor returned by `=INFINITY.TICKERS()`.
- **No Autocomplete?** Just type out the full `=INFINITY.LIVEPRICE(..)` formula completely and hit Enter. Excel sometimes drops custom autocomplete but the formula will still execute perfectly over the streaming engine.
