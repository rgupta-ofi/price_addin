# Olam Live Prices Excel Add-in Guide

Welcome to the Olam Live Prices Add-in! This tool allows you to stream real-time commodity prices directly into your Excel cells using a simple custom formula  no side panels, no manual refreshing, and no complex sign-ins required.

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
**=OLAM.LIVEPRICE(ticker, field)**

### Understanding the Inputs
1. **Ticker:** The ID of the commodity or security. You can use:
   - The exact Bloomberg Security ID (e.g., "USDJPY", "CCH6")
   - The Olam internal API key (e.g., "usd-jpy", "cc1")
   - *Note: Tickers are not case-sensitive.*
2. **Field:** The specific data point you want to stream. Supported fields include:
   - MID, BID, ASK
   - LAST_PRICE
   - OPEN, HIGH, LOW
   - VOLUME

### Examples
Pick any blank cell and type:

* To get the live Mid price of CC1:
  =OLAM.LIVEPRICE("cc1", "MID")

* To get the live Ask price of USD/JPY:
  =OLAM.LIVEPRICE("USDJPY", "ASK")

### What to Expect
1. When you hit Enter, the cell might momentarily display #BUSY! or Waiting... as it connects to the live server.
2. Within a second, it will display the live number.
3. As long as your file remains open, the cell will automatically update instantly whenever trading prices tick up or down. You do not need to refresh.

---

## Troubleshooting

- **#NAME? error:** This means Excel hasn't loaded the add-in. Go back to Insert > Add-ins and make sure "Olam Live Prices" is in your list.
- **#BUSY! or Waiting... never resolves:** Ensure you have internet access and that your network doesn't block WebSockets. (Ensure the text ticker isn't misspelled).
- **No Autocomplete?** Just type out the full =OLAM.LIVEPRICE(..) formula completely and hit Enter. Excel sometimes drops custom autocomplete but the formula will still execute perfectly over the streaming engine.
