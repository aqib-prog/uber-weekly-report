// main.js — Uber Weekly Reporter - CLEAN MVP VERSION (panel-scoped reads + fare/serviceFee/taxes)
// -----------------------------------------------------------------------
// Electron main process + Playwright scraping
// -----------------------------------------------------------------------

const DEBUG_RANGE =
  process.env.DEBUG_RANGE === "1" || process.env.PWDEBUG === "1";

const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const { chromium } = require("playwright"); // Browser automation
const ExcelJS = require("exceljs"); // Excel writer
const store = require("./secure-store"); // AES-GCM + keychain wrapper

// ---- Uber org constants ----
const ORG_ID = "4e7e783a-5d53-4c1c-a683-ef15c6ddbeae";
const ORG_URL = `https://supplier.uber.com/orgs/${ORG_ID}/vehicles`;
const EARNINGS_URL = `https://supplier.uber.com/orgs/${ORG_ID}/earnings`;

// ---- Window + state ----
let win = null;

// Visible Playwright instance for attended login
let pw = { browser: null, context: null, page: null };

// Warm, shared headless browser to make checks/runs snappy
let warm = { browser: null };
async function getWarmBrowser() {
  if (warm.browser) return warm.browser;
  warm.browser = await chromium.launch({
    headless: true,
    args: ["--disable-dev-shm-usage"],
  });
  return warm.browser;
}
app.on("before-quit", async () => {
  try {
    await warm.browser?.close();
  } catch {}
});

// ---- Electron window ----
function createWindow() {
  win = new BrowserWindow({
    width: 1180,
    height: 780,
    minWidth: 1020,
    minHeight: 680,
    backgroundColor: "#F8FAFC",
    title: "Uber Weekly Reporter",
    webPreferences: { preload: path.join(__dirname, "preload.js") },
  });
  win.removeMenu();
  win.loadFile("index.html");
  // win.webContents.openDevTools({ mode: "detach" }); // uncomment for debugging
}
app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ------------------------------------------------------------------
// Helpers
// ------------------------------------------------------------------

// Parse "AED 1,234.56" or "-AED 23.45" → Number (keep sign)
function moneyToNumber(txt) {
  if (!txt) return 0;
  const neg = /-\s*/.test(txt) || /\(\s*AED/i.test(txt);
  const n =
    parseFloat(
      String(txt)
        .replace(/[^0-9.,-]/g, "")
        .replace(/,/g, "")
    ) || 0;
  return neg ? -Math.abs(n) : n;
}

// Read first text from a locator safely
async function firstText(loc) {
  try {
    return (await loc.first().innerText()).trim();
  } catch {
    return "";
  }
}

// Select maximum rows per page (best-effort)
async function maximizeRowsPerPage(page) {
  console.log("Attempting to maximize rows per page...");
  try {
    const rowsControl = page
      .locator("button, div")
      .filter({ hasText: /\d+\s*rows?/i })
      .first();

    if (await rowsControl.isVisible().catch(() => false)) {
      await rowsControl.click({ timeout: 2000 });
      await page.waitForTimeout(400);

      const options = page
        .locator(
          '[role="listbox"] button, [role="menu"] button, .dropdown-item'
        )
        .filter({ hasText: /^\s*\d+\s*$/ });
      const n = await options.count();
      if (n > 0) {
        await options.nth(n - 1).click({ timeout: 1000 }); // pick largest
        await page.waitForTimeout(500);
      }
    }
  } catch (e) {
    console.log("maximizeRowsPerPage: " + e.message);
  }
}

// Find the right-side details drawer/panel (scope for all reads)
function detailsPanelLocator(page) {
  // Prefer BaseWeb drawer, then any modal/dialog that's wide enough
  return page
    .locator(
      [
        '[data-baseweb="drawer"]',
        'div[role="dialog"][aria-modal="true"]',
        'div[role="dialog"]',
      ].join(",")
    )
    .filter({
      has: page.locator("text=/Total earnings|Payout|Trips/i"),
    })
    .last();
}

// Expand a labeled section **inside the panel** so its children become visible
async function expandByLabelInPanel(page, panel, labelText) {
  try {
    const header = panel.getByText(new RegExp(`^${labelText}$`, "i")).first();
    const toggle = header.locator(
      "xpath=ancestor::*[self::div or self::section][1]//button[@aria-expanded]"
    );
    if (await toggle.count()) {
      const expanded = await toggle.first().getAttribute("aria-expanded");
      if (expanded !== "true") {
        await toggle.first().click({ timeout: 1200 });
        await page.waitForTimeout(200);
      }
      return true;
    }
  } catch {}
  // fallback: click the header itself
  try {
    await panel
      .getByText(new RegExp(`^${labelText}$`, "i"))
      .first()
      .click({
        timeout: 1200,
      });
    await page.waitForTimeout(150);
    return true;
  } catch {}
  return false;
}

// Read "AED ..." shown on the **same row** as the given label, inside the panel
async function readAedFromPanel(panel, labelRe) {
  try {
    const label = panel.getByText(labelRe).first();
    if (!(await label.isVisible().catch(() => false))) return 0;

    const row = label.locator(
      "xpath=ancestor::*[self::div or self::li or self::tr][1]"
    );
    const valueEl = row
      .locator(
        'xpath=.//*[contains(normalize-space(.), "AED") and not(self::script) and not(self::style)]'
      )
      .last();

    const txt = await firstText(valueEl);
    return moneyToNumber(txt);
  } catch {
    return 0;
  }
}

// Trips and distance from panel text
async function readTripsAndDistance(panel) {
  const t = (await panel.innerText().catch(() => "")) || "";
  const trips = parseInt((t.match(/Trips\s*(\d+)/i) || [])[1] || "0", 10) || 0;
  const distance =
    parseFloat(
      ((t.match(/([\d,]+(?:\.\d+)?)\s*km/i) || [])[1] || "0").replace(/,/g, "")
    ) || 0;
  return { trips, distance };
}

// NEW: Check if Next button is enabled
async function isNextButtonEnabled(page) {
  try {
    const nextButton = page.locator("button", { hasText: /Next/i }).first();
    if (!(await nextButton.isVisible().catch(() => false))) return false;
    const dis = await nextButton.getAttribute("disabled");
    const aria = await nextButton.getAttribute("aria-disabled");
    return dis === null && aria !== "true";
  } catch {
    return false;
  }
}

// NEW: Click Next button and wait for page change
async function gotoNextPage(page) {
  try {
    console.log("Attempting to go to next page...");
    const nextButton = page.locator("button", { hasText: /Next/i }).first();

    if (!(await isNextButtonEnabled(page))) {
      console.log("Next button is disabled - at last page");
      return false;
    }

    const firstBefore = await page
      .locator("div[role='row'], table tbody tr")
      .first()
      .innerText()
      .catch(() => "");

    await nextButton.click({ timeout: 1500 });
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.waitForTimeout(500);

    // Basic change detection
    const firstAfter = await page
      .locator("div[role='row'], table tbody tr")
      .first()
      .innerText()
      .catch(() => "");
    return firstAfter && firstAfter !== firstBefore;
  } catch (e) {
    console.log("gotoNextPage: " + e.message);
    return false;
  }
}

// ------------------------------------------------------------------
// RANGE HELPERS (as requested)
// ------------------------------------------------------------------
function toISO(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// UPDATED: Better toLong function to handle the display format
function toLong(d) {
  return d.toLocaleDateString("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
  });
}

// UPDATED: Enhanced date range reading to handle the actual Uber format
async function getSelectedRangeFromPage(page) {
  console.log("Reading selected date range from Uber page...");

  try {
    // Multiple strategies to find the date range element
    const dateSelectors = [
      // Look for buttons or divs containing the date pattern with ordinals (1st, 2nd, 3rd, 4th)
      "button:has-text(/[A-Za-z]{3}\\s+\\d{1,2}(st|nd|rd|th),\\s+\\d{4}.*\\d{2}:\\d{2}\\s+(AM|PM)/i)",
      'div[role="button"]:has-text(/[A-Za-z]{3}\\s+\\d{1,2}(st|nd|rd|th),\\s+\\d{4}.*\\d{2}:\\d{2}\\s+(AM|PM)/i)',
      // Fallback: any element containing the date pattern
      "*:has-text(/[A-Za-z]{3}\\s+\\d{1,2}(st|nd|rd|th),\\s+\\d{4}.*\\d{2}:\\d{2}\\s+(AM|PM)/i)",
      // Additional fallback for different formats
      "button:has-text(/[A-Za-z]{3,9}\\s+\\d{1,2},\\s+\\d{4}/)",
      'div[role="button"]:has-text(/[A-Za-z]{3,9}\\s+\\d{1,2},\\s+\\d{4}/)',
    ];

    let dateText = "";

    for (const selector of dateSelectors) {
      try {
        const element = page.locator(selector).first();
        if (await element.isVisible({ timeout: 2000 })) {
          dateText = await element.innerText();
          console.log(
            `Found date range with selector "${selector}": "${dateText}"`
          );
          break;
        }
      } catch (e) {
        console.log(`Selector "${selector}" failed: ${e.message}`);
        continue;
      }
    }

    if (!dateText) {
      console.log(
        "Could not find date range element, trying page text search..."
      );

      // Fallback: search the entire page text
      const pageText = await page.evaluate(() => document.body.textContent);
      const match = pageText.match(
        /([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4}\s+\d{2}:\d{2}\s+(?:AM|PM))\s*[-–—]\s*([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4}\s+\d{2}:\d{2}\s+(?:AM|PM))/
      );

      if (match) {
        dateText = `${match[1]} - ${match[2]}`;
        console.log(`Found date range in page text: "${dateText}"`);
      }
    }

    if (!dateText) {
      console.log("No date range found on page");
      return null;
    }

    // Parse the date range text
    // Handle format: "Sep 1st, 2025 04:01 AM - Sep 4th, 2025 06:36 PM"
    const cleanText = dateText.replace(/[\u2013\u2014]/g, "-"); // normalize dashes

    // Updated regex to handle ordinals (1st, 2nd, 3rd, 4th) and time
    const match = cleanText.match(
      /([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4})(?:\s+\d{2}:\d{2}\s+(?:AM|PM))?\s*[-–—]\s*([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4})(?:\s+\d{2}:\d{2}\s+(?:AM|PM))?/
    );

    if (!match) {
      // Fallback: try simpler format without ordinals
      const simpleMatch = cleanText.match(
        /([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})\s*[-–—]\s*([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})/
      );
      if (simpleMatch) {
        const start = new Date(simpleMatch[1]);
        const end = new Date(simpleMatch[2]);

        if (!isNaN(start) && !isNaN(end)) {
          return {
            start,
            end,
            startISO: toISO(start),
            endISO: toISO(end),
            displayText: dateText, // Keep original text for display
            long: `${toLong(start)} – ${toLong(end)}`,
          };
        }
      }

      console.log(`Could not parse date range: "${dateText}"`);
      return null;
    }

    // Remove ordinals for Date parsing (1st -> 1, 2nd -> 2, etc.)
    const startDateStr = match[1].replace(/(\d{1,2})(st|nd|rd|th)/, "$1");
    const endDateStr = match[2].replace(/(\d{1,2})(st|nd|rd|th)/, "$1");

    console.log(`Parsing dates: "${startDateStr}" and "${endDateStr}"`);

    const start = new Date(startDateStr);
    const end = new Date(endDateStr);

    if (isNaN(start) || isNaN(end)) {
      console.log(`Invalid dates parsed: start=${start}, end=${end}`);
      return null;
    }

    const result = {
      start,
      end,
      startISO: toISO(start),
      endISO: toISO(end),
      displayText: dateText, // Keep the original format for display
      long: `${toLong(start)} – ${toLong(end)}`,
    };

    console.log("Successfully parsed date range:", result);
    return result;
  } catch (error) {
    console.error("Error reading date range from page:", error.message);
    return null;
  }
}

// ------------------------------------------------------------------
// DRIVER DETAILS (panel-scoped, correct Tips & Payout + Fare/ServiceFee/Taxes)
// ------------------------------------------------------------------
async function extractDriverDetails(page, driverName) {
  console.log(`Extracting details for driver: ${driverName}`);

  try {
    // Find the right-side details panel and keep all queries scoped inside it
    const panel = detailsPanelLocator(page);
    await panel.waitFor({ state: "visible", timeout: 4000 });

    // Make sure "Total earnings" section is expanded (so breakdown rows are visible)
    await expandByLabelInPanel(page, panel, "Total earnings");

    // Read values strictly from inside the panel:
    const totalEarnings = await readAedFromPanel(panel, /^(Total earnings)$/i);

    // Breakdown items under Total earnings:
    const fare = await readAedFromPanel(panel, /^Fare$/i);
    const serviceFee = await readAedFromPanel(panel, /^Service\s*Fee$/i);

    // NEW: Other earnings field
    const otherEarnings = await readAedFromPanel(panel, /^Other\s*earnings$/i);

    const taxes = await readAedFromPanel(panel, /^Taxes?$/i);
    const tips = await readAedFromPanel(panel, /^Tip$/i);

    // Refunds & Expenses and Adjustments (usually separate blocks inside panel)
    const refundsExpenses =
      (await readAedFromPanel(panel, /^Refunds\s*&\s*Expenses$/i)) || 0;

    const adjustments =
      (await readAedFromPanel(
        panel,
        /^(Adjustments from previous periods|Adjustments)$/i
      )) || 0;

    // Payout (usually negative)
    const payout = await readAedFromPanel(panel, /^Payout$/i);

    // Trips + Distance from the same panel
    const { trips, distance } = await readTripsAndDistance(panel);

    // Keep your current net formula (does NOT subtract tips)
    const netEarnings = totalEarnings + refundsExpenses + adjustments + payout;

    const result = {
      name: driverName,
      totalEarnings,
      fare,
      serviceFee,
      otherEarnings, // NEW field
      taxes,
      tips,
      refundsExpenses,
      adjustments,
      payout,
      netEarnings,
      trips,
      distance,
    };

    console.log(`=== FINAL DATA for ${driverName} ===`);
    console.log(`Total Earnings: AED ${totalEarnings}`);
    console.log(`Fare: AED ${fare}`);
    console.log(`Service Fee: AED ${serviceFee}`);
    console.log(`Other Earnings: AED ${otherEarnings}`); // NEW log
    console.log(`Taxes: AED ${taxes}`);
    console.log(`Tips: AED ${tips}`);
    console.log(`Payout: AED ${payout}`);
    console.log(`Trips: ${trips} | Distance: ${distance} km`);
    console.log(`Net Earnings: AED ${netEarnings}`);

    return result;
  } catch (error) {
    console.error(`extractDriverDetails error: ${error.message}`);
    return {
      name: driverName,
      totalEarnings: 0,
      fare: 0,
      serviceFee: 0,
      otherEarnings: 0, // NEW field with default 0
      taxes: 0,
      tips: 0,
      refundsExpenses: 0,
      adjustments: 0,
      payout: 0,
      netEarnings: 0,
      trips: 0,
      distance: 0,
    };
  }
}
// ------------------------------------------
// IPC handlers used by the renderer (UI)
// ------------------------------------------

ipcMain.handle("has-session", async () => {
  try {
    return fs.existsSync(store.sessionPath());
  } catch {
    return false;
  }
});

ipcMain.handle("open-login", async () => {
  try {
    if (pw.browser && pw.page) {
      try {
        await pw.page.bringToFront();
      } catch {}
      return true;
    }
    pw.browser = await chromium.launch({ headless: false });
    pw.context = await pw.browser.newContext({
      viewport: { width: 1360, height: 900 },
      acceptDownloads: true,
    });
    pw.page = await pw.context.newPage();
    await pw.page.goto(ORG_URL, { waitUntil: "domcontentloaded" });
    return true;
  } catch (err) {
    console.error("open-login failed:", err);
    try {
      await pw.browser?.close();
    } catch {}
    pw = { browser: null, context: null, page: null };
    return false;
  }
});

ipcMain.handle("save-session", async () => {
  try {
    if (!pw.context || !pw.page)
      return { ok: false, msg: "Login window not open." };
    const url = pw.page.url();
    if (!/supplier\.uber\.com/.test(url)) {
      try {
        await pw.page.waitForSelector("text=Earnings", { timeout: 4000 });
      } catch {
        return {
          ok: false,
          msg: "Please finish sign-in to the Supplier dashboard, then try again.",
        };
      }
    }
    const state = await pw.context.storageState();
    const file = await store.encryptToFile(
      Buffer.from(JSON.stringify(state), "utf-8")
    );
    try {
      await pw.browser?.close();
    } catch {}
    pw = { browser: null, context: null, page: null };
    return { ok: true, file };
  } catch (err) {
    console.error("save-session failed:", err);
    try {
      await pw.browser?.close();
    } catch {}
    pw = { browser: null, context: null, page: null };
    return {
      ok: false,
      msg: "Failed to save session. See console for details.",
    };
  }
});

ipcMain.handle("smoke-earnings", async () => {
  try {
    const buf = await store.decryptFromFile();
    const state = JSON.parse(buf.toString("utf-8"));
    const browser = await getWarmBrowser();
    const context = await browser.newContext({ storageState: state });
    const page = await context.newPage();

    await page.goto(EARNINGS_URL, {
      waitUntil: "domcontentloaded",
      timeout: 6000,
    });
    if (/auth\.uber\.com/.test(page.url())) {
      await context.close();
      return {
        ok: false,
        reason: "expired",
        msg: "Session expired. Please connect again.",
      };
    }

    try {
      await page.waitForSelector("text=Driver earnings", { timeout: 5000 });
    } catch {
      await page.waitForSelector("text=Earnings", { timeout: 5000 });
    }

    await context.close();
    return { ok: true };
  } catch (err) {
    console.error("smoke-earnings failed:", err);
    if (/no such file|ENOENT/i.test(String(err)))
      return {
        ok: false,
        reason: "nosession",
        msg: "No saved session. Click Connect Uber, then Save Session.",
      };
    return { ok: false, msg: String(err.message || err) };
  }
});

// RESTORED: Full Uber functionality (manual date, then scrape)
ipcMain.handle("open-uber-for-manual-setup", async () => {
  console.log("=== STARTING UBER SETUP ===");

  try {
    const buf = await store.decryptFromFile();
    const state = JSON.parse(buf.toString("utf-8"));

    const browser = await chromium.launch({ headless: false, slowMo: 100 });
    const context = await browser.newContext({ storageState: state });
    const page = await context.newPage();

    await page.goto(EARNINGS_URL, {
      waitUntil: "domcontentloaded",
      timeout: 15000,
    });

    if (/auth\.uber\.com/.test(page.url())) {
      await browser.close();
      return {
        ok: false,
        reason: "expired",
        msg: "Session expired. Please connect again.",
      };
    }

    try {
      await page.waitForSelector("text=Driver earnings", { timeout: 10000 });
    } catch {
      await page
        .waitForSelector("text=Earnings", { timeout: 10000 })
        .catch(() => {});
    }

    global.manualBrowser = browser;
    global.manualContext = context;
    global.manualPage = page;

    return { ok: true, msg: "Uber page opened for manual date setting" };
  } catch (err) {
    if (/no such file|ENOENT/i.test(String(err))) {
      return {
        ok: false,
        reason: "nosession",
        msg: "No saved session found. Please connect Uber first.",
      };
    }
    return { ok: false, msg: String(err.message || err) };
  }
});

ipcMain.handle("run-automation", async () => {
  console.log("Starting automation with manually set date range...");

  try {
    if (!global.manualBrowser || !global.manualPage) {
      return {
        ok: false,
        msg: "No browser session found. Please open Uber page first.",
      };
    }

    const page = global.manualPage;
    const browser = global.manualBrowser;

    const url = page.url();
    if (!url.includes("supplier.uber.com") || !url.includes("earnings")) {
      return {
        ok: false,
        msg: "Browser is not on the earnings page. Please navigate to earnings page.",
      };
    }

    // Read the selected date-range from the page chip (once)
    const selectedRange = await getSelectedRangeFromPage(page);
    if (selectedRange) {
      console.log(
        "Selected range:",
        selectedRange.displayText || selectedRange.long
      );
    } else {
      console.log(
        "Could not parse date range chip; falling back to today for row stamps."
      );
    }

    // Find driver rows
    let driverRows = null;
    for (const selector of [
      "div[role='row']",
      "[data-testid*='driver']",
      "table tbody tr",
      "tr:has(img)",
    ]) {
      const rows = page.locator(selector);
      const count = await rows.count();
      console.log(`Trying selector "${selector}": ${count} rows`);
      if (count > 1) {
        driverRows = rows;
        break;
      }
    }
    if (!driverRows) {
      return {
        ok: false,
        reason: "nodata",
        msg: "Could not find driver table.",
      };
    }

    await maximizeRowsPerPage(page);
    await page.waitForTimeout(500);

    const allDriverData = [];
    let pageNumber = 1;

    do {
      const currentRows = page
        .locator("div[role='row']")
        .filter({ hasNotText: /Driver name|Total earnings/ });
      const count = await currentRows.count();
      console.log(`Rows on page ${pageNumber}: ${count}`);

      if (!count) break;

      for (let i = 0; i < count; i++) {
        try {
          const row = currentRows.nth(i);

          const nameEl = row
            .locator("div, span, td")
            .filter({ hasText: /^(?!AED|0\.00|[\d,]+\.\d{2}$)[A-Za-z]/ })
            .first();
          const driverName =
            ((await nameEl.textContent().catch(() => "")) || "").trim() ||
            `Driver ${i + 1}`;
          console.log(`Driver: ${driverName}`);

          // open right drawer (try last button with icon)
          let openOk = false;
          for (const s of [
            "button:last-child",
            "button:has(svg):last-child",
            'button[aria-label*="expand"]',
            'button[aria-label*="details"]',
            'button[aria-label*="more"]',
            "button:has(svg)",
            '[role="button"]:has(svg)',
          ]) {
            try {
              const btns = row.locator(s);
              if ((await btns.count()) > 0) {
                const b = btns.last();
                if (await b.isVisible().catch(() => false)) {
                  await b.scrollIntoViewIfNeeded().catch(() => {});
                  await b.click({ timeout: 1500 });
                  openOk = true;
                  break;
                }
              }
            } catch {}
          }
          if (!openOk) {
            // fallback: click row
            await row.click({ timeout: 1500 }).catch(() => {});
          }

          // ensure panel visible
          const panel = detailsPanelLocator(page);
          await panel.waitFor({ state: "visible", timeout: 4000 });

          // extract strictly from inside panel
          const details = await extractDriverDetails(page, driverName);

          // Stamp each driver row with the real dates
          details.startDate = selectedRange
            ? selectedRange.startISO
            : toISO(new Date());
          details.endDate = selectedRange
            ? selectedRange.endISO
            : toISO(new Date());

          allDriverData.push(details);

          // close panel
          await page.keyboard.press("Escape").catch(() => {});
          await page.waitForTimeout(300);
        } catch (e) {
          console.log("row error: " + e.message);
        }
      }

      const hasNext = await gotoNextPage(page);
      if (hasNext) {
        pageNumber++;
        await page
          .waitForSelector("div[role='row']", { timeout: 5000 })
          .catch(() => {});
        await page.waitForTimeout(400);
      } else {
        break;
      }
    } while (pageNumber <= 10);

    if (!allDriverData.length) {
      return {
        ok: false,
        reason: "nodata",
        msg: "No driver data was extracted.",
      };
    }

    // Excel writing (with new columns Fare/Service Fee/Taxes)
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Weekly Earnings Report");

    ws.columns = [
      { header: "Driver", key: "name", width: 30 },
      {
        header: "Total Earnings",
        key: "totalEarnings",
        width: 16,
        style: { numFmt: "#,##0.00" },
      },
      { header: "Fare", key: "fare", width: 14, style: { numFmt: "#,##0.00" } },
      {
        header: "Service Fee",
        key: "serviceFee",
        width: 14,
        style: { numFmt: "#,##0.00" },
      },
      {
        header: "Other Earnings", // NEW column
        key: "otherEarnings",
        width: 16,
        style: { numFmt: "#,##0.00" },
      },
      {
        header: "Taxes",
        key: "taxes",
        width: 12,
        style: { numFmt: "#,##0.00" },
      },
      { header: "Tips", key: "tips", width: 12, style: { numFmt: "#,##0.00" } },
      {
        header: "Refunds & Expenses",
        key: "refundsExpenses",
        width: 18,
        style: { numFmt: "#,##0.00" },
      },
      {
        header: "Adjustments",
        key: "adjustments",
        width: 15,
        style: { numFmt: "#,##0.00" },
      },
      {
        header: "Payout",
        key: "payout",
        width: 15,
        style: { numFmt: "#,##0.00" },
      },
      {
        header: "Net Earnings",
        key: "netEarnings",
        width: 16,
        style: { numFmt: "#,##0.00" },
      },
      { header: "Trips", key: "trips", width: 10 },
      {
        header: "Distance (km)",
        key: "distance",
        width: 15,
        style: { numFmt: "0.00" },
      },
    ];

    // --- Range banner above the table ---
    const bannerText = selectedRange
      ? `Range: ${selectedRange.displayText || selectedRange.long}`
      : "Range: (unknown)";
    ws.spliceRows(1, 0, [bannerText]); // insert as first row
    ws.mergeCells(1, 1, 1, ws.columns.length); // merge A1..last col
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { horizontal: "left" };
    // Color for the banner row
    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDBEEF4" }, // light blue banner
    };

    // Header is now on row 2
    ws.getRow(2).font = { bold: true };
    ws.getRow(2).alignment = { vertical: "middle", horizontal: "center" };
    ws.getRow(2).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6E6FA" },
    };

    allDriverData.forEach((driver, idx) => {
      const row = ws.addRow(driver);
      if (idx % 2 === 1) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF8F8F8" },
        };
      }
    });

    // Payout yellow: skip rows 1–2
    ws.getColumn("payout").eachCell((cell, r) => {
      if (r > 2) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
      }
    });

    // Borders
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Freeze the banner row
    ws.views = [{ state: "frozen", ySplit: 1 }];

    // Totals row
    const totalsRowNumber = ws.rowCount + 2;
    const totalsRow = ws.getRow(totalsRowNumber);
    totalsRow.getCell(1).value = "TOTALS";
    totalsRow.getCell(1).font = { bold: true };

    const sum = (arr, k) => arr.reduce((s, r) => s + (Number(r[k]) || 0), 0);

    totalsRow.getCell(2).value = sum(allDriverData, "totalEarnings");
    totalsRow.getCell(3).value = sum(allDriverData, "fare");
    totalsRow.getCell(4).value = sum(allDriverData, "serviceFee");
    totalsRow.getCell(5).value = sum(allDriverData, "otherEarnings");
    totalsRow.getCell(6).value = sum(allDriverData, "taxes");
    totalsRow.getCell(7).value = sum(allDriverData, "tips");
    totalsRow.getCell(8).value = sum(allDriverData, "refundsExpenses");
    totalsRow.getCell(9).value = sum(allDriverData, "adjustments");
    totalsRow.getCell(10).value = sum(allDriverData, "payout");
    totalsRow.getCell(11).value = sum(allDriverData, "netEarnings");
    totalsRow.getCell(12).value = sum(allDriverData, "trips");
    totalsRow.getCell(13).value = sum(allDriverData, "distance");

    totalsRow.font = { bold: true };
    totalsRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD3D3D3" },
    };
    // number formats
    [2, 3, 4, 5, 6, 7, 8, 9, 10, 11].forEach(
      (c) => (totalsRow.getCell(c).numFmt = "#,##0.00")
    );
    totalsRow.getCell(13).numFmt = "0.00";

    // Use actual date range in filename instead of today's date
    let fileName;
    if (selectedRange) {
      fileName = `Uber-Weekly-Report-${selectedRange.startISO}-to-${selectedRange.endISO}.xlsx`;
    } else {
      const timestamp = new Date().toISOString().split("T")[0];
      fileName = `Uber-Weekly-Report-${timestamp}.xlsx`;
    }
    const filePath = path.join(app.getPath("desktop"), fileName);

    // Add summary table below main table
    const summaryStartRow = ws.rowCount + 3;

    // Summary table header with date range
    const summaryHeaderRow = ws.getRow(summaryStartRow);
    const summaryBannerText = selectedRange
      ? `SUMMARY - ${selectedRange.displayText || selectedRange.long}`
      : "SUMMARY - Range: (unknown)";

    summaryHeaderRow.getCell(1).value = summaryBannerText;
    ws.mergeCells(summaryStartRow, 1, summaryStartRow, 7);
    summaryHeaderRow.font = { bold: true, size: 14 };
    summaryHeaderRow.alignment = { horizontal: "center" };
    summaryHeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFECEEDF" },
    };

    // Add summary table column headers
    const summaryColHeaderRow = ws.getRow(summaryStartRow + 1);
    const summaryHeaders = [
      "NAMES",
      "TOTAL EARNINGS",
      "REFUNDS/EXPENSES",
      "ADJUSTMENTS",
      "PAYOUT",
      "NET EARNINGS",
      "TOTAL TRIPS",
      "TIPS",
    ];

    summaryHeaders.forEach((header, index) => {
      const cell = summaryColHeaderRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD9E2F3" }, // Light blue
      };
    });

    // Set column widths for summary table
    const summaryWidths = [25, 16, 18, 15, 15, 16, 12, 12];
    for (let i = 0; i < summaryWidths.length; i++) {
      // Only set width if it's larger than current width
      const currentWidth = ws.getColumn(i + 1).width || 0;
      if (summaryWidths[i] > currentWidth) {
        ws.getColumn(i + 1).width = summaryWidths[i];
      }
    }

    // Add summary data rows
    allDriverData.forEach((driver, index) => {
      const rowNum = summaryStartRow + 2 + index;
      const row = ws.getRow(rowNum);

      // Calculate values using the formulas we identified
      const correctedTotalEarnings =
        driver.fare +
        driver.serviceFee +
        driver.otherEarnings +
        driver.taxes +
        driver.tips;
      const correctedNetEarnings =
        correctedTotalEarnings +
        driver.refundsExpenses +
        driver.adjustments +
        driver.payout;

      // Set values
      row.getCell(1).value = driver.name;
      row.getCell(2).value = correctedTotalEarnings;
      row.getCell(3).value = driver.refundsExpenses;
      row.getCell(4).value = driver.adjustments;
      row.getCell(5).value = driver.payout;
      row.getCell(6).value = correctedNetEarnings;
      row.getCell(7).value = driver.trips;
      row.getCell(8).value = driver.tips;

      // Apply number formatting
      [2, 3, 4, 5, 6, 8].forEach((col) => {
        row.getCell(col).numFmt = "#,##0.00";
      });

      // Highlight payout column in yellow
      row.getCell(5).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F8FAB4" }, // Yellow
      };

      // Alternate row colors
      if (index % 2 === 1) {
        [1, 2, 3, 4, 6, 7, 8].forEach((col) => {
          // Skip payout column (5)
          if (col !== 5) {
            row.getCell(col).fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFF8F8F8" }, // Light gray
            };
          }
        });
      }
    });

    // Add summary totals row
    const summaryTotalsRowNum = summaryStartRow + 2 + allDriverData.length + 1;
    const summaryTotalsRow = ws.getRow(summaryTotalsRowNum);

    summaryTotalsRow.getCell(1).value = "TOTALS";
    summaryTotalsRow.getCell(1).font = { bold: true };

    // Calculate corrected totals using proper formulas
    const totalCorrectedEarnings = allDriverData.reduce(
      (sum, driver) =>
        sum +
        (driver.fare +
          driver.serviceFee +
          driver.otherEarnings +
          driver.taxes +
          driver.tips),
      0
    );
    const totalCorrectedNet = allDriverData.reduce(
      (sum, driver) =>
        sum +
        (driver.fare +
          driver.serviceFee +
          driver.otherEarnings +
          driver.taxes +
          driver.tips +
          driver.refundsExpenses +
          driver.adjustments +
          driver.payout),
      0
    );

    summaryTotalsRow.getCell(2).value = totalCorrectedEarnings;
    summaryTotalsRow.getCell(3).value = sum(allDriverData, "refundsExpenses");
    summaryTotalsRow.getCell(4).value = sum(allDriverData, "adjustments");
    summaryTotalsRow.getCell(5).value = sum(allDriverData, "payout");
    summaryTotalsRow.getCell(6).value = totalCorrectedNet;
    summaryTotalsRow.getCell(7).value = sum(allDriverData, "trips");
    summaryTotalsRow.getCell(8).value = sum(allDriverData, "tips");

    // Format totals row
    summaryTotalsRow.font = { bold: true };
    summaryTotalsRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD3D3D3" }, // Gray
    };

    [2, 3, 4, 5, 6, 8].forEach((col) => {
      summaryTotalsRow.getCell(col).numFmt = "#,##0.00";
    });

    // Apply borders to all summary table cells
    for (let row = summaryStartRow; row <= summaryTotalsRowNum; row++) {
      for (let col = 1; col <= 8; col++) {
        const cell = ws.getCell(row, col);
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }
    }

    await wb.xlsx.writeFile(filePath);
    console.log(`📊 Excel file created: ${filePath}`);

    // Clean up browser
    await browser.close();
    global.manualBrowser = null;
    global.manualContext = null;
    global.manualPage = null;

    return {
      ok: true,
      file: filePath,
      driversProcessed: allDriverData.length,
      pagesProcessed: pageNumber,
    };
  } catch (err) {
    console.error("run-automation failed:", err);

    if (global.manualBrowser) {
      try {
        await global.manualBrowser.close();
      } catch {}
      global.manualBrowser = null;
      global.manualContext = null;
      global.manualPage = null;
    }
    return { ok: false, msg: String(err.message || err) };
  }
});

// Back-compat handler retained
ipcMain.handle("run-weekly", async () => {
  console.log("run-weekly called - redirecting to manual approach");
  return {
    ok: false,
    msg: "Please use the manual date setting approach instead",
  };
});

// NEW: Download handler for the Excel file
ipcMain.handle("download-file", async (_evt, filePath) => {
  try {
    if (!fs.existsSync(filePath)) {
      return { ok: false, msg: "File not found" };
    }

    // For Electron, we'll use the shell to open the file location
    const { shell } = require("electron");
    await shell.showItemInFolder(filePath);

    return { ok: true };
  } catch (err) {
    console.error("download-file failed:", err);
    return { ok: false, msg: String(err.message || err) };
  }
});

// PDF generation handler with proper headers and zero values
ipcMain.handle("generate-pdf", async (_evt, excelFilePath) => {
  try {
    if (!fs.existsSync(excelFilePath)) {
      return { ok: false, msg: "Excel file not found" };
    }

    // Generate PDF filename
    const pdfFilePath = excelFilePath.replace(".xlsx", ".pdf");

    // Read Excel file
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(excelFilePath);
    const ws = wb.getWorksheet("Weekly Earnings Report");

    // Build HTML table from Excel data
    let htmlContent = `
        <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; font-size: 12px; }
            .banner { font-size: 16px; font-weight: bold; margin-bottom: 20px; color: #333; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; }
            th, td { border: 1px solid #ddd; padding: 6px; text-align: left; font-size: 10px; }
            th { background-color: #f2f2f2; font-weight: bold; text-align: center; }
            .totals { background-color: #d3d3d3; font-weight: bold; }
            .payout { background-color: #ffff99; }
            .alternate { background-color: #f8f8f8; }
          </style>
        </head>
        <body>
      `;

    // Add banner (first row)
    const bannerCell = ws.getCell(1, 1);
    htmlContent += `<div class="banner">${
      bannerCell.value || "Weekly Earnings Report"
    }</div>`;

    // Define headers manually (since ws.columns might not have proper headers)
    const headers = [
      "Driver",
      "Total Earnings",
      "Fare",
      "Service Fee",
      "Other Earnings",
      "Taxes",
      "Tips",
      "Refunds & Expenses",
      "Adjustments",
      "Payout",
      "Net Earnings",
      "Trips",
      "Distance (km)",
    ];

    // Add table
    htmlContent += "<table>";

    // Add headers
    htmlContent += "<tr>";
    headers.forEach((header) => {
      htmlContent += `<th>${header}</th>`;
    });
    htmlContent += "</tr>";

    // Add data rows (starting from row 3, skip banner and header)
    let rowIndex = 0;
    ws.eachRow((row, rowNumber) => {
      if (rowNumber <= 2) return; // Skip banner and header rows

      const isTotal = String(row.getCell(1).value || "").includes("TOTAL");
      const isAlternate = rowIndex % 2 === 1 && !isTotal;
      const rowClass = isTotal ? "totals" : isAlternate ? "alternate" : "";

      htmlContent += `<tr class="${rowClass}">`;

      // Process each column
      for (let colIndex = 1; colIndex <= headers.length; colIndex++) {
        const cell = row.getCell(colIndex);
        const isPayout = colIndex === 10;
        const cellClass = isPayout && !isTotal ? "payout" : "";

        let value = cell.value;

        // Handle null/undefined values - show 0 for numeric columns, empty for text
        if (value === null || value === undefined || value === "") {
          if (colIndex === 1) {
            // Driver name column
            value = "";
          } else {
            value = 0;
          }
        }

        // Format numbers properly
        if (typeof value === "number" && colIndex > 1) {
          if (colIndex === 12) {
            // Distance column
            value = value.toFixed(2);
          } else if (colIndex === 11) {
            // Trips column
            value = Math.round(value);
          } else {
            // All money columns
            value = value.toLocaleString("en-US", {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            });
          }
        }

        htmlContent += `<td class="${cellClass}">${value}</td>`;
      }

      htmlContent += "</tr>";
      if (!isTotal) rowIndex++;
    });

    htmlContent += "</table></body></html>";

    // Use puppeteer to generate PDF
    const puppeteer = require("puppeteer");
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    await page.setContent(htmlContent);
    await page.pdf({
      path: pdfFilePath,
      format: "A3",
      landscape: true,
      margin: { top: "15px", right: "15px", bottom: "15px", left: "15px" },
    });

    await browser.close();

    return { ok: true, file: pdfFilePath };
  } catch (err) {
    console.error("generate-pdf failed:", err);
    return { ok: false, msg: String(err.message || err) };
  }
});
