
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
require('dotenv').config();

const DATA_DIR = path.join(__dirname, 'data');
const HISTORY_FILE = path.join(DATA_DIR, 'sofr_history.json');
const LATEST_FILE = path.join(DATA_DIR, 'sofr_latest.json');
const FORWARD_CSV = path.join(DATA_DIR, 'sofr_forward.csv');
const RAW_XLSX = path.join(DATA_DIR, 'chatham_raw.xlsx');

const MONTH_NAMES = ['January','February','March','April','May','June',
  'July','August','September','October','November','December'];

async function main() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  let snapshot = await downloadChathamExcel();

  if (!snapshot || snapshot.forwardCurve.length === 0) {
    console.error('✗ No forward curve data');
    process.exit(1);
  }

  // Write forward CSV
  let csv = 'Date,Year,SOFR\n';
  for (const pt of snapshot.forwardCurve) {
    csv += `${pt.date},${pt.year},${pt.sofr}\n`;
  }
  fs.writeFileSync(FORWARD_CSV, csv);
// ─── NEW: Write tenor-based forward curve ─────────

let tenorCsv = 'tenor,rate\n';

snapshot.forwardCurve.forEach((pt, index) => {

  let tenor;

  if (index < 12) {
    tenor = `${index + 1}M`;
  } else {
    const years = Math.floor((index + 1) / 12);
    tenor = `${years}Y`;
  }

  tenorCsv += `${tenor},${pt.sofr}\n`;
});

fs.writeFileSync(path.join(DATA_DIR, 'sofr_forward_tenor.csv'), tenorCsv);

console.log('Forward tenor CSV written → data/sofr_forward_tenor.csv');

  // ✅ Write swap CSV
  if (snapshot.swapRates && snapshot.swapRates.length > 0) {
    let swapCsv = 'date,tenor,rate\n';

    for (const s of snapshot.swapRates) {
      const tenor = s.tenor
        .replace(/years?/i, 'Y')
        .replace(/\s+/g, '')
        .toUpperCase();

      const rate = parseFloat(s.rate.replace('%', ''));

      if (!isNaN(rate)) {
        swapCsv += `${snapshot.date},${tenor},${rate}\n`;
      }
    }

    fs.writeFileSync(path.join(DATA_DIR, 'sofr_swaps.csv'), swapCsv);
    console.log(`Swap CSV written → data/sofr_swaps.csv`);
  }

  console.log(`Forward CSV written → data/sofr_forward.csv`);
}

// ─── MAIN SCRAPER ─────────────────────────────────────────────

async function downloadChathamExcel() {
  const { chromium } = require('playwright');

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    console.log('Logging in...');
    await page.goto('https://cf.com/');

    // --- LOGIN (kept simple for clarity, yours still works)
    await page.click('a:has-text("Login")', { timeout: 5000 }).catch(()=>{});
    await page.fill('input[name="username"]', process.env.CHATHAM_EMAIL);
    await page.click('button:has-text("Next")').catch(()=>{});
    await page.fill('input[type="password"]', process.env.CHATHAM_PASSWORD);
    await page.click('button[type="submit"]');

    await page.waitForTimeout(5000);

    console.log('✓ Logged in');

    // ─── STEP 1: GET SWAP RATES ─────────────────

    console.log('Navigating to swaps page...');
    await page.goto('https://cf.com/rates/us/sofr-swaps-annual-annual');
    await page.waitForTimeout(6000);

    const swapRates = await page.evaluate(() => {

      const results = [];

      // get tenors
      const buttons = document.querySelectorAll('button');
      const tenors = [];

      buttons.forEach(btn => {
        const aria = btn.getAttribute('aria-label');
        if (aria && aria.match(/^\d+\s*Year/i)) {
          tenors.push(aria);
        }
      });

      // get all % values
      const text = document.body.innerText;
      const rates = text.match(/\d+\.\d+%/g) || [];

      const n = Math.min(tenors.length, rates.length);

      for (let i = 0; i < n; i++) {
        results.push({
          tenor: tenors[i],
          rate: rates[i]
        });
      }

      return results;
    });

    console.log(`Found ${swapRates.length} swap rows`);

    // ─── STEP 2: FORWARD CURVE ─────────────────

    console.log('Navigating to rates page...');
    await page.goto('https://cf.com/rates/us');
    await page.waitForTimeout(4000);

    const link = await page.$('a[href*="1-month-term-sofr-swaps"]');
    const href = await link.getAttribute('href');

    await page.goto(`https://cf.com${href}`);
    await page.waitForTimeout(4000);

    const downloadBtn = await page.$('button:has-text("Download")');

    const [download] = await Promise.all([
      page.waitForEvent('download'),
      downloadBtn.click()
    ]);

    await download.saveAs(RAW_XLSX);

    const snapshot = parseExcel(RAW_XLSX);

    snapshot.swapRates = swapRates;

    return snapshot;

  } finally {
    await browser.close();
  }
}

// ─── PARSER ─────────────────────────────────────────

function parseExcel(file) {
  const wb = XLSX.readFile(file);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const forwardCurve = [];

  for (let i = 3; i < rows.length; i++) {
    const r = rows[i];
    if (!r[13] || !r[14]) continue;

    const rate = parseFloat(r[14]) * 100;

    const year = 2026 + Math.floor(i / 12);

    forwardCurve.push({
      date: `${MONTH_NAMES[i % 12]}-${String(year).slice(-2)}`,
      year,
      sofr: rate
    });
  }

  return {
    date: new Date().toISOString().split('T')[0],
    forwardCurve
  };
}

main();
