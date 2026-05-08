/**
 * Fetches the SiteList directly from the shared Google Sheet.
 * The sheet is exported as CSV and parsed in the browser.
 * 
 * Spreadsheet: Alex Database-Shenoz.xlsx
 * Sheet: SiteList
 * Columns: A=Short Code, B=Site Name, C=SC Office, D=OZ (Region)
 */

const SPREADSHEET_ID = '1XwQTSMI5Nz0WuKwSow06dVmTlnoyZOuj';
const SHEET_NAME = 'SiteList';
const CSV_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${SHEET_NAME}`;

/**
 * Parse a CSV string handling quoted fields with commas and newlines.
 */
function parseCSV(text) {
  const rows = [];
  let current = '';
  let inQuotes = false;
  const chars = text.split('');

  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (ch === '"') {
      if (inQuotes && chars[i + 1] === '"') {
        current += '"';
        i++; // skip escaped quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      if (!rows.length || rows[rows.length - 1] === undefined) {
        rows.push([]);
      }
      rows[rows.length - 1].push(current);
      current = '';
    } else if (ch === '\n' && !inQuotes) {
      if (!rows.length || rows[rows.length - 1] === undefined) {
        rows.push([]);
      }
      rows[rows.length - 1].push(current);
      current = '';
      rows.push(undefined); // marker for new row
    } else if (ch === '\r' && !inQuotes) {
      // skip carriage return
    } else {
      current += ch;
    }
  }

  // Push last field
  if (current || (rows.length && rows[rows.length - 1] === undefined)) {
    if (!rows.length || rows[rows.length - 1] === undefined) {
      rows.push([]);
    }
    rows[rows.length - 1].push(current);
  }

  return rows.filter(r => r !== undefined && r.length > 0);
}

/**
 * Fetches and parses the SiteList from Google Sheets.
 * @param {Object} config - Configuration object containing spreadsheet info and column mapping
 */
export async function fetchSiteList(config) {
  const spreadsheetId = config?.spreadsheetId || '1XwQTSMI5Nz0WuKwSow06dVmTlnoyZOuj';
  const sheetName = config?.sheetName || 'SiteList';
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${sheetName}`;

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch SiteList: ${response.status} ${response.statusText}`);
  }

  const csvText = await response.text();
  const rows = parseCSV(csvText);

  if (rows.length < 2) {
    throw new Error('SiteList sheet appears to be empty.');
  }

  const dataRows = rows.slice(1);

  // Mapping using config indices (0-based)
  const map = config?.mapping || {
    shortCode: 0,
    siteName: 1,
    scOffice: 2,
    region: 3,
    nodalDeg: 4,
    pld: 18,
    batteries: 19,
    bdt: 27,
    vip: 33
  };

  const sites = dataRows
    .map(row => {
      const shortCode = (row[map.shortCode] || '').trim();
      if (!shortCode) return null;

      return {
        shortCode,
        siteName: (row[map.siteName] || '').trim(),
        scOffice: (row[map.scOffice] || '').trim(),
        region: (row[map.region] || '').trim(),
        nodalDeg: (row[map.nodalDeg] || '').trim(),
        pld: (row[map.pld] || '').trim().toLowerCase() === 'true' || (row[map.pld] || '').trim().toLowerCase() === 'yes', 
        batteries: (row[map.batteries] || '').trim(),
        bdt: parseFloat((row[map.bdt] || '').replace(/[^\d.]/g, '')) || 0,
        vip: (row[map.vip] || '').trim().toLowerCase() === 'true' || (row[map.vip] || '').trim().toLowerCase() === 'yes',
        raw: row,
      };
    })
    .filter(Boolean);

  return sites;
}
