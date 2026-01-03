/**
 * Google Sheets (Apps Script) phone number formatter (US/Canada only).
 *
 * Features:
 * - Detects likely phone-number columns by header name (phone/mobile/cell/sms/tel).
 * - Converts vanity letters to digits (e.g. 1-800-FLOWERS -> 8003569377).
 * - Normalizes to a consistent output format.
 * - Writes "INVALID" when it can't produce a valid 10-digit NANP number.
 */

'use strict';

const DEFAULT_OUTPUT_STYLE = 'digits';
const INVALID_TEXT = 'INVALID';

const DEFAULT_HEADER_ROW = 1;
const DEFAULT_FIRST_DATA_ROW = 2;

/** Adds a menu when the spreadsheet is opened. */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Phone Formatter')
    .addItem('Format phone numbers (active sheet)', 'formatPhoneNumbersInActiveSheet')
    .addItem('Format phone numbers (all sheets)', 'formatPhoneNumbersInSpreadsheet')
    .addToUi();
}

/** Formats phone-number columns on the active sheet. */
function formatPhoneNumbersInActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  formatPhoneNumbersInSheet_(sheet, DEFAULT_OUTPUT_STYLE, DEFAULT_HEADER_ROW, DEFAULT_FIRST_DATA_ROW);
}

/** Formats phone-number columns on every sheet in the active spreadsheet. */
function formatPhoneNumbersInSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach((sheet) => {
    formatPhoneNumbersInSheet_(sheet, DEFAULT_OUTPUT_STYLE, DEFAULT_HEADER_ROW, DEFAULT_FIRST_DATA_ROW);
  });
}

/**
 * Custom function for cell-by-cell use.
 * Example: =NORMALIZE_PHONE_USCA(A2)
 */
function NORMALIZE_PHONE_USCA(input, outputStyle) {
  return normalizePhoneUSCA_(input, outputStyle || DEFAULT_OUTPUT_STYLE);
}

function formatPhoneNumbersInSheet_(sheet, outputStyle, headerRow, firstDataRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < firstDataRow || lastCol < 1) return;

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const phoneCols = [];
  for (let c = 0; c < headers.length; c++) {
    const header = String(headers[c] ?? '');
    if (isPhoneHeader_(header)) phoneCols.push(c + 1);
  }

  if (phoneCols.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No phone-number columns found on sheet '${sheet.getName()}' (searched header row ${headerRow}).`
    );
    return;
  }

  const numRows = lastRow - firstDataRow + 1;

  phoneCols.forEach((col) => {
    const range = sheet.getRange(firstDataRow, col, numRows, 1);
    // Keep as text to avoid formatting surprises.
    range.setNumberFormat('@');

    const values = range.getValues();
    for (let r = 0; r < values.length; r++) {
      values[r][0] = normalizePhoneUSCA_(values[r][0], outputStyle);
    }
    range.setValues(values);
  });
}

function isPhoneHeader_(header) {
  const h = String(header).trim().toLowerCase();
  if (!h) return false;

  return (
    h.includes('phone') ||
    h.includes('mobile') ||
    h.includes('cell') ||
    h.includes('sms') ||
    h.includes('telephone') ||
    h.includes('tel')
  );
}

function normalizePhoneUSCA_(raw, outputStyle) {
  if (raw === null || raw === undefined) return '';

  let s;
  if (typeof raw === 'number' && isFinite(raw)) {
    s = String(Math.trunc(raw));
  } else {
    s = String(raw);
  }

  s = s.trim();
  if (!s) return '';

  const digits = alphaNumericToDigitsUntilExtension_(s);
  const tenDigits = extractNANPTenDigits_(digits);

  if (!tenDigits) return INVALID_TEXT;

  return applyOutputStyle_(tenDigits, outputStyle);
}

/**
 * Extracts a US/Canada (NANP) 10-digit number from a digit string.
 * Accepts either:
 *   - 10 digits
 *   - 11 digits starting with "1" (drops leading 1)
 */
function extractNANPTenDigits_(digits) {
  const d = String(digits || '').trim();
  if (d.length === 10) return d;
  if (d.length === 11 && d.startsWith('1')) return d.slice(1);
  return '';
}

function applyOutputStyle_(tenDigits, outputStyle) {
  const style = String(outputStyle || '').trim().toLowerCase();

  switch (style) {
    case 'digits':
      return tenDigits;
    case 'dash':
      return `${tenDigits.slice(0, 3)}-${tenDigits.slice(3, 6)}-${tenDigits.slice(6)}`;
    case 'paren':
      return `(${tenDigits.slice(0, 3)}) ${tenDigits.slice(3, 6)}-${tenDigits.slice(6)}`;
    case 'e164':
      return `+1${tenDigits}`;
    default:
      return tenDigits;
  }
}

/**
 * Converts a phone-ish string into digits by:
 * - keeping digits
 * - converting A-Z to phone keypad digits
 * - ignoring everything else
 * - stopping before an extension (e.g. "ext 123", "x89") once we've already collected 10+ digits
 */
function alphaNumericToDigitsUntilExtension_(s) {
  let out = '';
  const str = String(s);

  for (let i = 0; i < str.length; i++) {
    // Extension stop (only after we already have a full base number)
    if (out.length >= 10) {
      if (isExtensionMarkerAt_(str, i) && remainingHasDigit_(str, i + 1)) {
        break;
      }
    }

    const ch = str[i];

    if (ch >= '0' && ch <= '9') {
      out += ch;
      continue;
    }

    const up = ch.toUpperCase();
    if (up >= 'A' && up <= 'Z') {
      out += alphaToKeypadDigit_(up);
    }
  }

  return out;
}

/**
 * Returns true if s contains an extension marker starting at index i.
 * Recognizes "x" and "ext"/"extension" (case-insensitive).
 */
function isExtensionMarkerAt_(s, i) {
  const rem = String(s).slice(i).toLowerCase();
  return rem.startsWith('x') || rem.startsWith('ext') || rem.startsWith('extension');
}

function remainingHasDigit_(s, startIndex) {
  const str = String(s);
  for (let i = startIndex; i < str.length; i++) {
    const ch = str[i];
    if (ch >= '0' && ch <= '9') return true;
  }
  return false;
}

/** Standard phone keypad mapping. */
function alphaToKeypadDigit_(up) {
  switch (up) {
    case 'A':
    case 'B':
    case 'C':
      return '2';
    case 'D':
    case 'E':
    case 'F':
      return '3';
    case 'G':
    case 'H':
    case 'I':
      return '4';
    case 'J':
    case 'K':
    case 'L':
      return '5';
    case 'M':
    case 'N':
    case 'O':
      return '6';
    case 'P':
    case 'Q':
    case 'R':
    case 'S':
      return '7';
    case 'T':
    case 'U':
    case 'V':
      return '8';
    case 'W':
    case 'X':
    case 'Y':
    case 'Z':
      return '9';
    default:
      return '';
  }
}
