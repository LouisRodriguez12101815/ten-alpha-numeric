# ten-alpha-numeric
Mobile phone number formatting (US/Canada).

## Excel VBA macro
This repo includes an Excel VBA module that:
- detects likely phone-number columns by header name (phone/mobile/cell/sms/tel)
- converts vanity letters to digits (e.g. `1-800-FLOWERS`)
- normalizes to a consistent output format
- writes `INVALID` when it can't produce a valid US/Canada 10-digit number

### Files
- `excel/PhoneNumberFormatter.bas`

### How to use
1. Open your workbook in Excel.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor: `File` → `Import File...` → choose `excel/PhoneNumberFormatter.bas`.
4. Run one of these macros:
   - `FormatPhoneNumbersInActiveSheet` (active sheet only)
   - `FormatPhoneNumbersInWorkbook` (all sheets)

### Output format
Default output is **10 digits** (example: `4155550123`).
You can change `DEFAULT_OUTPUT_STYLE` in `excel/PhoneNumberFormatter.bas` to:
- `"digits"`
- `"dash"`
- `"paren"`
- `"e164"`

## Google Sheets (Apps Script)
Equivalent Google Sheets code lives here:
- `google-sheets/PhoneNumberFormatter.gs`

### How to use
1. Open your Google Sheet.
2. Go to `Extensions` → `Apps Script`.
3. Create a new script file (or replace `Code.gs`) and paste in the contents of `google-sheets/PhoneNumberFormatter.gs`.
4. Save.
5. Reload the spreadsheet.
6. Use the new menu: `Phone Formatter` →
   - `Format phone numbers (active sheet)` or
   - `Format phone numbers (all sheets)`

### Custom function (optional)
You can also normalize a single cell with:
- `=NORMALIZE_PHONE_USCA(A2)`

### Output format
Default output is **10 digits**.
To change it, set `DEFAULT_OUTPUT_STYLE` in `google-sheets/PhoneNumberFormatter.gs` to:
- `"digits"`
- `"dash"`
- `"paren"`
- `"e164"`
