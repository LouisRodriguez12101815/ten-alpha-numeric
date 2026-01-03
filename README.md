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
