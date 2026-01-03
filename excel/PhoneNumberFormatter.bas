Attribute VB_Name = "PhoneNumberFormatter"
Option Explicit

' Output styles:
'   "digits" -> 4155550123
'   "dash"   -> 415-555-0123
'   "paren"  -> (415) 555-0123
'   "e164"   -> +14155550123
Private Const DEFAULT_OUTPUT_STYLE As String = "digits"
Private Const INVALID_TEXT As String = "INVALID"

Private Const DEFAULT_HEADER_ROW As Long = 1
Private Const DEFAULT_FIRST_DATA_ROW As Long = 2

' Entry point: formats phone-number columns on the active sheet.
Public Sub FormatPhoneNumbersInActiveSheet()
    FormatPhoneNumbersInWorksheet ActiveSheet, DEFAULT_OUTPUT_STYLE, DEFAULT_HEADER_ROW, DEFAULT_FIRST_DATA_ROW
End Sub

' Entry point: formats phone-number columns on every worksheet in the active workbook.
Public Sub FormatPhoneNumbersInWorkbook()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        FormatPhoneNumbersInWorksheet ws, DEFAULT_OUTPUT_STYLE, DEFAULT_HEADER_ROW, DEFAULT_FIRST_DATA_ROW
    Next ws
End Sub

' Formats any column whose header looks like a phone-number column.
' - Looks for headers containing: phone, mobile, cell, sms, tel/telephone
' - Rewrites values in-place, converting vanity letters to digits.
' - Marks invalid inputs as "INVALID".
Public Sub FormatPhoneNumbersInWorksheet(ByVal ws As Worksheet, _
                                        Optional ByVal outputStyle As String = DEFAULT_OUTPUT_STYLE, _
                                        Optional ByVal headerRow As Long = DEFAULT_HEADER_ROW, _
                                        Optional ByVal firstDataRow As Long = DEFAULT_FIRST_DATA_ROW)
    Dim phoneCols As Collection
    Set phoneCols = FindPhoneNumberColumns(ws, headerRow)

    If phoneCols.Count = 0 Then
        MsgBox "No phone-number columns found on sheet '" & ws.Name & "' (searched header row " & headerRow & ").", vbInformation
        Exit Sub
    End If

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation

    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim col As Variant
    For Each col In phoneCols
        CleanPhoneColumn ws, CLng(col), firstDataRow, outputStyle
    Next col

CleanUp:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Error while formatting phone numbers: " & Err.Description, vbExclamation
    End If
End Sub

' Quick sanity check (open VBA editor, run this sub, view output in Immediate Window: Ctrl+G)
Public Sub PhoneFormatter_SelfTest()
    Debug.Print NormalizePhoneUSCA("(415) 555-0123")
    Debug.Print NormalizePhoneUSCA("1-800-FLOWERS")
    Debug.Print NormalizePhoneUSCA("415.555.0123 x89")
    Debug.Print NormalizePhoneUSCA("INVALID INPUT")
End Sub

Private Sub CleanPhoneColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal firstDataRow As Long, ByVal outputStyle As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    If lastRow < firstDataRow Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastRow, col))

    ' Force text so Excel doesn't mangle large numbers.
    rng.NumberFormat = "@"

    Dim v As Variant
    v = rng.Value2

    Dim r As Long
    If IsArray(v) Then
        For r = 1 To UBound(v, 1)
            v(r, 1) = NormalizePhoneUSCA(v(r, 1), outputStyle)
        Next r
        rng.Value2 = v
    Else
        rng.Value2 = NormalizePhoneUSCA(v, outputStyle)
    End If
End Sub

Private Function FindPhoneNumberColumns(ByVal ws As Worksheet, ByVal headerRow As Long) As Collection
    Dim cols As New Collection

    Dim used As Range
    On Error Resume Next
    Set used = ws.UsedRange
    On Error GoTo 0

    If used Is Nothing Then
        Set FindPhoneNumberColumns = cols
        Exit Function
    End If

    Dim firstCol As Long
    firstCol = used.Column

    Dim lastCol As Long
    lastCol = used.Column + used.Columns.Count - 1

    Dim c As Long
    For c = firstCol To lastCol
        Dim header As String
        header = CStr(ws.Cells(headerRow, c).Value2)
        If IsPhoneHeader(header) Then
            cols.Add c
        End If
    Next c

    Set FindPhoneNumberColumns = cols
End Function

Private Function IsPhoneHeader(ByVal header As String) As Boolean
    Dim h As String
    h = LCase$(Trim$(header))

    If h = vbNullString Then Exit Function

    If InStr(1, h, "phone", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
    If InStr(1, h, "mobile", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
    If InStr(1, h, "cell", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
    If InStr(1, h, "sms", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
    If InStr(1, h, "telephone", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
    If InStr(1, h, "tel", vbTextCompare) > 0 Then IsPhoneHeader = True: Exit Function
End Function

Private Function NormalizePhoneUSCA(ByVal raw As Variant, Optional ByVal outputStyle As String = DEFAULT_OUTPUT_STYLE) As String
    If IsError(raw) Then
        NormalizePhoneUSCA = INVALID_TEXT
        Exit Function
    End If

    Dim s As String
    If IsNumeric(raw) Then
        ' Prevent scientific notation for large numbers.
        s = Format$(raw, "0")
    Else
        s = CStr(raw)
    End If

    s = Trim$(s)
    If s = vbNullString Then
        NormalizePhoneUSCA = vbNullString
        Exit Function
    End If

    Dim digits As String
    digits = AlphaNumericToDigitsUntilExtension(s)

    Dim tenDigits As String
    tenDigits = ExtractNANPTenDigits(digits)

    If tenDigits = vbNullString Then
        NormalizePhoneUSCA = INVALID_TEXT
        Exit Function
    End If

    NormalizePhoneUSCA = ApplyOutputStyle(tenDigits, outputStyle)
End Function

' Extracts a US/Canada (NANP) 10-digit number from a digit string.
' Accepts either:
'   - 10 digits
'   - 11 digits starting with "1" (drops leading 1)
Private Function ExtractNANPTenDigits(ByVal digits As String) As String
    digits = Trim$(digits)

    If Len(digits) = 10 Then
        ExtractNANPTenDigits = digits
        Exit Function
    End If

    If Len(digits) = 11 And Left$(digits, 1) = "1" Then
        ExtractNANPTenDigits = Mid$(digits, 2)
        Exit Function
    End If

    ExtractNANPTenDigits = vbNullString
End Function

Private Function ApplyOutputStyle(ByVal tenDigits As String, ByVal outputStyle As String) As String
    Dim style As String
    style = LCase$(Trim$(outputStyle))

    Select Case style
        Case "digits"
            ApplyOutputStyle = tenDigits
        Case "dash"
            ApplyOutputStyle = Left$(tenDigits, 3) & "-" & Mid$(tenDigits, 4, 3) & "-" & Right$(tenDigits, 4)
        Case "paren"
            ApplyOutputStyle = "(" & Left$(tenDigits, 3) & ") " & Mid$(tenDigits, 4, 3) & "-" & Right$(tenDigits, 4)
        Case "e164"
            ApplyOutputStyle = "+1" & tenDigits
        Case Else
            ApplyOutputStyle = tenDigits
    End Select
End Function

' Converts a phone-ish string into digits by:
' - keeping digits
' - converting A-Z to phone keypad digits
' - ignoring everything else
' - stopping before an extension (e.g. "ext 123", "x89") once we've already collected 10+ digits
Private Function AlphaNumericToDigitsUntilExtension(ByVal s As String) As String
    Dim out As String
    out = vbNullString

    Dim i As Long
    Dim n As Long
    n = Len(s)

    For i = 1 To n
        ' Extension stop (only after we already have a full base number)
        If Len(out) >= 10 Then
            If IsExtensionMarkerAt(s, i) And RemainingHasDigit(s, i + 1) Then
                Exit For
            End If
        End If

        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            out = out & ch
        Else
            Dim up As String
            up = UCase$(ch)
            If up >= "A" And up <= "Z" Then
                out = out & AlphaToKeypadDigit(up)
            End If
        End If
    Next i

    AlphaNumericToDigitsUntilExtension = out
End Function

' Returns True if s contains an extension marker starting at index i.
' Recognizes "x" and "ext"/"extension" (case-insensitive).
Private Function IsExtensionMarkerAt(ByVal s As String, ByVal i As Long) As Boolean
    Dim rem As String
    rem = LCase$(Mid$(s, i))

    If Left$(rem, 1) = "x" Then
        IsExtensionMarkerAt = True
        Exit Function
    End If

    If Left$(rem, 3) = "ext" Then
        IsExtensionMarkerAt = True
        Exit Function
    End If

    If Left$(rem, 9) = "extension" Then
        IsExtensionMarkerAt = True
        Exit Function
    End If
End Function

Private Function RemainingHasDigit(ByVal s As String, ByVal startIndex As Long) As Boolean
    Dim i As Long
    For i = startIndex To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            RemainingHasDigit = True
            Exit Function
        End If
    Next i
End Function

' Standard phone keypad mapping.
Private Function AlphaToKeypadDigit(ByVal up As String) As String
    Select Case up
        Case "A", "B", "C": AlphaToKeypadDigit = "2"
        Case "D", "E", "F": AlphaToKeypadDigit = "3"
        Case "G", "H", "I": AlphaToKeypadDigit = "4"
        Case "J", "K", "L": AlphaToKeypadDigit = "5"
        Case "M", "N", "O": AlphaToKeypadDigit = "6"
        Case "P", "Q", "R", "S": AlphaToKeypadDigit = "7"
        Case "T", "U", "V": AlphaToKeypadDigit = "8"
        Case "W", "X", "Y", "Z": AlphaToKeypadDigit = "9"
        Case Else: AlphaToKeypadDigit = ""
    End Select
End Function
