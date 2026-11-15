Attribute VB_Name = "PromoMacros"
Option Explicit
'==============================================================================
' MM Promo Tool - Add / Update / Load promotion macros (Input2 <-> Database)
'
' SETUP (one time):  save the workbook as .xlsm, then in the VBA editor use
' File > Import File... and choose this PromoMacros.bas.
' The three buttons on Input2 are already assigned to AddPromotion,
' UpdatePromotion and LoadSelectedPromotion.
'==============================================================================

Private Const SH_INPUT As String = "Input2"
Private Const SH_DB As String = "Database"
Private Const DB_FIRST_ROW As Long = 6        'first data row on Database
Private Const DB_LAST_ROW As Long = 1000      'last row the Input2 formulas read
Private Const DB_FIRST_COL As Long = 3        'column C
Private Const DB_LAST_COL As Long = 64        'column BL
Private Const LIST_FIRST_ROW As Long = 30     'bottom list on Input2 (B30:C129)
Private Const LIST_LAST_ROW As Long = 129
Private Const CONFIRM_ON_ALERTS As Boolean = True 'ask before saving when alerts show

'------------------------------------------------------------------------------
' Buttons
'------------------------------------------------------------------------------
Public Sub AddPromotion()
    Dim wsI As Worksheet, wsD As Worksheet
    Dim msg As String, code As String, serial As Double
    Dim r As Long, rec As Variant
    On Error GoTo EH
    Set wsI = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsD = ThisWorkbook.Worksheets(SH_DB)

    msg = ValidateForm(wsI)
    If Len(msg) > 0 Then
        MsgBox "Promotion NOT added. Please fix:" & vbLf & vbLf & msg, vbExclamation, "Add Promotion"
        Exit Sub
    End If
    serial = CDbl(wsI.Range("K5").Value2)
    code = Txt(wsI.Range("D8").Value2)
    If IsLockedMonth(serial) Then
        MsgBox "The selected month is the current month or in the past, so changes cannot be saved.", vbExclamation, "Add Promotion"
        Exit Sub
    End If
    If FindDbRow(wsD, code, serial) > 0 Then
        MsgBox code & " already has a promotion in " & Format$(CDate(serial), "mmm yyyy") & ". Use Update Promotion to change it.", vbExclamation, "Add Promotion"
        Exit Sub
    End If
    If Not ConfirmAlerts(wsI, "Add Promotion") Then Exit Sub

    r = NextDbRow(wsD)
    If r > DB_LAST_ROW Then
        MsgBox "The Database is full (row " & DB_LAST_ROW & ").", vbExclamation, "Add Promotion"
        Exit Sub
    End If
    rec = BuildRecord(wsI)
    Application.ScreenUpdating = False
    If r > DB_FIRST_ROW Then CopyRowFormats wsD, r
    wsD.Range(wsD.Cells(r, DB_FIRST_COL), wsD.Cells(r, DB_LAST_COL)).Value = rec
    Application.ScreenUpdating = True
    MsgBox "Promotion added: " & code & " - " & Format$(CDate(serial), "mmm yyyy") & " (Database row " & r & ").", vbInformation, "Add Promotion"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Add Promotion failed: " & Err.Description, vbCritical, "Add Promotion"
End Sub

Public Sub UpdatePromotion()
    Dim wsI As Worksheet, wsD As Worksheet
    Dim msg As String, code As String, serial As Double
    Dim r As Long, rec As Variant
    On Error GoTo EH
    Set wsI = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsD = ThisWorkbook.Worksheets(SH_DB)

    msg = ValidateForm(wsI)
    If Len(msg) > 0 Then
        MsgBox "Promotion NOT updated. Please fix:" & vbLf & vbLf & msg, vbExclamation, "Update Promotion"
        Exit Sub
    End If
    serial = CDbl(wsI.Range("K5").Value2)
    code = Txt(wsI.Range("D8").Value2)
    If IsLockedMonth(serial) Then
        MsgBox "The selected month is the current month or in the past, so changes cannot be saved.", vbExclamation, "Update Promotion"
        Exit Sub
    End If
    r = FindDbRow(wsD, code, serial)
    If r = 0 Then
        MsgBox "No existing promotion for " & code & " in " & Format$(CDate(serial), "mmm yyyy") & ". Use Add Promotion instead.", vbExclamation, "Update Promotion"
        Exit Sub
    End If
    If Not ConfirmAlerts(wsI, "Update Promotion") Then Exit Sub
    If MsgBox("Overwrite the existing " & code & " promotion for " & Format$(CDate(serial), "mmm yyyy") & " with the data on the form?", vbYesNo + vbQuestion, "Update Promotion") <> vbYes Then Exit Sub

    rec = BuildRecord(wsI)
    Application.ScreenUpdating = False
    wsD.Range(wsD.Cells(r, DB_FIRST_COL), wsD.Cells(r, DB_LAST_COL)).Value = rec
    Application.ScreenUpdating = True
    MsgBox "Promotion updated: " & code & " - " & Format$(CDate(serial), "mmm yyyy") & " (Database row " & r & ").", vbInformation, "Update Promotion"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Update Promotion failed: " & Err.Description, vbCritical, "Update Promotion"
End Sub

Public Sub LoadSelectedPromotion()
    Dim wsI As Worksheet, wsD As Worksheet
    Dim c As Range, rowNo As Long, idx As Variant, dbRow As Long
    Dim entries() As String, parts() As String, i As Long
    Dim v As Variant, code As String, lk As Variant, target As Range, ph As String
    On Error GoTo EH
    Set wsI = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsD = ThisWorkbook.Worksheets(SH_DB)
    If ActiveSheet.Name <> SH_INPUT Then
        MsgBox "Go to the Input2 sheet, click a Product code in the list, then press Load Selected Promotion.", vbInformation, "Load Selected Promotion"
        Exit Sub
    End If
    Set c = ActiveCell
    If Intersect(c, wsI.Range("B" & LIST_FIRST_ROW & ":C" & LIST_LAST_ROW)) Is Nothing Then
        MsgBox "Click a Product code in the 'Promotions in Selected Month' list first.", vbInformation, "Load Selected Promotion"
        Exit Sub
    End If
    rowNo = c.Row
    idx = wsI.Cells(rowNo, 23).Value2                  'column W = Database row index
    If Not IsNum(idx) Then
        MsgBox "There is no promotion on that row.", vbInformation, "Load Selected Promotion"
        Exit Sub
    End If
    dbRow = DB_FIRST_ROW - 1 + CLng(idx)

    Application.ScreenUpdating = False
    'product
    lk = wsD.Cells(dbRow, 4).Value2                    'D = Product Lookup
    code = Txt(wsD.Cells(dbRow, 5).Value2)             'E = Product Code
    If IsBlank(lk) Then lk = ProdMasterLookup(code)
    If Not IsBlank(lk) Then wsI.Range("D7").Value2 = lk
    'inputs
    entries = MapEntries()
    For i = LBound(entries) To UBound(entries)
        parts = Split(entries(i), "|")
        If UBound(parts) >= 2 Then
            If InStr(parts(2), "I") > 0 Then
                v = wsD.Cells(dbRow, ColNum(parts(0))).Value2
                Set target = wsI.Range(parts(1))
                If IsBlank(v) Then
                    ph = PlaceholderFor(parts(1))
                    If Len(ph) > 0 Then target.Value2 = ph Else target.ClearContents
                Else
                    target.Value2 = v
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Loaded " & code & " - " & Format$(wsD.Cells(dbRow, DB_FIRST_COL).Value2, "mmm yyyy") & " into the form.", vbInformation, "Load Selected Promotion"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Load Selected Promotion failed: " & Err.Description, vbCritical, "Load Selected Promotion"
End Sub

'------------------------------------------------------------------------------
' Form <-> Database mapping.  Each entry is  DatabaseColumn|InputCell|flags
'   I = input cell on the form (copied back by Load Selected Promotion)
'   N = numeric only (text placeholders such as "$Y", "Y%" or "X1" are ignored)
'   P = ignore the "(placeholder)" text
' Database columns not listed here: F (Product Name), BK and BL (totals).
'------------------------------------------------------------------------------
Private Function MapEntries() As String()
    Dim s As String
    s = "C|K5|,D|D7|,E|D8|,G|D11|,H|D12|,I|D13|,J|D14|,K|D15|,L|D16|,M|D18|,N|D19|,O|D20|,P|K7|,Q|K8|P,"
    s = s & "R|S11|I,S|S12|I,T|S13|I,U|S15|I,V|S16|I,W|S17|I,X|S18|I,Y|S20|I,Z|T20|I,AA|U20|I,"
    s = s & "AB|S21|I,AC|T21|I,AD|U21|I,AE|E25|I,AF|E24|I,AG|E26|,"
    s = s & "AH|I24|IN,AI|I25|IN,AJ|I26|IN,AK|L24|IN,AL|L25|IN,AM|L26|IN,"
    s = s & "AN|P24|IN,AO|P25|IN,AP|P26|IN,AQ|S24|IN,AR|S25|IN,AS|S26|IN,"
    s = s & "AT|L11|,AU|L12|,AV|L13|,AW|L14|,AX|M11|,AY|M12|,AZ|M13|,BA|M14|,"
    s = s & "BB|N11|,BC|N12|,BD|N13|,BE|N14|,BF|L15|,BG|M15|,BH|M16|,BI|M17|,BJ|M18"
    MapEntries = Split(s, ",")
End Function

Private Function PlaceholderFor(ByVal addr As String) As String
    Select Case addr
        Case "L24", "L25", "L26": PlaceholderFor = "$Y"
        Case "S24", "S25", "S26": PlaceholderFor = "Y%"
        Case "I24", "P24": PlaceholderFor = "X1"
        Case "I25", "P25": PlaceholderFor = "X2"
        Case "I26", "P26": PlaceholderFor = "X3"
        Case Else: PlaceholderFor = ""
    End Select
End Function

Private Function BuildRecord(ByVal ws As Worksheet) As Variant
    Dim rec(1 To 1, 1 To DB_LAST_COL - DB_FIRST_COL + 1) As Variant
    Dim entries() As String, parts() As String, i As Long, flags As String
    Dim code As String, nm As Variant
    entries = MapEntries()
    For i = LBound(entries) To UBound(entries)
        parts = Split(entries(i), "|")
        flags = ""
        If UBound(parts) >= 2 Then flags = parts(2)
        rec(1, ColNum(parts(0)) - DB_FIRST_COL + 1) = CleanValue(ws.Range(parts(1)).Value2, flags)
    Next i
    rec(1, 1) = CDate(ws.Range("K5").Value2)          'C  promo month as a date
    code = Txt(ws.Range("D8").Value2)
    nm = ProdMasterValue(code, "Description")         'F  product name
    If Not IsEmpty(nm) Then rec(1, ColNum("F") - DB_FIRST_COL + 1) = nm
    If IsNum(ws.Range("M15").Value2) Or IsNum(ws.Range("M17").Value2) Then _
        rec(1, ColNum("BK") - DB_FIRST_COL + 1) = NumOr0(ws.Range("M15").Value2) + NumOr0(ws.Range("M17").Value2)
    If IsNum(ws.Range("M16").Value2) Or IsNum(ws.Range("M18").Value2) Then _
        rec(1, ColNum("BL") - DB_FIRST_COL + 1) = NumOr0(ws.Range("M16").Value2) + NumOr0(ws.Range("M18").Value2)
    BuildRecord = rec
End Function

'------------------------------------------------------------------------------
' Validation
'------------------------------------------------------------------------------
Private Function ValidateForm(ByVal ws As Worksheet) As String
    Dim e As String, t As String, h As String
    Dim up As Variant, pq As Variant, ok As Boolean, c As Range, bad As Boolean

    If Not IsNum(ws.Range("K5").Value2) Then e = e & "- Month (K5) must be selected" & vbLf
    If IsBlank(ws.Range("D8").Value2) Then e = e & "- Select a valid SKU / Vendor Lookup" & vbLf

    t = LCase$(Txt(ws.Range("S11").Value2))
    h = LCase$(Txt(ws.Range("S13").Value2))
    If t = "" Then e = e & "- Promotional Type must be filled" & vbLf
    If Txt(ws.Range("S12").Value2) = "" Then e = e & "- Promotional Intent must be filled" & vbLf
    If h = "" Then e = e & "- Hotspot must be filled" & vbLf

    'Uplift % or Promo Qty - exactly one
    up = ws.Range("S16").Value2
    pq = ws.Range("S17").Value2
    If Not IsBlank(up) And Not IsBlank(pq) Then
        e = e & "- Uplift % and Promo Qty are both filled - clear one of them" & vbLf
    ElseIf IsBlank(up) And IsBlank(pq) Then
        e = e & "- Either Uplift % or Promo Qty must be filled" & vbLf
    ElseIf (Not IsBlank(up) And Not IsNum(up)) Or (Not IsBlank(pq) And Not IsNum(pq)) Then
        e = e & "- Uplift % / Promo Qty must be a number" & vbLf
    End If

    'Presentation qty unless Bargain Bin / Shelf only
    If h <> "" And h <> "bargain bin" And h <> "shelf only" Then
        If Not IsNum(ws.Range("S20").Value2) Or Not IsNum(ws.Range("T20").Value2) Or Not IsNum(ws.Range("U20").Value2) Then
            e = e & "- Presentation Qty (Small, Medium and Large) must be filled for this Hotspot" & vbLf
        End If
    End If

    'Promotion-type specific inputs
    If t = "discount" Then
        ok = IsPositive(ws.Range("E24").Value2) Or IsPositive(ws.Range("E25").Value2)
        If Not ok Then e = e & "- Discount: enter Discount ($) or % Off" & vbLf
        If IsPositive(ws.Range("E25").Value2) Then
            If CDbl(ws.Range("E25").Value2) > 1 Then e = e & "- % Off must be between 0% and 100% (e.g. 10%)" & vbLf
        End If
    ElseIf Left$(t, 5) = "buy 2" Then
        ok = False
        For Each c In ws.Range("L24:L26").Cells
            If IsPositive(c.Value2) Then ok = True
        Next c
        If Not ok Then e = e & "- Buy 2 or more for $XX: fill at least one $Y row in 'Promo Input - Multi-buy x for y'" & vbLf
    ElseIf Left$(t, 5) = "buy x" Then
        ok = False: bad = False
        For Each c In ws.Range("S24:S26").Cells
            If IsPositive(c.Value2) Then
                ok = True
                If CDbl(c.Value2) > 1 Then bad = True
            End If
        Next c
        If Not ok Then e = e & "- Buy X get XX% discount: fill at least one Y% row in 'Promo Input - Multi-buy x and y%'" & vbLf
        If bad Then e = e & "- Y% must be between 0% and 100% (e.g. 25%)" & vbLf
    End If
    ValidateForm = e
End Function

Private Function ConfirmAlerts(ByVal ws As Worksheet, ByVal title As String) As Boolean
    Dim r As Long, t As String, s As String
    ConfirmAlerts = True
    If Not CONFIRM_ON_ALERTS Then Exit Function
    For r = 5 To 8
        t = Txt(ws.Cells(r, 17).Value2)                'Alerts box lines (Q5:Q8)
        If Len(t) > 0 And Left$(t, 2) <> "OK" Then s = s & "- " & t & vbLf
    Next r
    If Len(s) > 0 Then
        ConfirmAlerts = (MsgBox("This promotion has pre-promotion alerts:" & vbLf & vbLf & s & vbLf & "Save it anyway?", vbYesNo + vbExclamation, title) = vbYes)
    End If
End Function

Private Function IsLockedMonth(ByVal serial As Double) As Boolean
    Dim d As Date
    d = CDate(serial)
    IsLockedMonth = (Year(d) * 12 + Month(d) <= Year(Date) * 12 + Month(Date))
End Function

'------------------------------------------------------------------------------
' Database helpers
'------------------------------------------------------------------------------
Private Function LastDbRow(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long, m As Long
    For c = 3 To 5                                     'columns C:E
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > m Then m = r
    Next c
    LastDbRow = m
End Function

Private Function NextDbRow(ByVal ws As Worksheet) As Long
    NextDbRow = LastDbRow(ws) + 1
    If NextDbRow < DB_FIRST_ROW Then NextDbRow = DB_FIRST_ROW
End Function

Private Function FindDbRow(ByVal ws As Worksheet, ByVal code As String, ByVal serial As Double) As Long
    Dim last As Long, v As Variant, r As Long
    last = LastDbRow(ws)
    If last < DB_FIRST_ROW Then Exit Function
    v = ws.Range(ws.Cells(DB_FIRST_ROW, 3), ws.Cells(last, 5)).Value2     'C:E
    For r = 1 To UBound(v, 1)
        If IsNum(v(r, 1)) Then
            If MonthKey(CDbl(v(r, 1))) = MonthKey(serial) Then
                If StrComp(Txt(v(r, 3)), code, vbTextCompare) = 0 Then
                    FindDbRow = DB_FIRST_ROW + r - 1
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function MonthKey(ByVal serial As Double) As Long
    MonthKey = Year(CDate(serial)) * 12 + Month(CDate(serial))
End Function

Private Sub CopyRowFormats(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    ws.Range(ws.Cells(DB_FIRST_ROW, DB_FIRST_COL), ws.Cells(DB_FIRST_ROW, DB_LAST_COL)).Copy
    ws.Range(ws.Cells(r, DB_FIRST_COL), ws.Cells(r, DB_LAST_COL)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Function ProdMasterValue(ByVal code As String, ByVal returnCol As String) As Variant
    Dim lo As ListObject, m As Variant
    On Error GoTo NF
    Set lo = ThisWorkbook.Worksheets("Product Master").ListObjects("ProdMaster")
    m = Application.Match(code, lo.ListColumns("Product Code2").DataBodyRange, 0)
    If IsError(m) Then GoTo NF
    ProdMasterValue = lo.ListColumns(returnCol).DataBodyRange.Cells(CLng(m), 1).Value2
    Exit Function
NF:
    ProdMasterValue = Empty
End Function

Private Function ProdMasterLookup(ByVal code As String) As Variant
    ProdMasterLookup = ProdMasterValue(code, "Lookup")
End Function

'------------------------------------------------------------------------------
' Small value helpers
'------------------------------------------------------------------------------
Private Function ColNum(ByVal letters As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(letters)
        n = n * 26 + (Asc(UCase$(Mid$(letters, i, 1))) - 64)
    Next i
    ColNum = n
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlank = True
    ElseIf IsEmpty(v) Then
        IsBlank = True
    Else
        IsBlank = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        IsNum = False
    ElseIf VarType(v) = vbString Then
        IsNum = False
    Else
        IsNum = IsNumeric(v)
    End If
End Function

Private Function IsPositive(ByVal v As Variant) As Boolean
    If IsNum(v) Then IsPositive = (CDbl(v) > 0)
End Function

Private Function NumOr0(ByVal v As Variant) As Double
    If IsNum(v) Then NumOr0 = CDbl(v)
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        Txt = ""
    Else
        Txt = Trim$(CStr(v))
    End If
End Function

Private Function CleanValue(ByVal v As Variant, ByVal flags As String) As Variant
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then
        CleanValue = Empty
    ElseIf VarType(v) = vbString Then
        s = Trim$(v)
        If Len(s) = 0 Then
            CleanValue = Empty
        ElseIf InStr(flags, "N") > 0 Then
            CleanValue = Empty
        ElseIf InStr(flags, "P") > 0 And Left$(s, 1) = "(" Then
            CleanValue = Empty
        Else
            CleanValue = s
        End If
    Else
        CleanValue = v
    End If
End Function
