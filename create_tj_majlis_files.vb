Option Explicit

'=========================================================
' STEP 1: Create Region folders + Majlis files
'=========================================================
Sub STEP1_Create_Majlis_Files()

    Dim src As Worksheet
    Set src = ThisWorkbook.Sheets("Tajneed")

    Dim rootPath As String
    rootPath = ThisWorkbook.path & Application.PathSeparator
    
    ' Password for protecting columns
    Dim password As String
    password = "mkad!2026" ' Change this to your desired password

    '---------------------------------------
    ' 1. Build column map
    '---------------------------------------
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, c As Long
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        colMap(LCase(Trim(src.Cells(1, c).Value))) = c
    Next c

    Dim required
    required = Array("region", "majlis", "tanziem", "jamaat id", "name")

    For c = LBound(required) To UBound(required)
        If Not colMap.Exists(required(c)) Then
            MsgBox "Missing required column: " & required(c), vbCritical
            Exit Sub
        End If
    Next c

    Dim lastRow As Long
    lastRow = src.Cells(src.Rows.Count, colMap("region")).End(xlUp).Row

    '---------------------------------------
    ' 2. Build dictionary of rows per Majlis
    '---------------------------------------
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim reg As String, maj As String, tz As String

    For r = 2 To lastRow

        tz = Trim(CStr(src.Cells(r, colMap("tanziem")).Value))
        If LCase(tz) = "nannamujahid" Then GoTo NextRow

        reg = Trim(CStr(src.Cells(r, colMap("region")).Value))
        maj = Trim(CStr(src.Cells(r, colMap("majlis")).Value))

        If reg = "" Or maj = "" Then GoTo NextRow

        Dim key As String
        key = CleanName(reg) & "||" & CleanName(maj)

        If Not dict.Exists(key) Then
            dict.Add key, New Collection
        End If

        dict(key).Add r

NextRow:
    Next r

    '---------------------------------------
    ' 3. Create folders + write each Majlis file
    '---------------------------------------

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim entry As Variant
    Dim regionFolder As String, majlisFile As String
    Dim wb As Workbook, ws As Worksheet, tbl As ListObject

    Dim regionsRoot As String
    regionsRoot = rootPath & "Regions" & Application.PathSeparator
    CreateFolderIfMissing regionsRoot

    For Each entry In dict.Keys

        Dim arr() As String
        arr = Split(entry, "||")

        reg = arr(0)
        maj = arr(1)

        regionFolder = regionsRoot & reg & Application.PathSeparator
        CreateFolderIfMissing regionFolder

        majlisFile = regionFolder & maj & ".xlsx"

        '-------------------------------
        ' Create new file or open existing
        '-------------------------------
        If Dir(majlisFile) = "" Then
            Set wb = Workbooks.Add(xlWBATWorksheet)
            Set ws = wb.Sheets(1)
            ws.Name = "Data"
            CreateMajlisTable ws
        Else
            Set wb = Workbooks.Open(majlisFile)
            Set ws = wb.Sheets("Data")
        End If

        Set tbl = ws.ListObjects("MajlisTable")

        '-------------------------------
        ' Write ALL rows for this Majlis
        '-------------------------------
        Dim rowIndex As Variant
        For Each rowIndex In dict(entry)

            Dim newRow As ListRow
            Set newRow = tbl.ListRows.Add

            newRow.Range(1, 1).Value = src.Cells(rowIndex, colMap("region")).Value
            newRow.Range(1, 2).Value = src.Cells(rowIndex, colMap("majlis")).Value
            newRow.Range(1, 3).Value = src.Cells(rowIndex, colMap("tanziem")).Value
            newRow.Range(1, 4).Value = src.Cells(rowIndex, colMap("jamaat id")).Value
            newRow.Range(1, 5).Value = src.Cells(rowIndex, colMap("name")).Value

            AddFormulaRow ws, newRow, tbl
        Next rowIndex

        ws.Columns(1).hidden = True
        ws.Columns(2).hidden = True

        StyleMajlisTable ws
        
        ' Apply data validation to columns 6-17 (Budget and Nov-Oct months)
        ApplyDataValidation ws
        
        ' Protect first 5 and last 3 columns with password
        ProtectSpecificColumns ws, password

        wb.SaveAs majlisFile, xlOpenXMLWorkbook
        wb.Close False

    Next entry

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "STEP 1 completed successfully!", vbInformation

End Sub



'=========================================================
' SUPPORTING FUNCTIONS
'=========================================================

Private Sub CreateMajlisTable(ws As Worksheet)

    Dim headers
    headers = Array("Region", "Majlis", "Tanziem", "Jamaat ID", "Name", _
                    "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", _
                    "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", _
                    "Paid", "Rest", "Percentage")
    

    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(1, UBound(headers) + 1), , xlYes)
    tbl.Name = "MajlisTable"

End Sub


Private Sub AddFormulaRow(ws As Worksheet, newRow As ListRow, tbl As ListObject)

    Dim rw As Long: rw = newRow.Range.Row

    Dim colBudget As Long: colBudget = tbl.ListColumns("Budget").Range.Column
    Dim colTotal As Long: colTotal = tbl.ListColumns("Paid").Range.Column
    Dim colRest As Long: colRest = tbl.ListColumns("Rest").Range.Column
    Dim colPct As Long: colPct = tbl.ListColumns("Percentage").Range.Column

    Dim colFirst As Long, colLast As Long

    ' Nov is column 7, Oct is column 18
    colFirst = tbl.ListColumns("Nov").Range.Column
    colLast = tbl.ListColumns("Oct").Range.Column

    ws.Cells(rw, colTotal).Formula = "=SUM(" & ws.Cells(rw, colFirst).Address(False, False) & ":" & ws.Cells(rw, colLast).Address(False, False) & ")"
    ws.Cells(rw, colRest).Formula = "=" & ws.Cells(rw, colBudget).Address(False, False) & "-" & ws.Cells(rw, colTotal).Address(False, False)
    ws.Cells(rw, colPct).Formula = "=IF(" & ws.Cells(rw, colBudget).Address(False, False) & "=0,""""," & ws.Cells(rw, colTotal).Address(False, False) & "/" & ws.Cells(rw, colBudget).Address(False, False) & ")"

    ws.Cells(rw, colPct).NumberFormat = "0.00%"

End Sub


Private Sub StyleMajlisTable(ws As Worksheet)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MajlisTable")

    With tbl.HeaderRowRange
        .Interior.Color = RGB(200, 220, 255)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    With tbl.Range.Borders
        .Color = RGB(180, 180, 180)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ws.Columns.AutoFit

End Sub


Private Sub ApplyDataValidation(ws As Worksheet)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MajlisTable")
    
    Dim totalCols As Long
    totalCols = tbl.Range.Columns.Count
    
    ' Apply data validation to columns 7-18 (Nov through Oct months)
    Dim startCol As Long, endCol As Long
    startCol = 7
    endCol = Application.Min(18, totalCols)
    
    Dim col As Long
    For col = startCol To endCol
        With tbl.ListColumns(col).DataBodyRange.Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="0", Formula2:="999999999"
        End With
    Next col
    
    ' Also apply to Budget column specifically (column 6) if it's within range
    If totalCols >= 6 Then
        With tbl.ListColumns(6).DataBodyRange.Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="0", Formula2:="999999999"
        End With
    End If
End Sub


Private Sub ProtectSpecificColumns(ws As Worksheet, pwd As String)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MajlisTable")
    
    Dim totalCols As Long
    totalCols = tbl.Range.Columns.Count
    
    ' Unprotect sheet first if it was protected
    ws.Unprotect pwd
    
    ' Unlock all cells first
    tbl.Range.Locked = False
    
    ' Protect first 5 columns (Region, Majlis, Tanziem, Jamaat ID, Name)
    If totalCols >= 5 Then
        Dim i As Long
        For i = 1 To 5
            tbl.ListColumns(i).DataBodyRange.Locked = True
        Next i
    End If
    
    ' Protect last 3 columns (Total, Rest, %)
'    If totalCols >= 18 Then 'for TJ
    If totalCols >= 14 Then 'for 100 Moschee
        tbl.ListColumns(totalCols - 2).DataBodyRange.Locked = True
        tbl.ListColumns(totalCols - 1).DataBodyRange.Locked = True
        tbl.ListColumns(totalCols).DataBodyRange.Locked = True
    End If
    
    ' Protect the worksheet with password
    ws.Protect pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub


Private Sub CreateFolderIfMissing(path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then MkDir path
End Sub

Private Function CleanName(s As String) As String
    Dim b
    s = Trim(CStr(s))

    For Each b In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, b, "_")
    Next b

    CleanName = s
End Function

