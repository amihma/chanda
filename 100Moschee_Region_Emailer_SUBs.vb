Option Explicit

'====================================================================
' MODULE: Region_Email_Module
'====================================================================


Public Const FROM_EMAIL As String = "your_email@domain.com"

'====================================================================
' FINANCIAL YEAR FORMATTER
'====================================================================
Function FinancialYear() As String
    Dim y As Long: y = Year(Date)
    If Month(Date) >= 7 Then
        FinancialYear = y & "/" & (y + 1)
    Else
        FinancialYear = (y - 1) & "/" & y
    End If
End Function

'====================================================================
' DOUBLE-CLICK HANDLER CALLED BY SHEET
'====================================================================
Sub Region_Email_DblClick(rowNum As Long)

    Dim ws As Worksheet: Set ws = Sheets("Region_Emails")

    Dim region As String, email As String
    region = ws.Cells(rowNum, 1).Value
    email = ws.Cells(rowNum, 2).Value

    If region = "" Or email = "" Then
        MsgBox "Region or email missing!", vbCritical
        Exit Sub
    End If

    Call Send_Region_Email_Single(region, email)

End Sub


'====================================================================
' PREPARE REGION EMAIL HTML
'====================================================================
Function Prepare_Region_HTML(regionName As String) As String

    Dim ws As Worksheet: Set ws = Sheets("Majalis")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim fy As String: fy = FinancialYear()
    
    Dim html As String: html = ""

    '================================================================
    ' HEADER TABLE (NO BORDER)
    '================================================================
    html = html & "<table style='width:100%; border-collapse:collapse;'>"
    html = html & "<tr>"
    html = html & "<td><img src='logo.png' style='height:80px;'></td>"
    html = html & "<td><strong>text<br>text<br>text</strong></td>"
    html = html & "</tr>"
    html = html & "</table>"

    html = html & "<h2>Asslam.o.Alaikum</h2>"

    '================================================================
    ' START DATA TABLE (WITH SOLID BLACK BORDER)
    '================================================================
    html = html & "<table style='border-collapse:collapse;' border='1' cellpadding='5'>"

    ' HEADING ROW
    html = html & "<tr><td colspan='13' style='text-align:center;'><h2>Chanda 100 Moschee Übersicht Jahr " & fy & "</h2></td></tr>"
    html = html & "<tr><td colspan='13' style='text-align:center;'>Region: </strong>" & regionName & "</strong></td></tr>"

    ' TOP HEADERS
    html = html & "<tr>"
    html = html & "<td rowspan='2'><strong>Majlis</strong></td>"
    html = html & "<td colspan='6'><strong>Khuddam</strong></td>"
    html = html & "<td colspan='6'><strong>Atfal</strong></td>"
    html = html & "</tr>"

    ' SUB-HEADERS
    html = html & "<tr>"
    html = html & "<td><strong>Tajneed</strong></td><td><strong>Nicht-Zahler</strong></td><td><strong>Budget</strong></td><td><strong>Bezahlt</strong></td><td><strong>Rest</strong></td><td><strong>%</strong></td>"
    html = html & "<td><strong>Tajneed</strong></td><td><strong>Nicht-Zahler</strong></td><td><strong>Budget</strong></td><td><strong>Bezahlt</strong></td><td><strong>Rest</strong></td><td><strong>%</strong></td>"
    html = html & "</tr>"

    '================================================================
    ' BUILD ONE ROW PER MAJLIS (MERGING KHADIM + TIFL)
    '================================================================
    Dim r As Long
    Dim maj As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    ' STEP 1 — collect Majlis names for this region
    For r = 2 To lastRow
        If ws.Cells(r, "B").Value = regionName Then
            maj = ws.Cells(r, "A").Value
            If Not dict.Exists(maj) Then dict.Add maj, 1
        End If
    Next r

    ' STEP 2 — for each Majlis build a combined row
    Dim khRow As Long, tfRow As Long
    Dim i As Long

    For i = 0 To dict.Count - 1
        
        maj = dict.Keys()(i)

        khRow = 0
        tfRow = 0

        ' Find rows for this majlis
        For r = 2 To lastRow
            If ws.Cells(r, "B").Value = regionName And ws.Cells(r, "A").Value = maj Then
                If ws.Cells(r, "C").Value = "Khadim" Then
                    khRow = r
                ElseIf ws.Cells(r, "C").Value = "Tifl" Then
                    tfRow = r
                End If
            End If
        Next r

        ' Extract Khadim values
        Dim k_taj As Variant, k_nz As Variant, k_bd As Variant, k_bz As Variant, k_rs As Variant, k_pc As Variant

        If khRow > 0 Then
            k_taj = ws.Cells(khRow, "D").Value
            k_nz = ws.Cells(khRow, "E").Value
            k_bd = ws.Cells(khRow, "F").Value
            k_bz = ws.Cells(khRow, "O").Value
            k_rs = ws.Cells(khRow, "P").Value
            k_pc = Round(ws.Cells(khRow, "Q").Value * 100, 2)

        Else
            k_taj = "": k_nz = "": k_bd = "": k_bz = "": k_rs = "": k_pc = ""
        End If

        ' Extract Tifl values
        Dim t_taj As Variant, t_nz As Variant, t_bd As Variant, t_bz As Variant, t_rs As Variant, t_pc As Variant

        If tfRow > 0 Then
            t_taj = ws.Cells(tfRow, "D").Value
            t_nz = ws.Cells(tfRow, "E").Value
            t_bd = ws.Cells(tfRow, "F").Value
            t_bz = ws.Cells(tfRow, "O").Value
            t_rs = ws.Cells(tfRow, "P").Value
            t_pc = Round(ws.Cells(tfRow, "Q").Value * 100, 2)

        Else
            ' No Tifl row ? leave blank
            t_taj = "": t_nz = "": t_bd = "": t_bz = "": t_rs = "": t_pc = ""
        End If

        '================================================================
        ' ADD ROW TO HTML
        '================================================================
        html = html & "<tr>"

        html = html & "<td><strong>" & maj & "</strong></td>"

        'Khuddam
        html = html & "<td>" & k_taj & "</td><td>" & k_nz & "</td><td>" & k_bd _
                     & "</td><td>" & k_bz & "</td><td>" & k_rs & "</td><td>" & k_pc & "%</td>"

        'Atfal
        html = html & "<td>" & t_taj & "</td><td>" & t_nz & "</td><td>" & t_bd _
                     & "</td><td>" & t_bz & "</td><td>" & t_rs & "</td><td>" & t_pc & "%</td>"

        html = html & "</tr>"

    Next i

    html = html & "</table>"

    '================================================================
    ' SIGNATURE
    '================================================================
    html = html & "<p>Wassalam<br>Name<br>Kontakt</p>"

    Prepare_Region_HTML = html

End Function


'====================================================================
' SEND REGION EMAIL (SINGLE)
'====================================================================
Sub Send_Region_Email_Single(regionName As String, sendTo As String)

    Dim html As String
    html = Prepare_Region_HTML(regionName)

    If html = "" Then
        MsgBox "Region not found!", vbCritical
        Exit Sub
    End If

    Dim o As Object, m As Object
    Set o = CreateObject("Outlook.Application")
    Set m = o.CreateItem(0)

    With m
        .To = sendTo
        .Sender = FROM_EMAIL
        .Subject = "Chanda Summary – Region " & regionName
        .htmlBody = html
        .Display
    End With

End Sub


'====================================================================
' BULK REGION EMAIL
'====================================================================
Sub Send_Region_Email_Bulk()

    If MsgBox("Send bulk region emails?", vbYesNo) = vbNo Then Exit Sub

    Dim ws As Worksheet: Set ws = Sheets("Region_Emails")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, "C").Value <> "" Then
            Call Send_Region_Email_Single(ws.Cells(r, "A").Value, ws.Cells(r, "B").Value)
        End If
    Next r

End Sub



