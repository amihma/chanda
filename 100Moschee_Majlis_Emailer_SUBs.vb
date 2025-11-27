'====================================================================
' MODULE: Majlis_Email_Module
'====================================================================

Option Explicit

Public Const FROM_EMAIL As String = "tj@khuddam.de"

'====================================================================
' FINANCIAL YEAR STRING MAKER
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
' DOUBLE-CLICK HANDLER (called by sheet code)
'====================================================================
Sub Majlis_Email_DblClick(rowNum As Long)

    Dim ws As Worksheet: Set ws = Sheets("Majlis_Emails")

    Dim majlis As String, email As String
    majlis = ws.Cells(rowNum, 1).Value
    email = ws.Cells(rowNum, 2).Value

    If majlis = "" Or email = "" Then
        MsgBox "Majlis name or email missing!", vbCritical
        Exit Sub
    End If

    Call Send_Majlis_Email_Single(majlis, email)

End Sub


'====================================================================
' PREPARE HTML FOR MAJLIS EMAIL
'====================================================================
Function Prepare_Majlis_HTML(majlisName As String) As String

    Dim ws As Worksheet: Set ws = Sheets("Majalis")
    Dim r As Variant
    r = Application.Match(majlisName, ws.Columns("A"), 0)
    If IsError(r) Then Exit Function

    Dim tanziem As String: tanziem = ws.Cells(r, "C").Value

    Dim isK As Boolean: isK = (tanziem = "Khadim")
    Dim isT As Boolean: isT = (tanziem = "Tifl")

    'Extract values
    Dim taj As Long: taj = ws.Cells(r, "D").Value
    Dim nz As Long: nz = ws.Cells(r, "E").Value
    Dim bdg As Double: bdg = ws.Cells(r, "F").Value
    Dim bez As Double: bez = ws.Cells(r, "O").Value
    Dim rst As Double: rst = ws.Cells(r, "P").Value
    Dim pct As Double: pct = Round(ws.Cells(r, "Q").Value * 100, 2)

    Dim fy As String: fy = FinancialYear()

    '====================================================================
    ' BUILD HTML
    '====================================================================
    Dim html As String: html = ""

    'HEADER TABLE (no border)
    html = html & "<table style='width:100%; background:#ED8F0C; border-collapse:collapse;'>"
    html = html & "<tr>"
    html = html & "<td style='width:50px; vertical-align:top;'><img src='https://khuddam.de/wp-content/uploads/2024/01/cropped-MKADLogo-150x150.png' width='100px' height='100px'></td>"
    html = html & "<td style='vertical-align:middle; font-family:Arial; font-size:22px; align:right;'>Abteilung Tehrik-e-Jadid <br />Majlis Khuddam-ul-Ahmadiyya <br />Deutschland</td>"
    html = html & "</tr></table>"

    html = html & "<h2>Asslam.o.Alaikum</h2>"

    'DATA TABLE (black border)
    html = html & "<table style='border-collapse:collapse;' border='1' cellpadding='5'>"

    html = html & "<tr><td colspan='4' style='text-align:center;'><h2>Chanda 100 Moschee Übersicht Jahr " & fy & "</h2></td></tr>"

    html = html & "<tr><td><strong>Majlis</strong></td><td colspan='3'><strong>" & majlisName & "</strong></td></tr>"

    html = html & "<tr><td colspan='4'></td></tr>"

    html = html & "<tr><td><strong>Feld</strong></td><td><strong>Khuddam</strong></td><td><strong>Atfal</strong></td><td><strong>Total</strong></td></tr>"

    html = html & "<tr><td><strong>Tajneed</strong></td><td>" & IIf(isK, taj, "") & "</td><td>" & IIf(isT, taj, "") _
            & "</td><td>" & taj & "</td></tr>"

    html = html & "<tr><td><strong>Nicht-Zahler</strong></td><td>" & IIf(isK, nz, "") & "</td><td>" & IIf(isT, nz, "") _
            & "</td><td>" & nz & "</td></tr>"

    html = html & "<tr><td><strong>Budget</strong></td><td>" & IIf(isK, bdg, "") & "</td><td>" & IIf(isT, bdg, "") _
            & "</td><td>" & bdg & "</td></tr>"

    html = html & "<tr><td><strong>Bezahlt</strong></td><td>" & IIf(isK, bez, "") & "</td><td>" & IIf(isT, bez, "") _
            & "</td><td>" & bez & "</td></tr>"

    html = html & "<tr><td><strong>Rest</strong></td><td>" & IIf(isK, rst, "") & "</td><td>" & IIf(isT, rst, "") _
            & "</td><td>" & rst & "</td></tr>"

    html = html & "<tr><td><strong>%</strong></td><td>" & IIf(isK, pct & "%", "") & "</td><td>" & IIf(isT, pct & "%", "") _
            & "</td><td>" & pct & "%</td></tr>"

    html = html & "</table>"

    'SIGNATURE
    html = html & "<p>Wassalam<br>Name<br>Kontakt</p>"

    Prepare_Majlis_HTML = html

End Function

Sub Send_Majlis_Email_Single(majlis As String, sendTo As String)

    Dim html As String
    html = Prepare_Majlis_HTML(majlis)

    If html = "" Then
        MsgBox "Majlis data not found!", vbCritical
        Exit Sub
    End If

    Dim o As Object, m As Object, acc As Object
    Set o = CreateObject("Outlook.Application")
    Set m = o.CreateItem(0)

    With m
        .To = sendTo
        .Subject = "Chanda Summary – " & majlis
        .htmlBody = html

        '==========================================
        ' FORCE SPECIFIC FROM EMAIL (works 100%)
        '==========================================
        For Each acc In o.Session.Accounts
            If LCase(acc.SmtpAddress) = LCase(FROM_EMAIL) Then
                Set .SendUsingAccount = acc
                Exit For
            End If
        Next acc

        .Display
    End With

End Sub



'====================================================================
' BULK SEND
'====================================================================
Sub Send_Majlis_Email_Bulk()

    If MsgBox("Send bulk Majlis emails?", vbYesNo) = vbNo Then Exit Sub

    Dim ws As Worksheet: Set ws = Sheets("Majlis_Emails")
    Dim r As Long, lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, "C").Value <> "" Then
            Call Send_Majlis_Email_Single(ws.Cells(r, "A").Value, ws.Cells(r, "B").Value)
        End If
    Next r

End Sub



