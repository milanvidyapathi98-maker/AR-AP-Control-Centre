Attribute VB_Name = "modARAP"
Option Explicit

' Returns the column number of a table column, or a clear error if it's missing
Private Function ColIx(lo As ListObject, colName As String) As Long
    On Error Resume Next
    ColIx = lo.ListColumns(colName).Index
    On Error GoTo 0
    If ColIx = 0 Then Err.Raise vbObjectError + 1, , _
        "Column '" & colName & "' missing in " & lo.Name
End Function

' Finds a table by name on any sheet
Private Function GetTable(tblName As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetTable = ws.ListObjects(tblName)
        On Error GoTo 0
        If Not GetTable Is Nothing Then Exit Function
    Next ws
    Err.Raise vbObjectError + 2, , "Table '" & tblName & "' not found"
End Function

' Returns a sheet (creates it if missing) and clears it
Private Function GetSheet(nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheet.Name = nm
    End If
    GetSheet.Cells.Clear
End Function

Public Sub RefreshEverything()
    Dim t As Double, cn As WorkbookConnection
    On Error GoTo Fail
    t = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing queries..."

    ' 1) Turn off background refresh so VBA waits for each query
    For Each cn In ThisWorkbook.Connections
        If cn.Type = xlConnectionTypeOLEDB Then cn.OLEDBConnection.BackgroundQuery = False
    Next cn

    ' 2) Refresh all queries and recalculate everything
    ThisWorkbook.RefreshAll
    Application.CalculateFull

    ' 3) Write the timestamp
    With ThisWorkbook.Worksheets("Parameters")
        .Range("D2").Value = "Last refreshed"
        .Range("E2").Value = Now
        .Range("E2").NumberFormat = "dd.mm.yyyy hh:mm"
    End With

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Refresh complete in " & Format(Timer - t, "0.0") & " seconds.", vbInformation
    Exit Sub

Fail:
    ' Always restore the screen and status bar, even if something breaks
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub

Public Sub ReconcileBank()
    Dim loB As ListObject, loR As ListObject, loP As ListObject
    Dim vB As Variant, vR As Variant, vP As Variant
    Dim dR As Object, dRID As Object, dP As Object, dPID As Object, seen As Object, cnt As Object
    Dim i As Long, j As Long, n As Long, k As String, tol As Double
    Dim amt As Double, d As Long, ref As String, refs As String, parts() As String
    Dim total As Double, ok As Boolean, st As String, matchID As String
    Dim outB() As Variant, ws As Worksheet, wsX As Worksheet, msg As String, key As Variant

    On Error GoTo Fail
    Application.ScreenUpdating = False
    tol = ThisWorkbook.names("Tolerance").RefersToRange.Value

    Set loB = GetTable("tblBank")
    Set loR = GetTable("tblARReceipts")
    Set loP = GetTable("tblAPPayments")
    vB = loB.DataBodyRange.Value          ' read whole tables into memory (fast)
    vR = loR.DataBodyRange.Value
    vP = loP.DataBodyRange.Value

    Set dR = CreateObject("Scripting.Dictionary")   ' InvoiceNo|Date -> EUR amount
    Set dRID = CreateObject("Scripting.Dictionary") ' InvoiceNo|Date -> ReceiptID(s)
    Set dP = CreateObject("Scripting.Dictionary")   ' BillID|Date -> EUR amount
    Set dPID = CreateObject("Scripting.Dictionary") ' BillID|Date -> PaymentID
    Set seen = CreateObject("Scripting.Dictionary") ' receipts found on the bank statement
    Set cnt = CreateObject("Scripting.Dictionary")  ' count per status

    ' ---- 1) Load ERP receipts and vendor payments into dictionaries ----
    Dim cInv As Long, cRD As Long, cRA As Long, cRID As Long
    cInv = ColIx(loR, "AppliedInvoiceNo"): cRD = ColIx(loR, "ReceiptDate")
    cRA = ColIx(loR, "AmountEUR"):        cRID = ColIx(loR, "ReceiptID")
    For i = 1 To UBound(vR, 1)
        k = vR(i, cInv) & "|" & CLng(vR(i, cRD))
        dR(k) = dR(k) + vR(i, cRA)
        If Len(dRID(k)) > 0 Then dRID(k) = dRID(k) & ", "
        dRID(k) = dRID(k) & vR(i, cRID)
    Next i

    Dim cPB As Long, cPD As Long, cPA As Long, cPID As Long
    cPB = ColIx(loP, "BillID"):        cPD = ColIx(loP, "PaymentDate")
    cPA = ColIx(loP, "AmountPaidEUR"): cPID = ColIx(loP, "PaymentID")
    For i = 1 To UBound(vP, 1)
        k = vP(i, cPB) & "|" & CLng(vP(i, cPD))
        dP(k) = dP(k) + vP(i, cPA)
        dPID(k) = vP(i, cPID)
    Next i

    ' ---- 2) Loop through every bank line and decide its status ----
    Dim cBD As Long, cCP As Long, cRef As Long, cAmt As Long, cRefs As Long, cLID As Long
    cBD = ColIx(loB, "BookingDate"): cCP = ColIx(loB, "Counterparty")
    cRef = ColIx(loB, "Reference"):  cAmt = ColIx(loB, "AmountEUR")
    cRefs = ColIx(loB, "InvoiceRefs"): cLID = ColIx(loB, "LineID")

    ReDim outB(1 To UBound(vB, 1), 1 To 7)
    For i = 1 To UBound(vB, 1)
        amt = vB(i, cAmt)
        d = CLng(vB(i, cBD))
        ref = Trim(CStr(vB(i, cRef)))
        refs = Trim(CStr(vB(i, cRefs)))
        matchID = ""

        If amt > 0 Then                                   ' money in
            If Len(refs) = 0 Then
                st = "Unapplied cash"
            Else
                parts = Split(refs, ",")
                total = 0: ok = True
                For j = LBound(parts) To UBound(parts)
                    k = Trim(parts(j)) & "|" & d
                    If dR.Exists(k) Then
                        total = total + dR(k)
                        If Len(matchID) > 0 Then matchID = matchID & ", "
                        matchID = matchID & dRID(k)
                        seen(k) = True
                    Else
                        ok = False
                    End If
                Next j
                If ok And Abs(total - amt) <= tol Then st = "Matched AR" Else st = "Unmatched AR"
            End If
        ElseIf UCase(Left(ref, 5)) = "AP26-" Then          ' money out to a vendor
            k = Left(ref, 10) & "|" & d
            If dP.Exists(k) Then
                st = "Matched AP": matchID = dPID(k)
            Else
                st = "Unmatched AP"
            End If
        Else                                              ' payroll, tax, bank fees
            st = "Other (no sub-ledger)"
        End If

        cnt(st) = cnt(st) + 1
        outB(i, 1) = vB(i, cLID): outB(i, 2) = vB(i, cBD): outB(i, 3) = vB(i, cCP)
        outB(i, 4) = vB(i, cRef): outB(i, 5) = amt: outB(i, 6) = st: outB(i, 7) = matchID
    Next i

    ' ---- 3) Write results in one go and colour the status ----
    Set ws = GetSheet("Bank_Recon")
    ws.Range("A1:G1").Value = Array("LineID", "BookingDate", "Counterparty", "Reference", _
                                    "AmountEUR", "Status", "MatchID")
    ws.Range("A2").Resize(UBound(outB, 1), 7).Value = outB
    ws.Columns("B").NumberFormat = "dd.mm.yyyy"
    ws.Columns("E").NumberFormat = "#,##0.00"
    For i = 1 To UBound(outB, 1)
        With ws.Cells(i + 1, 6).Interior
            Select Case Left(outB(i, 6), 7)
                Case "Matched": .Color = RGB(198, 239, 206)   ' green
                Case "Other (": .Color = RGB(217, 217, 217)   ' grey
                Case Else:      .Color = RGB(255, 199, 206)   ' red
            End Select
        End With
    Next i
    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit

    ' ---- 4) Exceptions: unmatched bank lines + receipts not on the bank statement ----
    Set wsX = GetSheet("Recon_Exceptions")
    wsX.Range("A1:E1").Value = Array("Type", "ID", "Date", "Reference / Invoice", "AmountEUR")
    n = 1
    For i = 1 To UBound(outB, 1)
        If Left(outB(i, 6), 7) <> "Matched" And outB(i, 6) <> "Other (no sub-ledger)" Then
            n = n + 1
            wsX.Cells(n, 1).Resize(1, 5).Value = Array("Bank: " & outB(i, 6), outB(i, 1), _
                                                       outB(i, 2), outB(i, 4), outB(i, 5))
        End If
    Next i
    Dim nNotOnBank As Long
    For i = 1 To UBound(vR, 1)
        k = vR(i, cInv) & "|" & CLng(vR(i, cRD))
        If Not seen.Exists(k) Then
            n = n + 1: nNotOnBank = nNotOnBank + 1
            wsX.Cells(n, 1).Resize(1, 5).Value = Array("ERP receipt not on bank statement", _
                vR(i, cRID), vR(i, cRD), vR(i, cInv), vR(i, cRA))
        End If
    Next i
    wsX.Columns("C").NumberFormat = "dd.mm.yyyy"
    wsX.Columns("E").NumberFormat = "#,##0.00"
    wsX.Range("A1:E1").Font.Bold = True
    wsX.Columns("A:E").AutoFit

    ' ---- Summary ----
    For Each key In cnt.Keys
        msg = msg & key & ": " & cnt(key) & vbCrLf
    Next key
    msg = msg & "ERP receipts not on bank statement: " & nNotOnBank
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, "Bank reconciliation"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Reconciliation failed: " & Err.Description, vbCritical
End Sub

Public Sub ExportCustomerStatements()
    Dim lo As ListObject, v As Variant, i As Long, r As Long, n As Long
    Dim worst As Object, names As Object, key As Variant
    Dim cCust As Long, cName As Long, cInv As Long, cType As Long, cID As Long, cDD As Long
    Dim colCur As Long, cOpen As Long, cOpenE As Long, cDPD As Long, cDisp As Long
    Dim dpd As Variant, lvl As Long, lvlText As String, folder As String
    Dim ws As Worksheet, tot As Double

    On Error GoTo Fail
    If ThisWorkbook.Path = "" Then MsgBox "Please save the workbook first.", vbExclamation: Exit Sub
    Application.ScreenUpdating = False

    Set lo = GetTable("tblAROpenItems")
    v = lo.DataBodyRange.Value
    cCust = ColIx(lo, "CustomerID"):   cName = ColIx(lo, "CustomerName")
    cInv = ColIx(lo, "InvoiceNo"):     cType = ColIx(lo, "DocType")
    cID = ColIx(lo, "InvoiceDate"):    cDD = ColIx(lo, "DueDate")
    colCur = ColIx(lo, "Currency"):      cOpen = ColIx(lo, "OpenAmount")
    cOpenE = ColIx(lo, "OpenEUR"):     cDPD = ColIx(lo, "DaysPastDue")
    cDisp = ColIx(lo, "Disputed")

    ' ---- 1) Customers with overdue, non-disputed items + their worst days overdue ----
    Set worst = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        dpd = v(i, cDPD)
        If v(i, cOpen) <> 0 And UCase(v(i, cDisp)) <> "Y" _
           And IsNumeric(dpd) And Len(CStr(dpd)) > 0 Then
            If dpd > 0 Then
                If dpd > worst(v(i, cCust)) Then worst(v(i, cCust)) = dpd
                names(v(i, cCust)) = v(i, cName)
            End If
        End If
    Next i
    If worst.Count = 0 Then MsgBox "No overdue customers.", vbInformation: GoTo Done

    ' ---- 2) Output folder Statements_<today> next to the workbook ----
    folder = ThisWorkbook.Path & "\Statements_" & Format(Date, "yyyy-mm-dd")
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Set ws = GetSheet("Statement")
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ' ---- 3) One statement per customer, exported as PDF ----
    For Each key In worst.Keys
        ' Dunning level from the worst days overdue
        Select Case worst(key)
            Case Is <= 30: lvl = 1: lvlText = "Payment reminder"
            Case Is <= 60: lvl = 2: lvlText = "Second reminder"
            Case Else:     lvl = 3: lvlText = "Final notice"
        End Select

        ws.Cells.Clear
        ws.Range("A1").Value = ThisWorkbook.Worksheets("Parameters").Range("B2").Value
        ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
        ws.Range("A3").Value = "Customer statement – " & lvlText
        ws.Range("A3").Font.Bold = True
        ws.Range("A4").Value = "Customer:":  ws.Range("B4").Value = key & "  " & names(key)
        ws.Range("A5").Value = "Date:":      ws.Range("B5").Value = Format(ThisWorkbook.Names("AsOfDate").RefersToRange.Value, "dd.mm.yyyy")

        ws.Range("A7:H7").Value = Array("Document", "Type", "Invoice date", "Due date", _
                                        "Currency", "Open amount", "Days overdue", "Open EUR")
        ws.Range("A7:H7").Font.Bold = True
        ws.Range("A7:H7").Interior.Color = RGB(221, 235, 247)

        r = 7: tot = 0
        For i = 1 To UBound(v, 1)
            If v(i, cCust) = key And v(i, cOpen) <> 0 Then     ' all open items incl. credit notes
                r = r + 1
                ws.Cells(r, 1).Resize(1, 8).Value = Array(v(i, cInv), v(i, cType), v(i, cID), _
                    v(i, cDD), v(i, colCur), v(i, cOpen), v(i, cDPD), v(i, cOpenE))
                tot = tot + v(i, cOpenE)
            End If
        Next i
        ws.Cells(r + 2, 7).Value = "Total EUR"
        ws.Cells(r + 2, 8).Value = tot
        ws.Cells(r + 2, 7).Resize(1, 2).Font.Bold = True
        ws.Cells(r + 4, 1).Value = "Please transfer the overdue amount within 7 days, " & _
            "quoting the document numbers. If you have already paid, please ignore this letter."

        ws.Range("C:D").NumberFormat = "dd.mm.yyyy"
        ws.Range("F:F,H:H").NumberFormat = "#,##0.00"
        ws.Columns("A:H").AutoFit

        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folder & "\" & key & "_Level" & lvl & ".pdf"
        n = n + 1
    Next key

    MsgBox n & " statements exported to:" & vbCrLf & folder, vbInformation
Done:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Statement export failed: " & Err.Description, vbCritical
End Sub

Public Sub CreatePaymentRun()
    Dim lo As ListObject, v As Variant, i As Long, n As Long, f As Integer
    Dim horizon As Variant, runDate As Date, endDate As Date
    Dim cBill As Long, cVen As Long, cVIN As Long, colCur As Long, cAppr As Long, cDup As Long
    Dim cOpen As Long, cOpenE As Long, cDisc As Long, cDiscD As Long, cDue As Long
    Dim action As String, payDate As Date, pay As Double, payE As Double, totE As Double
    Dim outP() As Variant, ws As Worksheet, csvPath As String

    On Error GoTo Fail
    If ThisWorkbook.Path = "" Then MsgBox "Please save the workbook first.", vbExclamation: Exit Sub

    ' ---- 1) Ask for the horizon ----
    horizon = Application.InputBox("Pay bills due within how many days?", "Payment run", 7, Type:=1)
    If horizon = False Then Exit Sub                 ' user clicked Cancel
    runDate = ThisWorkbook.names("ForecastStart").RefersToRange.Value
    endDate = runDate + horizon
    Application.ScreenUpdating = False

    Set lo = GetTable("tblAPBills")
    v = lo.DataBodyRange.Value
    cBill = ColIx(lo, "BillID"):         cVen = ColIx(lo, "VendorID")
    cVIN = ColIx(lo, "VendorInvoiceNo"): colCur = ColIx(lo, "Currency")
    cAppr = ColIx(lo, "ApprovalStatus"): cDup = ColIx(lo, "DupFlag")
    cOpen = ColIx(lo, "OpenAmount"):     cOpenE = ColIx(lo, "OpenEUR")
    cDisc = ColIx(lo, "DiscPct"):        cDiscD = ColIx(lo, "DiscountDate")
    cDue = ColIx(lo, "DueDate")

    ' ---- 2) Select bills: Approved, open, not a duplicate, due inside the horizon ----
    ReDim outP(1 To UBound(v, 1), 1 To 9)
    For i = 1 To UBound(v, 1)
        If v(i, cAppr) = "Approved" And v(i, cOpen) > 0 And v(i, cDup) <> "Duplicate" Then
            action = ""
            If v(i, cDisc) > 0 And IsDate(v(i, cDiscD)) Then
                If v(i, cDiscD) >= runDate And v(i, cDiscD) <= endDate Then
                    action = "Pay with discount"
                    payDate = v(i, cDiscD)
                    pay = Round(v(i, cOpen) * (1 - v(i, cDisc)), 2)
                    payE = Round(v(i, cOpenE) * (1 - v(i, cDisc)), 2)
                End If
            End If
            If action = "" And v(i, cDue) <= endDate Then
                If v(i, cDue) < runDate Then
                    action = "Overdue - pay now": payDate = runDate
                Else
                    action = "Pay in full": payDate = v(i, cDue)
                End If
                pay = v(i, cOpen): payE = v(i, cOpenE)
            End If

            If action <> "" Then
                n = n + 1
                outP(n, 1) = v(i, cVen):  outP(n, 2) = v(i, cBill): outP(n, 3) = v(i, cVIN)
                outP(n, 4) = v(i, colCur):  outP(n, 5) = v(i, cOpen): outP(n, 6) = action
                outP(n, 7) = payDate:     outP(n, 8) = pay:         outP(n, 9) = payE
                totE = totE + payE
            End If
        End If
    Next i

    ' ---- 3) Write the proposal to Payment_Run ----
    Set ws = GetSheet("Payment_Run")
    ws.Range("A1").Value = "Payment run " & Format(runDate, "dd.mm.yyyy") & _
                           " – horizon " & horizon & " days (until " & Format(endDate, "dd.mm.yyyy") & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:I3").Value = Array("VendorID", "BillID", "VendorInvoiceNo", "Currency", _
                                    "OpenAmount", "Action", "PayDate", "AmountToPay", "AmountToPayEUR")
    ws.Range("A3:I3").Font.Bold = True
    ws.Range("A3:I3").Interior.Color = RGB(221, 235, 247)
    If n > 0 Then ws.Range("A4").Resize(n, 9).Value = outP
    ws.Cells(n + 5, 8).Value = "Total EUR"
    ws.Cells(n + 5, 9).Value = totE
    ws.Cells(n + 5, 8).Resize(1, 2).Font.Bold = True
    ws.Columns("G").NumberFormat = "dd.mm.yyyy"
    ws.Range("E:E,H:I").NumberFormat = "#,##0.00"
    ws.Columns("A:I").AutoFit

    ' ---- 4) Bank upload file PaymentRun_<date>.csv ----
    csvPath = ThisWorkbook.Path & "\PaymentRun_" & Format(runDate, "yyyy-mm-dd") & ".csv"
    f = FreeFile
    Open csvPath For Output As #f
    Print #f, "VendorID;BillID;VendorInvoiceNo;Currency;AmountToPay"
    For i = 1 To n
        Print #f, outP(i, 1) & ";" & outP(i, 2) & ";" & outP(i, 3) & ";" & _
                  outP(i, 4) & ";" & Format(outP(i, 8), "0.00")
    Next i
    Close #f

    Application.ScreenUpdating = True
    MsgBox n & " bills selected, total EUR " & Format(totE, "#,##0.00") & vbCrLf & _
           "CSV saved: " & csvPath, vbInformation
    Exit Sub
Fail:
    If f > 0 Then Close #f
    Application.ScreenUpdating = True
    MsgBox "Payment run failed: " & Err.Description, vbCritical
End Sub

Public Sub RunMonthEnd()
    RefreshEverything
    ReconcileBank
    ExportCustomerStatements
    CreatePaymentRun
End Sub

Public Sub ShowSearch(): frmSearch.Show: End Sub
