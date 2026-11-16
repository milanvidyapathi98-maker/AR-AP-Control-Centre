VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Customer Search"
   ClientHeight    =   5784
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10656
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
    Dim lo As ListObject, v As Variant, i As Long, r As Long
    Dim q As String, n As Long, tot As Double
    Dim cID As Long, cName As Long, cInv As Long, cDue As Long
    Dim cOpen As Long, cEUR As Long, cStat As Long

    Set lo = ThisWorkbook.Worksheets("AR_OpenItems").ListObjects("tblAROpenItems")
    lstItems.Clear
    lblSummary.Caption = ""

    q = UCase$(Trim$(txtSearch.Value))
    If Len(q) = 0 Then
        lblSummary.Caption = "Enter a customer ID or name."
        txtSearch.SetFocus
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then
        lblSummary.Caption = "No open items in the table."
        Exit Sub
    End If

    ' Look up column positions once
    cID = lo.ListColumns("CustomerID").Index
    cName = lo.ListColumns("CustomerName").Index
    cInv = lo.ListColumns("InvoiceNo").Index
    cDue = lo.ListColumns("DueDate").Index
    cOpen = lo.ListColumns("OpenAmount").Index
    cEUR = lo.ListColumns("OpenEUR").Index
    cStat = lo.ListColumns("Status").Index

    v = lo.DataBodyRange.Value

    For i = 1 To UBound(v, 1)
        ' Skip rows with error values (#N/A etc.)
        If Not IsError(v(i, cID)) And Not IsError(v(i, cName)) _
           And Not IsError(v(i, cOpen)) And Not IsError(v(i, cEUR)) Then

            If Val(v(i, cOpen)) <> 0 Then
                If UCase$(CStr(v(i, cID))) = q _
                   Or InStr(1, UCase$(CStr(v(i, cName))), q) > 0 Then

                    lstItems.AddItem CStr(v(i, cInv))
                    r = lstItems.ListCount - 1
                    lstItems.List(r, 1) = CStr(v(i, cName))
                    If IsDate(v(i, cDue)) Then lstItems.List(r, 2) = Format(v(i, cDue), "dd.mm.yyyy")
                    lstItems.List(r, 3) = Format(v(i, cOpen), "#,##0.00")
                    If Not IsError(v(i, cStat)) Then lstItems.List(r, 4) = CStr(v(i, cStat))

                    n = n + 1
                    tot = tot + v(i, cEUR)
                End If
            End If
        End If
    Next i

    If n = 0 Then
        lblSummary.Caption = "No open items found for """ & txtSearch.Value & """."
    Else
        lblSummary.Caption = n & " open items  |  EUR " & Format(tot, "#,##0.00")
    End If
End Sub


