Attribute VB_Name = "Module1"
Option Explicit

'======== CONFIGURATION CONSTANTS ==========
Const SH_ORDERS As String = "Main Order Sheet (Downloaded)"
Const SH_PAYMENTS As String = "Payments Received"
Const SH_PACKAGES As String = "Pre-Paid Packages"

Const SH_PKG_ORDERS As String = "Package Orders"
Const SH_PAID_ORDERS As String = "Paid Orders"
Const SH_BAL_PENDING As String = "Balance Pending Orders"
Const SH_PEND_DEL As String = "Pending Delivery Orders"
Const SH_DEL_NO_PAY As String = "Delivered Without Payment"
Const SH_DEL_WITH_BAL As String = "Delivered With Balance"
Const SH_PAID_PKGS As String = "Paid Packages"
Const SH_PENDING_PKGS As String = "Pending Packages"
Const SH_ADV_PAY As String = "Advance Payments Details"
Const SH_ADJUST As String = "Total Adjustments"
Const SH_OUTSTANDING As String = "Customer Outstanding Summary"

' Column numbers in Main Order Sheet
Const COL_ORDERDATE As Long = 2           ' "Order Date / Time"
Const COL_ORDERNO As Long = 3             ' "Order No."
Const COL_NAME As Long = 4                ' "Name"
Const COL_DELIVEREDON As Long = 9         ' "Delivered On"
Const COL_GROSS As Long = 12              ' "Gross Amount"
Const COL_DISCOUNT As Long = 13           ' "Discount"
Const COL_TAX As Long = 14                ' "Tax"
Const COL_NET As Long = 15                ' "Net Amount"
Const COL_ADVANCE As Long = 16            ' "Advance"
Const COL_PAID As Long = 17               ' "Paid"
Const COL_ADJUSTMENT As Long = 18         ' "Adjustment"
Const COL_BALANCE As Long = 19            ' "Balance"
Const COL_HOMEDEL As Long = 25            ' "Home Delivery"
Const COL_PACKAGE As Long = 31            ' "Package"
Const COL_ORDERSTATUS As Long = 39        ' "Order Status"

' Columns in Payments Received
Const PAY_COL_ORDERNO As Long = 3         ' "Order No."
Const PAY_COL_ISADV As Long = 7           ' "Is Advance?"
Const PAY_COL_ISPKG As Long = 8           ' "Is Package Payment?"
Const PAY_COL_AMOUNT As Long = 9          ' "Payment Amount"

' Columns in Pre-Paid Packages
Const PKG_COL_STATUS As Long = 9          ' "Payment Status"

'==========================================
' MASTER MACRO – call this from a button
'==========================================
Public Sub Refresh_All_Reports()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanExit
    
    ' 1. Ensure all report sheets exist
    EnsureReportSheet SH_PKG_ORDERS
    EnsureReportSheet SH_PAID_ORDERS
    EnsureReportSheet SH_BAL_PENDING
    EnsureReportSheet SH_PEND_DEL
    EnsureReportSheet SH_DEL_NO_PAY
    EnsureReportSheet SH_DEL_WITH_BAL
    EnsureReportSheet SH_PAID_PKGS
    EnsureReportSheet SH_PENDING_PKGS
    EnsureReportSheet SH_ADV_PAY
    EnsureReportSheet SH_ADJUST
    EnsureReportSheet SH_OUTSTANDING
    
    ' 2. Clear old data (keep headers if any)
    ClearSheetData SH_PKG_ORDERS
    ClearSheetData SH_PAID_ORDERS
    ClearSheetData SH_BAL_PENDING
    ClearSheetData SH_PEND_DEL
    ClearSheetData SH_DEL_NO_PAY
    ClearSheetData SH_DEL_WITH_BAL
    ClearSheetData SH_PAID_PKGS
    ClearSheetData SH_PENDING_PKGS
    ClearSheetData SH_ADV_PAY
    ClearSheetData SH_ADJUST
    ClearSheetData SH_OUTSTANDING
    
    ' 3. Build each report
    Build_PackageOrders
    Build_PaidOrders
    Build_BalancePendingOrders
    Build_PendingDeliveryOrders
    Build_DeliveredWithoutPayment
    Build_DeliveredWithBalance
    Build_PaidPackages
    Build_PendingPackages
    Build_AdvancePayments
    Build_TotalAdjustments
    Build_CustomerOutstandingSummary
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'==========================================
' HELPER: Ensure sheet exists, create if not
'==========================================
Private Sub EnsureReportSheet(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheetName
    End If
End Sub

'==========================================
' HELPER: Clear data but keep header row (row 1)
'==========================================
Private Sub ClearSheetData(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws.UsedRange.Rows.Count > 1 Then
        ws.Range("A2:" & ws.Cells(ws.Rows.Count, ws.Columns.Count).Address).ClearContents
    End If
End Sub

'==========================================
' GET ORDERS SHEET + LAST ROW
'==========================================
Private Function OrdersSheet() As Worksheet
    Set OrdersSheet = ThisWorkbook.Worksheets(SH_ORDERS)
End Function

Private Function LastRowOrders() As Long
    Dim ws As Worksheet
    Set ws = OrdersSheet
    
    LastRowOrders = ws.Cells(ws.Rows.Count, COL_ORDERNO).End(xlUp).Row
    If LastRowOrders < 4 Then LastRowOrders = 3
End Function

'==========================================
' COPY ROW FROM ORDERS TO TARGET SHEET
' (Copies full row 3 header first if needed)
'==========================================
Private Sub CopyOrderRow(ByVal sourceRow As Long, ByVal targetSheetName As String, ByRef targetRow As Long)
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Set wsSrc = OrdersSheet
    Set wsTgt = ThisWorkbook.Worksheets(targetSheetName)
    
    ' If targetRow = 1, create header row by copying row 3 from orders
    If targetRow = 1 Then
        wsTgt.Rows(1).Value = wsSrc.Rows(3).Value
        targetRow = 2
    End If
    
    wsTgt.Rows(targetRow).Value = wsSrc.Rows(sourceRow).Value
    targetRow = targetRow + 1
End Sub

'==========================================
' 1. Package Orders (Package = "Yes")
'==========================================
Private Sub Build_PackageOrders()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim tgtRow As Long
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        If OrdersSheet.Cells(r, COL_PACKAGE).Value = "Yes" Then
            CopyOrderRow r, SH_PKG_ORDERS, tgtRow
        End If
    Next r
End Sub

'==========================================
' 2. Paid Orders (Balance <= 0, not Cancelled)
'==========================================
Private Sub Build_PaidOrders()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim bal As Variant, status As String
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        bal = OrdersSheet.Cells(r, COL_BALANCE).Value
        status = CStr(OrdersSheet.Cells(r, COL_ORDERSTATUS).Value)
        
        If status <> "Cancelled" Then
            If IsNumeric(bal) Then
                If bal <= 0 Then
                    CopyOrderRow r, SH_PAID_ORDERS, tgtRow
                End If
            End If
        End If
    Next r
End Sub

'==========================================
' 3. Balance Pending Orders (Balance > 0, not Cancelled)
'==========================================
Private Sub Build_BalancePendingOrders()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim bal As Variant, status As String
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        bal = OrdersSheet.Cells(r, COL_BALANCE).Value
        status = CStr(OrdersSheet.Cells(r, COL_ORDERSTATUS).Value)
        
        If status <> "Cancelled" Then
            If IsNumeric(bal) Then
                If bal > 0 Then
                    CopyOrderRow r, SH_BAL_PENDING, tgtRow
                End If
            End If
        End If
    Next r
End Sub

'==========================================
' 4. Pending Delivery Orders (Delivered On blank, not Cancelled)
'==========================================
Private Sub Build_PendingDeliveryOrders()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim status As String
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        status = CStr(OrdersSheet.Cells(r, COL_ORDERSTATUS).Value)
        If status <> "Cancelled" Then
            If IsEmpty(OrdersSheet.Cells(r, COL_DELIVEREDON).Value) _
               Or OrdersSheet.Cells(r, COL_DELIVEREDON).Value = "" Then
                CopyOrderRow r, SH_PEND_DEL, tgtRow
            End If
        End If
    Next r
End Sub

'==========================================
' 5. Delivered Without Payment (Delivered, Paid=0)
'==========================================
Private Sub Build_DeliveredWithoutPayment()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim paid As Variant
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        If Not IsEmpty(OrdersSheet.Cells(r, COL_DELIVEREDON).Value) Then
            paid = OrdersSheet.Cells(r, COL_PAID).Value
            If (Not IsNumeric(paid)) Or paid = 0 Then
                CopyOrderRow r, SH_DEL_NO_PAY, tgtRow
            End If
        End If
    Next r
End Sub

'==========================================
' 6. Delivered With Balance (Delivered, Balance > 0)
'==========================================
Private Sub Build_DeliveredWithBalance()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim bal As Variant
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        If Not IsEmpty(OrdersSheet.Cells(r, COL_DELIVEREDON).Value) Then
            bal = OrdersSheet.Cells(r, COL_BALANCE).Value
            If IsNumeric(bal) Then
                If bal > 0 Then
                    CopyOrderRow r, SH_DEL_WITH_BAL, tgtRow
                End If
            End If
        End If
    Next r
End Sub

'==========================================
' 7. Paid Packages (Payment Status = "Paid")
'==========================================
Private Sub Build_PaidPackages()
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim lastRow As Long, r As Long, tgtRow As Long
    
    Set wsSrc = ThisWorkbook.Worksheets(SH_PACKAGES)
    Set wsTgt = ThisWorkbook.Worksheets(SH_PAID_PKGS)
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    tgtRow = 1
    
    ' Header
    wsTgt.Rows(1).Value = wsSrc.Rows(1).Value
    tgtRow = 2
    
    For r = 2 To lastRow
        If wsSrc.Cells(r, PKG_COL_STATUS).Value = "Paid" Then
            wsTgt.Rows(tgtRow).Value = wsSrc.Rows(r).Value
            tgtRow = tgtRow + 1
        End If
    Next r
End Sub

'==========================================
' 8. Pending Packages (Payment Status = "Pending")
'==========================================
Private Sub Build_PendingPackages()
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim lastRow As Long, r As Long, tgtRow As Long
    
    Set wsSrc = ThisWorkbook.Worksheets(SH_PACKAGES)
    Set wsTgt = ThisWorkbook.Worksheets(SH_PENDING_PKGS)
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    tgtRow = 1
    
    ' Header
    wsTgt.Rows(1).Value = wsSrc.Rows(1).Value
    tgtRow = 2
    
    For r = 2 To lastRow
        If wsSrc.Cells(r, PKG_COL_STATUS).Value = "Pending" Then
            wsTgt.Rows(tgtRow).Value = wsSrc.Rows(r).Value
            tgtRow = tgtRow + 1
        End If
    Next r
End Sub

'==========================================
' 9. Advance Payments Details (Is Advance? = "Yes")
'==========================================
Private Sub Build_AdvancePayments()
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim lastRow As Long, r As Long, tgtRow As Long
    
    Set wsSrc = ThisWorkbook.Worksheets(SH_PAYMENTS)
    Set wsTgt = ThisWorkbook.Worksheets(SH_ADV_PAY)
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    
    ' Header
    wsTgt.Rows(1).Value = wsSrc.Rows(1).Value
    tgtRow = 2
    
    For r = 2 To lastRow
        If UCase(wsSrc.Cells(r, PAY_COL_ISADV).Value) = "YES" Then
            wsTgt.Rows(tgtRow).Value = wsSrc.Rows(r).Value
            tgtRow = tgtRow + 1
        End If
    Next r
End Sub

'==========================================
' 10. Total Adjustments (Adjustment <> 0)
'==========================================
Private Sub Build_TotalAdjustments()
    Dim r As Long, lastRow As Long, tgtRow As Long
    Dim adj As Variant
    
    lastRow = LastRowOrders()
    tgtRow = 1
    
    For r = 4 To lastRow
        adj = OrdersSheet.Cells(r, COL_ADJUSTMENT).Value
        If IsNumeric(adj) Then
            If adj <> 0 Then
                CopyOrderRow r, SH_ADJUST, tgtRow
            End If
        End If
    Next r
End Sub

'==========================================
' 11. Customer Outstanding Summary
'     Group by Name, Balance>0
'==========================================
Private Sub Build_CustomerOutstandingSummary()
    Dim wsTgt As Worksheet
    Dim lastRow As Long, r As Long
    Dim dict As Object
    Dim custName As String
    Dim bal As Variant
    
    Set wsTgt = ThisWorkbook.Worksheets(SH_OUTSTANDING)
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = LastRowOrders()
    
    ' Build dictionary: key = customer name, value = array(TotalBalance, CountOrders)
    For r = 4 To lastRow
        custName = Trim(CStr(OrdersSheet.Cells(r, COL_NAME).Value))
        bal = OrdersSheet.Cells(r, COL_BALANCE).Value
        
        If custName <> "" Then
            If IsNumeric(bal) Then
                If bal > 0 Then
                    If Not dict.Exists(custName) Then
                        dict.Add custName, Array(CDbl(bal), 1)
                    Else
                        Dim arr
                        arr = dict(custName)
                        arr(0) = arr(0) + CDbl(bal)
                        arr(1) = arr(1) + 1
                        dict(custName) = arr
                    End If
                End If
            End If
        End If
    Next r
    
    ' Write headers
    wsTgt.Range("A1").Value = "Customer Name"
    wsTgt.Range("B1").Value = "Total Outstanding Amount"
    wsTgt.Range("C1").Value = "Number of Unpaid Orders"
    
    ' Dump dictionary to sheet
    Dim i As Long, rowOut As Long
    rowOut = 2
    For i = 0 To dict.Count - 1
        custName = dict.Keys()(i)
        Dim vals
        vals = dict.Items()(i)
        
        wsTgt.Cells(rowOut, 1).Value = custName
        wsTgt.Cells(rowOut, 2).Value = vals(0)
        wsTgt.Cells(rowOut, 3).Value = vals(1)
        
        rowOut = rowOut + 1
    Next i
    
    ' Sort by Total Outstanding (col B) descending
    If rowOut > 2 Then
        wsTgt.Range("A1:C" & rowOut - 1).Sort _
            Key1:=wsTgt.Range("B2"), Order1:=xlDescending, Header:=xlYes
    End If
End Sub
