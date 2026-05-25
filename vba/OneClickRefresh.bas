Attribute VB_Name = "Module1"
Sub OneClickRefresh()
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Olist pipeline..."
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    ' Refresh all Power Query connections
    ThisWorkbook.RefreshAll

    ' Wait for queries to finish processing
    Application.Wait Now + TimeValue("00:00:05")
    DoEvents

    ' Refresh all PivotTables on every sheet
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    ' Write to the audit log
    Dim logWs As Worksheet
    Set logWs = ThisWorkbook.Sheets("Audit_Log")
    Dim nr As Long
    nr = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    logWs.Cells(nr, 1).Value = Now()
    logWs.Cells(nr, 2).Value = Environ("USERNAME")
    logWs.Cells(nr, 3).Value = _
        ThisWorkbook.Sheets("Reconciled_Output") _
        .ListObjects(1).ListRows.Count

    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True

    MsgBox "Pipeline refreshed successfully." & vbNewLine & _
           "Timestamp: " & Format(Now(), "dd/mm/yyyy hh:mm:ss"), _
           vbInformation, "Olist Reconciliation"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    MsgBox "Refresh failed: " & Err.Description, vbCritical, "Error"
End Sub

