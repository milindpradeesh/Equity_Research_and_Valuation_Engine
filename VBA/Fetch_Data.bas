Attribute VB_Name = "Module28"
Sub Fetch_Data()

    Dim ticker As String
    Dim pythonPath As String
    Dim scriptPath As String
    Dim excelPath As String
    Dim cmd As String
    Dim ws As Worksheet
    Dim wsh As Object
    Dim exec As Object
    Dim sht As Worksheet
    Dim statusCell As Range

    Set ws = ThisWorkbook.Sheets("Model Control")
    Set statusCell = ws.Range("B4") ' Progress display cell
    
    ticker = ws.Range("B3").Value
    
    If ticker = "" Then
        MsgBox "Please enter a ticker symbol in Control Panel!B4", vbExclamation
        Exit Sub
    End If

    statusCell.Value = "Fetching data from API..."
    DoEvents

    pythonPath = "<>"
    scriptPath = "<>"
    excelPath = ThisWorkbook.FullName

    cmd = """" & pythonPath & """ -u """ & scriptPath & """ """ & ticker & """ """ & excelPath & """"

    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec(cmd)

    ' Wait until Python finishes while keeping Excel responsive
    Do While exec.Status = 0
        statusCell.Value = "Fetching & processing financial data..."
        DoEvents
    Loop

    ' Check Python exit code
    If exec.ExitCode <> 0 Then
        statusCell.Value = "? Python script failed. Check API/Data."
        MsgBox "Python process failed. Please check API key, internet, or script.", vbCritical
        Exit Sub
    End If

    statusCell.Value = "Data fetched. Formatting sheets..."
    DoEvents

    ' ===== FORMAT RAW SHEETS =====
    Call Format_Raw_Sheets

    ' ===== FORMAT TIMESTAMP IN RAW_QUOTE =====
    Dim quoteSht As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error Resume Next
    Set quoteSht = ThisWorkbook.Sheets("Raw_Quote")
    On Error GoTo 0

    If Not quoteSht Is Nothing Then
        lastRow = quoteSht.Cells(quoteSht.Rows.Count, 1).End(xlUp).Row
        
        For i = 1 To lastRow
            If LCase(quoteSht.Cells(i, 1).Value) = "timestamp" Then
                quoteSht.Cells(i, 2).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
                Exit For
            End If
        Next i
    End If

    statusCell.Value = "Building financial model..."
    DoEvents

    ' ===== BUILD MODEL =====
    Call Build_Income_Statement
    Call Build_Balance_Sheet
    Call Build_Cashflow_Statement
    Call Build_Dashboard

    ' ===== HIDE RAW DATA SHEETS =====
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = "Raw_Income" _
        Or sht.Name = "Raw_Balance" _
        Or sht.Name = "Raw_CashFlow" _
        Or sht.Name = "Raw_Quote" Then
            sht.Visible = xlSheetHidden
        End If
    Next sht

    statusCell.Value = "Ready!"
    MsgBox "Ready"

End Sub




