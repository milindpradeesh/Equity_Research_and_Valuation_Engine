Attribute VB_Name = "Module3"
Sub Build_Income_Statement()

    Dim rawSht As Worksheet
    Dim modelSht As Worksheet

    ' --- Find Raw_Income sheet safely ---
    On Error Resume Next
    Set rawSht = ThisWorkbook.Worksheets("Raw_Income")
    On Error GoTo 0

    If rawSht Is Nothing Then
        MsgBox "Sheet 'Raw_Income' not found. Python may not have written the data.", vbCritical
        Exit Sub
    End If

    ' --- Create or clear Income Statement sheet ---
    On Error Resume Next
    Set modelSht = ThisWorkbook.Worksheets("Income Statement")
    On Error GoTo 0

    If modelSht Is Nothing Then
        Set modelSht = ThisWorkbook.Worksheets.Add(After:=rawSht)
        modelSht.Name = "Income Statement"
    Else
        modelSht.Cells.Clear
    End If

    ' --- Copy values from Raw to Model ---
    rawSht.UsedRange.Copy
    modelSht.Range("A3").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' --- Formatting ---
    modelSht.Range("A1").Value = "INCOME STATEMENT"
    modelSht.Range("A1").Font.Size = 16
    modelSht.Range("A1").Font.Bold = True

    modelSht.Rows(3).Font.Bold = True
    modelSht.Columns.AutoFit
    modelSht.Range("B4:Z200").NumberFormat = "#,##0;(#,##0)"

End Sub

