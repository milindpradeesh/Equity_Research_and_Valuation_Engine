Attribute VB_Name = "Module4"
Sub Build_Cashflow_Statement()

    Dim rawSht As Worksheet, modelSht As Worksheet
    
    Set rawSht = Sheets("Raw_CashFlow")
    
    On Error Resume Next
    Set modelSht = Sheets("Cash Flow Statement")
    If modelSht Is Nothing Then
        Set modelSht = Sheets.Add(After:=rawSht)
        modelSht.Name = "Cash Flow Statement"
    Else
        modelSht.Cells.Clear
    End If
    On Error GoTo 0

    rawSht.UsedRange.Copy
    modelSht.Range("A3").PasteSpecial xlPasteValues

    modelSht.Range("A1").Value = "CASH FLOW STATEMENT"
    modelSht.Range("A1").Font.Size = 16
    modelSht.Range("A1").Font.Bold = True

    modelSht.Rows(3).Font.Bold = True
    modelSht.Columns.AutoFit
    modelSht.Range("B4:Z200").NumberFormat = "#,##0;(#,##0)"

End Sub

