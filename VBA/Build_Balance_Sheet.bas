Attribute VB_Name = "Module6"
Sub Build_Balance_Sheet()

    Dim rawSht As Worksheet, modelSht As Worksheet
    
    Set rawSht = Sheets("Raw_Balance")
    
    On Error Resume Next
    Set modelSht = Sheets("Balance Sheet")
    If modelSht Is Nothing Then
        Set modelSht = Sheets.Add(After:=rawSht)
        modelSht.Name = "Balance Sheet"
    Else
        modelSht.Cells.Clear
    End If
    On Error GoTo 0

    rawSht.UsedRange.Copy
    modelSht.Range("A3").PasteSpecial xlPasteValues

    modelSht.Range("A1").Value = "BALANCE SHEET"
    modelSht.Range("A1").Font.Size = 16
    modelSht.Range("A1").Font.Bold = True

    modelSht.Rows(3).Font.Bold = True
    modelSht.Columns.AutoFit
    modelSht.Range("B4:Z200").NumberFormat = "#,##0;(#,##0)"

End Sub

