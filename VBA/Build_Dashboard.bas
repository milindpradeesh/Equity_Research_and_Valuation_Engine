Attribute VB_Name = "Module5"
Sub Build_Dashboard()

    Dim rawSht As Worksheet, dash As Worksheet
    
    Set rawSht = Sheets("Raw_Quote")
    
    On Error Resume Next
    Set dash = Sheets("Dashboard")
    If dash Is Nothing Then
        Set dash = Sheets.Add(Before:=Sheets(1))
        dash.Name = "Dashboard"
    Else
        dash.Cells.Clear
    End If
    On Error GoTo 0

    dash.Range("A1").Value = "COMPANY DASHBOARD"
    dash.Range("A1").Font.Size = 18
    dash.Range("A1").Font.Bold = True

    rawSht.Range("A1:B20").Copy dash.Range("A3")
    dash.Columns("A:B").AutoFit

End Sub

