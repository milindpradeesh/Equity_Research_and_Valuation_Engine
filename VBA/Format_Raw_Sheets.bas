Attribute VB_Name = "Module2"
Sub Format_Raw_Sheets()

    Dim sht As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    For Each sht In ThisWorkbook.Worksheets
        If Left(sht.Name, 4) = "Raw_" Then
            
            With sht
                .Cells.Font.Name = "Calibri"
                
                ' Header row formatting
                .Rows(1).Font.Bold = True
                .Rows(1).HorizontalAlignment = xlCenter
                
                ' Auto fit columns
                .Columns.AutoFit
                
                ' Find last used row and column
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                
                ' Apply number format only if data exists beyond column A
                If lastCol > 1 And lastRow > 1 Then
                    .Range(.Cells(2, 2), .Cells(lastRow, lastCol)).NumberFormat = "#,##0;(#,##0)"
                End If
                
                ' Freeze panes safely
                .Activate
                .Range("B2").Select
                ActiveWindow.FreezePanes = True
            End With
            
        End If
    Next sht

End Sub

