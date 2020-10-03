Attribute VB_Name = "Module1"
Sub Ticker_Value()
    With ActiveSheet
        .Range("A2", .Range("A2").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("I2"), Unique:=True
        .Range("I2").Delete Shift:=xlShiftUp
    End With

End Sub

Sub Yearly_Change()
  With ActiveSheet
        Dim x As Integer
        Dim open_no As Variant
        Dim result As Variant
        NumRows = Range("I2", Range("I2").End(xlDown)).Rows.Count
        Range("I2").Select
        For x = 1 To NumRows
            open_no = Application.VLookup(Cells((x + 1), 9), ActiveSheet.Range("A2:C1000000"), 3, False)
            close_no = Application.VLookup(Cells((x + 1), 9), ActiveSheet.Range("A2:F1000000"), 6, True)
            close_no = CVar(close_no)
            result = close_no - open_no
            Cells((x + 1), 10).Value = result
            If open_no = 0 Then
                perc_ch = "N/A"
            Else
                perc_ch = result / open_no
            End If
            Cells((x + 1), 11).Value = perc_ch
            Cells((x + 1), 11).NumberFormat = "0.00%"
            ActiveCell.Offset(1, 0).Select
        Next
    End With
End Sub

Sub total_count()
Dim SumIfResult As Variant
    Dim x As Integer
    With ActiveSheet
        NumRows = Range("I2", Range("I2").End(xlDown)).Rows.Count
        Range("I2").Select
        For x = 1 To NumRows
            SumIfResult = Application.WorksheetFunction.SumIf(Range("A2:A1000000"), Cells((x + 1), 9), Range("G2:G1000000"))
            Cells((x + 1), 12).Value = SumIfResult
            ActiveCell.Offset(1, 0).Select
        Next
    End With
End Sub
