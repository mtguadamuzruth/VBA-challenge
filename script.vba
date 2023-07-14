Sub stock_executable()
'establish where we are going to store ticker yearly change, % change. total stock Volume
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

' Set an initial variable for ticker
Dim Ticker As String
Dim LastRow As Variant

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'variable "Tindex stands for Ticker Index
Dim TIndex As Integer
Dim OpenVal As Long
Dim TotalSV As Double

TIndex = 2
OpenVal = 0
TotalSV = 0

'loop though all ticker
For i = 2 To LastRow
    TotalSV = TotalSV + Cells(i, "G").Value
    
    If OpenVal = 0 Then
        OpenVal = Cells(i, "C").Value
    End If


    ' check if we are still using the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    'add ticker symbol
        Cells(TIndex, "j").Value = Cells(i, "a").Value
        'add yearly change from opening price at the diginning of the given year to the closing price at the end of the year
        YearlyCh = Cells(i, "F").Value - OpenVal
         Cells(TIndex, "K").Value = YearlyCh
         
         'add the percentage change note we do not *100 because the L is formatted in percentages
         Cells(TIndex, "L").Value = (YearlyCh / OpenVal)
         'add the Total total Stock Value
         
         Cells(TIndex, "M").Value = TotalSV
         
        TIndex = TIndex + 1
        OpenVal = 0
        TotalSV = 0
    End If

Next i

End Sub
