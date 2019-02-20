Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

Dim ws As Worksheet


For Each ws In worksheets

  Dim ticker_Name As String

  Dim vol_Total As Double
  
  vol_Total = 0
  
  Dim I As Double
  
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  For I = 2 To lastrow

    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ticker_Name = Cells(I, 1).Value

      vol_Total = vol_Total + Cells(I, 7).Value

      Cells(Summary_Table_Row, 9).Value = ticker_Name

      Cells(Summary_Table_Row, 10).Value = vol_Total

      Summary_Table_Row = Summary_Table_Row + 1

      vol_Total = 0

    Else

      vol_Total = vol_Total + Cells(I, 7).Value

    End If

  Next I
Next ws


End Sub

