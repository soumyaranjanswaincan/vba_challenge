Attribute VB_Name = "Module2"
Sub Create_stock_summary(sheetname As String)

    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim totalVolume As Double
        
    Dim skipOpenPrice As Boolean
      
    Dim maxVal As Double
    Dim minVal As Double
    Dim maxVolume As Double
    Dim columnRange As Range
    Dim searchStockChangeVal As String
    Dim searchStockTotalVal As String
    
    Dim ws As Worksheet
    ' Referencing worksheet by name
    Set ws = ThisWorkbook.Worksheets(sheetname)
    
    ' Activate the sheet
    ws.Activate
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Initialize values
    totalVolume = 0
    skipOpenPrice = False
    
    
    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    For i = 2 To lastrow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker value
      ticker = Cells(i, 1).Value
      
      ' Set the close price value
      close_price = Cells(i, 6).Value

      ' Add to the Total
      totalVolume = totalVolume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker

      ' Print the total volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = totalVolume
      
      yearly_change = (close_price - open_price)
      percent_change = (yearly_change / open_price)
      
       ' Print the ticker in the Summary Table
      Range("J" & Summary_Table_Row).Value = yearly_change

      ' Print the total volume to the Summary Table
      Range("K" & Summary_Table_Row).Value = percent_change
      
      If yearly_change > 0 Then
         Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      totalVolume = 0
      skipOpenPrice = False

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      totalVolume = totalVolume + Cells(i, 7).Value
      
      If skipOpenPrice = False Then
            
        ' Set the open price value
        open_price = Cells(i, 3).Value
        
        skipOpenPrice = True
        
      End If
        
      
    End If

  Next i

  
      ' Set the range for the column where you want to find the maximum value
      Set columnRange = Range("K2:K" & Cells(ws.Rows.Count, "K").End(xlUp).Row)
      
      ' Set the range for the column where you want to find the maximum value
      Set totalColumnRange = Range("L2:L" & Cells(ws.Rows.Count, "L").End(xlUp).Row)
      
      ' Find the maximum value in the column
      maxVal = Application.WorksheetFunction.Max(columnRange)
      
      ' Find the minimum value in the column
      minVal = Application.WorksheetFunction.Min(columnRange)
      
      ' Find the minimum value in the column
      maxVolume = Application.WorksheetFunction.Max(totalColumnRange)

      
      Range("P2").Value = maxVal
      Range("P3").Value = minVal
      Range("P4").Value = maxVolume
      
        For i = 2 To columnRange.Rows.Count
    
            searchStockChangeVal = ws.Cells(i, 11).Value
            searchStockTotalVal = ws.Cells(i, 12).Value
        
            If (CStr(searchStockChangeVal) = CStr(maxVal)) Then
                Cells(2, 15).Value = Cells(i, 9).Value
            Else

            End If
            If (CStr(searchStockChangeVal) = CStr(minVal)) Then
                Cells(3, 15).Value = Cells(i, 9).Value
            Else

            End If
            
            If (CStr(searchStockTotalVal) = CStr(maxVolume)) Then
                Cells(4, 15).Value = Cells(i, 9).Value
            Else

            End If
        
        Next i
      

End Sub


