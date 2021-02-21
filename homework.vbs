Sub Sum_Ticker()

For Each ws In Worksheets

  ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String
  
  ' Set Headers For Summary Table
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    
  ' Set an initial variable for holding the last row of data
    Dim Last_Row As Long
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
  

  ' Set an initial variable for holding the total ticker volume
    Dim Ticker_Volume_Total As Double

  ' Set an initial variable for holding the Yearly Change
    Dim Ticker_Yearly_Change As Double
  
  ' Set an initial variable for holding the Yearly Change
    Dim Ticker_Percent_Change As Double
  
    ' Set an initial variable for holding the Opening_Yearly_Price
    Dim Opening_Yearly_Price As Double

    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    
    ' Set variable to check Percent Change & Total Volume Rows for Min and Max Values
    Dim MaxMin_PercentChange_RowCheck As Long
    MaxMin_PercentChange_RowCheck = ws.Cells(Rows.Count, 12).End(xlUp).Row
    Dim MaxVolume_RowCheck As Long
    MaxVolume_RowCheck = ws.Cells(Rows.Count, 13).End(xlUp).Row
    Dim PercentMin As Double
    PercentMin = 0
    Dim PercentMax As Double
    PercentMax = 0
    Dim MaxVolume As Double
    MaxVolume = 0
  
  ' Set initial values for Ticker_Volume_Total and Summary_Table_Row
    Ticker_Volume_Total = 0
    Ticker_Yearly_Change = 0
    Summary_Table_Row = 2

    ' Store Open Ticker Price
    Opening_Yearly_Price = ws.Cells(2, 3).Value

  ' Loop through all ticker data
  For i = 2 To Last_Row


    ' Check if we are still within the same Ticker name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
        Ticker_Name = ws.Cells(i, 1).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
      
      ' Calculate Yearly Change value
      Ticker_Yearly_Change = ws.Cells(i, 6).Value - Opening_Yearly_Price
      
      ' Print the ticker Yearly Change
      ws.Range("K" & Summary_Table_Row).Value = Ticker_Yearly_Change
      
      ' Calculate Percent Change for the year
      
      If Opening_Yearly_Price = 0 Then
      Ticker_Percent_Change = 0
      Else
      Ticker_Percent_Change = ((Ticker_Yearly_Change / Opening_Yearly_Price))
      End If
        
      
      ' Print the Percent Change for the Year
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Percent_Change
      ws.Range("L" & Summary_Table_Row).Style = "Percent"
      ' ws.Range("L" & Summary_Table_Row).Value = Format(ws.Range("L" & Summary_Table_Row).Value / 100.00, "#.##%")
    
     ' Add to the Ticker Volume Total
      Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

      ' Print the Ticker Volume Total to the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Ticker_Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Summary Table Variables
      Ticker_Yearly_Change = 0
      Ticker_Volume_Total = 0
      Opening_Yearly_Price = ws.Cells(i + 1, 3).Value
      
      Else
      
        Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
  Dim Percentage_Change_Row As Long
  Yearly_Change_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  For i = 2 To Yearly_Change_Row
    If ws.Cells(i, 11).Value >= 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 11).Interior.ColorIndex = 3
    End If
    
    Next i
    
    For i = 2 To MaxMin_PercentChange_RowCheck
    If PercentMax < ws.Cells(i, 12).Value Then
        PercentMax = ws.Cells(i, 12).Value
        ws.Cells(2, 18).Value = PercentMax
        ws.Cells(2, 18).Style = "Percent"
        ' Pull ticker name from summary table
        ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
    
    ElseIf percent_min > ws.Cells(i, 12).Value Then
        PercentMin = ws.Cells(i, 12).Value
        ws.Cells(3, 18).Value = PercentMin
        ws.Cells(3, 18).Style = "Percent"
        ' Pull ticker name from summary table
        ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
        
    End If
Next i

For i = 2 To MaxVolume_RowCheck
        If MaxVolume < ws.Cells(i, 13).Value Then
            MaxVolume = ws.Cells(i, 13).Value
            ws.Cells(4, 18).Value = MaxVolume
            ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
    End If
Next i

    
Next ws

End Sub