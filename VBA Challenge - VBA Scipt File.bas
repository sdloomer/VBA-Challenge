Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data():

' Set variable for Worksheets
Dim ws As Worksheet

' Begin loop for all worksheets
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

    ' Set variable for Ticker Name
    Dim Ticker As String
    Dim i As Long
    
    ' Set variable for holding Total Stock per Ticker
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    ' Set variable for holding Yearly Change per Ticker
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    
    ' Set variable for holding Percent Change per Ticker
    Dim Percent_Change As Double
    Percent_Change = 0
    
    ' Keep track of location of each Ticker in Table
    Dim Table_Row As Integer
    Table_Row = 2
    
    ' Determine Last Row in Complete List of Tickers
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Loop through all Tickers
        For i = 2 To LastRow
    
            ' Check to make sure still within same Ticker, if not
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
                ' Set Ticker Name
                Ticker = Cells(i, 1).Value
                
                ' Find Yearly Change
                Yearly_Change = Cells(i, 6).Value - Open_Price
                
                ' Find Percent Change
                Percent_Change = (Yearly_Change / Open_Price)
                
                ' Change to Next Ticker's Open Price
                Open_Price = Cells(i + 1, 3).Value
                
                ' Add to Stock Total
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
        
                ' Print Ticker Name in Table
                Range("I" & Table_Row).Value = Ticker
                
                ' Print Yearly Change in Table
                Range("J" & Table_Row).Value = Yearly_Change
                
                ' Print Percent Change in Table
                Range("K" & Table_Row).Value = Format(Percent_Change, "0.00%")
                
                ' Print Stock Total to Table
                Range("L" & Table_Row).Value = Stock_Volume
                
                ' Color Positive and Negative Yearly Change
                If Range("J" & Table_Row).Value < 0 Then
                    
                    Range("J" & Table_Row).Interior.ColorIndex = 3
                        
                End If
                
                If Range("J" & Table_Row).Value > 0 Then
                    
                    Range("J" & Table_Row).Interior.ColorIndex = 4
                
                End If
                
                ' Color Positive and Negative Percent Change
                If Range("K" & Table_Row).Value < 0 Then
                    
                    Range("K" & Table_Row).Interior.ColorIndex = 3
                        
                End If
                
                If Range("K" & Table_Row).Value > 0 Then
                    
                    Range("K" & Table_Row).Interior.ColorIndex = 4
                
                End If
                
                ' Add one to Table Row
                Table_Row = Table_Row + 1
                
                ' Reset Stock Total
                Stock_Volume = 0
                
            Else
            
                ' Add to Stock Total
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
                
            End If
            
        Next i
        
    ' Set variables for Greatest % Increase, % Decrease, Greatest Total Volume
    Dim Max As Double
    Dim Largest_Increase_Ticker As String
    Max = Cells(2, 11).Value
    Dim Min As Double
    Dim Largest_Decrease_Ticker As String
    Min = Cells(2, 11).Value
    Dim Vol As Double
    Dim Greatest_Volume As String
    Vol = Cells(2, 12).Value
    
    Dim j As Long
      
    ' Keep track of location of each Ticker in Second Table
    Dim Table_Row2 As Integer
    Table_Row2 = 2
    
    ' Determine Last Row in Second List of Tickers
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Loop through all Percent Changes
        For j = 2 To LastRow
            
            ' Find Greatest % Increase
            If Cells(j, 11).Value > Max Then
                
                ' Hold Greatest % Increase
                Largest_Increase_Ticker = Cells(j, 9).Value
                
                ' Set Greatest % Increase
                Max = Cells(j, 11).Value
                
            End If
            
            ' Find Greatest % Decrease
            If Cells(j, 11).Value < Min Then
                
                ' Hold Greatest % Decrease
                Largest_Decrease_Ticker = Cells(j, 9).Value
                
                ' Set Greatest % Decrease
                Min = Cells(j, 11).Value
            
            End If
            
            ' Find Greatest Total Volume
            If Cells(j, 12).Value > Vol Then
                
                ' Hold Greatest Total Volume
                Greatest_Volume = Cells(j, 9).Value
                
                ' Set Greatest Total Volume
                Vol = Cells(j, 12).Value
            
            End If
                
        Next j
        
    ' Print Greatest % Increase
        Cells(2, 17).Value = Max
    
    ' Print Greatest % Increase Ticker
        Cells(2, 16).Value = Largest_Increase_Ticker
    
    ' Print Greatest % Decrease
        Cells(3, 17).Value = Min
    
    ' Print Greatest % Decrease Ticker
        Cells(3, 16).Value = Largest_Decrease_Ticker
        
    ' Print Greatest Total Volume
        Cells(4, 17).Value = Vol
    
    ' Print Greatest Total Volume Ticker
        Cells(4, 16).Value = Greatest_Volume
        
    ' Print Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
    
    ' Autofit columns
        ws.Columns("A:Q").AutoFit
    
Next ws
    
End Sub

' Sources used
' https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475


