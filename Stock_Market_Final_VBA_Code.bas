Attribute VB_Name = "Module1"
Sub Total_Stock_Volume():

'Set Headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 13).Value = "Total Stock Volume"

'Set Variables'
    Dim Ticker_Symbol As String
    Dim Total_Volume As Double
        Total_Volume = 0
 
'Summary Row Table'
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
'Being able to loop till the last row'
    Dim LastRow As Long
        LastRow = Range("A1", Range("A1").End(xlDown)).Rows.Count

'Loop through the sheet'
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Identify the Ticker Symbol'
            Ticker_Symbol = Cells(i, 1).Value
            
            'Add to Total Volume'
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            'Print Ticker Symbol in Summary Table'
            Range("J" & Summary_Table_Row).Value = Ticker_Symbol
            
            'Print Total Volume in Summary Table'
            Range("M" & Summary_Table_Row).Value = Total_Volume
            
            'New row to Summary Table'
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Total Volume'
            Total_Volume = 0
        
        'If the cell is the same'
        Else
            
            'Add to Total Volume'
            Total_Volume = Total_Volume + Cells(i, 7).Value
        
        End If
    
    Next i
End Sub

Sub Yearly_Change():

'Set Headers
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    
'Set Variables
    Dim Open_Stock As Double
        Open_Stock = Cells(2, 3).Value
    Dim Closing_Stock As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double


'Summary Row Table'
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

'Being able to loop till the last row'
    Dim LastRow As Long
        LastRow = Range("A1", Range("A1").End(xlDown)).Rows.Count
        
'Loop through, finding yearly close and yearly open'
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Define the value for Closing_Stock, Yearly_Change, Percent_Change'
            Closing_Stock = Cells(i, 6).Value
            Yearly_Change = Closing_Stock - Open_Stock
            Percent_Change = (Yearly_Change / Open_Stock)
            
            'Print Yearly Change Value'
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            
            'Print Percent Change Value'
            Range("L" & Summary_Table_Row).Value = Percent_Change
            Cells(Summary_Table_Row, 12).Value = Format(Cells(Summary_Table_Row, 12).Value, "Percent")
            
            'New row to Summary Table'
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Values'
            Open_Stock = Cells(i + 1, 3).Value
            Closing_Stock = 0
            Yearly_Change = 0
        
        End If
        
    Next i
     
'Change colors of yearly change
    For i = 2 To LastRow
    
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        Else
            Cells(i, 11).Interior.ColorIndex = 3
        
        End If
    
    Next i
        
    
End Sub

Sub Combined_Process():

'Set Headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
'Set Variables'
    Dim Ticker_Symbol As String
    Dim Total_Volume As Double
        Total_Volume = 0
    Dim Open_Stock As Double
        Open_Stock = Cells(2, 3).Value
    Dim Closing_Stock As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
'Summary Row Table'
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
'Being able to loop till the last row'
    Dim LastRow As Long
        LastRow = Range("A1", Range("A1").End(xlDown)).Rows.Count

'Loop through the sheet'
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Identify the Ticker Symbol'
            Ticker_Symbol = Cells(i, 1).Value
            
            'Add to Total Volume'
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            'Print Ticker Symbol in Summary Table'
            Range("J" & Summary_Table_Row).Value = Ticker_Symbol
            
            'Print Total Volume in Summary Table'
            Range("M" & Summary_Table_Row).Value = Total_Volume
            
            'Reset Total Volume'
            Total_Volume = 0
            
            'Define the value for Closing_Stock, Yearly_Change, Percent_Change'
            Closing_Stock = Cells(i, 6).Value
            Yearly_Change = Closing_Stock - Open_Stock
                If Open_Stock = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = (Yearly_Change / Open_Stock)
                End If
            
            'Print Yearly Change Value'
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            
            'Print Percent Change Value'
            Range("L" & Summary_Table_Row).Value = Percent_Change
            Cells(Summary_Table_Row, 12).Value = Format(Cells(Summary_Table_Row, 12).Value, "Percent")
            
            'New row to Summary Table'
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Values'
            Open_Stock = Cells(i + 1, 3).Value
            Closing_Stock = 0
            Yearly_Change = 0
        
        
        'If the cell is the same'
        Else
            
            'Add to Total Volume'
            Total_Volume = Total_Volume + Cells(i, 7).Value
        
        End If
    
    Next i
    
'Change colors of yearly change
    For i = 2 To LastRow
    
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        Else
            Cells(i, 11).Interior.ColorIndex = 3
        
        End If
    
    Next i

End Sub
