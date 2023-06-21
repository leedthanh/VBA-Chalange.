Attribute VB_Name = "Module1"
Option Explicit

Sub stonkproject()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate


Dim ticker As String
ticker = " "
Dim total_ticker As Double
total_ticker = 0

Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
percent_change = 0
Dim total_volume As Double
total_volume = 0

   
 Dim RowCount As Integer
  RowCount = 3400
  
  Dim increase_number As Integer
  
  Dim decrease_number As Integer
  
  Dim totalvol As Integer
  
  Dim increase_value As Integer
  
  Dim decrease_value As Integer
  
  Dim totalvalue As Integer
  
  
  


Dim summary_table_row As Long
summary_table_row = 2

Dim lastrow As Long
Dim i As Long

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent"
        Cells(1, 12).Value = "total volume"
        
        
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"


open_price = ws.Cells(2, 3).Value

    For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    Range("I" & summary_table_row).Value = ticker
    
    'find total volume
   total_volume = total_volume + ws.Cells(i, 7).Value
   Range("L" & summary_table_row).Value = total_volume
    'calculate price change
    close_price = ws.Cells(i, 6).Value
    yearly_change = close_price - open_price
    Range("J" & summary_table_row).Value = yearly_change
    If Range("J" & summary_table_row).Value > 0 Then
    Range("J" & summary_table_row).Interior.ColorIndex = 4
    Else
    Range("J" & summary_table_row).Interior.ColorIndex = 3
    End If
    'check for division of zero
    ' find percent change
    If open_price <> 0 Then
        percent_change = (yearly_change / open_price) * 100
        Range("K" & summary_table_row).Value = percent_change
    Else
        percent_change = 0
        
    End If
   
   'add 1 to summary table
   summary_table_row = summary_table_row + 1
   
    total_volume = 0
   
   open_price = ws.Cells(i + 1, 3)
   
  Else
   
  total_volume = total_volume + ws.Cells(i, 7).Value
     
    End If
                

            Next i
            
    ws.Cells(i, 11).NumberFormat = "0.00%"
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
  
  'math function to find greatest increase decrease and math with ticker
 
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
  Range("O2") = Cells(increase_number + 1, 9)
  decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
    Range("O3") = Cells(decrease_number + 1, 9)
   
   totalvol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
   
   Range("O4") = Cells(totalvol + 1, 9)

   
   
   
   'max function to find value of greatest increase decrease and total
   
increase_value = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
   Range("P2") = Cells(increase_value + 1, 11).Value

decrease_value = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
   
   Range("P3") = Cells(decrease_value + 1, 11).Value
   
   totalvalue = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
Range("P4") = Cells(totalvalue + 1, 12).Value
        
     
          
        Next ws
End Sub




