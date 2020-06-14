'Execute in Module
Sub start()

Dim ws As Worksheet

Application.ScreenUpdating = False

For Each ws In Worksheets
         ws.Select
         Call SortBy
         Call Report
         Call Challenge
Next

Application.ScreenUpdating = True
    
End Sub

Sub SortBy()            'Sort criteria: Ticker and Date

With ActiveSheet.Sort

     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range("A1:G" & Cells(Rows.Count, "G").End(xlUp).Row)
     .Header = xlYes
     .Apply
     
End With

End Sub

Sub Report()

'Variable declaration
Dim ticker As String
Dim Total_Stock_Volume, Open_Value, Close_Value, ratio As Double
Dim last_row, i, j As Long
  
'Initialize counters & variables
Total_Stock_Volume = 0
Open_Value = 0
Close_Value = 0
ratio = 0
j = 2   'Auxiliary counter for Report/Summary section
last_row = Cells(Rows.Count, 1).End(xlUp).Row


'Write column names in for Report/Summary section
Range("M1").Value = "Ticker"
Range("N1").Value = "Total Stock Volume"
Range("O1").Value = "Yearly Change"
Range("P1").Value = "Percent Change"
  
  
For i = 2 To last_row       'Record unique tickers w/total volumes
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       ticker = Cells(i, 1).Value
       Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       Close_Value = Cells(i, 6).Value
       Range("M" & j).Value = ticker
       Range("N" & j).Value = Total_Stock_Volume
       Range("O" & j).Value = Open_Value - Close_Value
       
       If Open_Value <> 0 Then         ' to avoid overflow caused by 0/0
          Range("P" & j).Value = ((Close_Value - Open_Value) / (Abs(Open_Value)))
       End If
       
       
       j = j + 1
       Total_Stock_Volume = 0
       Close_Value = 0
       Open_Value = 0
       
    Else
       Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       
            If Open_Value = 0 Then
               Open_Value = Cells(i, 6).Value
            Else
               Open_Value = Open_Value
            End If
            
    End If
    
Next i
    
'Next, set formating conditions for Report/Summary Column

For Each O In Range("O2:O" & Cells(Rows.Count, "O").End(xlUp).Row)      'cell fill color based on value

    If O.Value > 0 Then
       O.Interior.ColorIndex = 4
    ElseIf O.Value < 0 Then
       O.Interior.ColorIndex = 3
    Else
       O.Interior.ColorIndex = xlNone
    End If
    
Next O
        
    Range("P2:P" & j).NumberFormat = "##.##%"    'Format %change column as percentage

End Sub
Sub Challenge()

'Write column names in for Challenge #1

    Range("S2").Value = "Greatest % increase"
    Range("S3").Value = "Lowest % decrease"
    Range("S4").Value = "Greatest Total Volume"
    Range("T1").Value = "Ticker"
    Range("U1").Value = "Value"
    
'Evaluate cells based on criteria. Worksheet functions use since values are already calculated
        
    Cells(2, 21).Value = WorksheetFunction.Max(Range("P2:P" & Cells(Rows.Count, 16).End(xlUp).Row))
    Cells(2, 20).Value = WorksheetFunction.index(Range("M2:P" & Cells(Rows.Count, 13).End(xlUp).Row), WorksheetFunction.Match(Cells(2, 21).Value, Range("P2:P" & Cells(Rows.Count, 14).End(xlUp).Row), 0), 1)
    Range("U2").NumberFormat = "###.##%"        'Convert values in Challenge 1 to %

    Cells(3, 21).Value = WorksheetFunction.Min(Range("P2:P" & Cells(Rows.Count, 16).End(xlUp).Row))
    Cells(3, 20).Value = WorksheetFunction.index(Range("M2:P" & Cells(Rows.Count, 13).End(xlUp).Row), WorksheetFunction.Match(Cells(3, 21).Value, Range("P2:P" & Cells(Rows.Count, 14).End(xlUp).Row), 0), 1)
    Range("U3").NumberFormat = "###.##%"        'Convert values in Challenge 1 to %

    Cells(4, 21).Value = WorksheetFunction.Max(Range("N2:N" & Cells(Rows.Count, 14).End(xlUp).Row))
    Cells(4, 20).Value = WorksheetFunction.index(Range("M2:P" & Cells(Rows.Count, 16).End(xlUp).Row), WorksheetFunction.Match(Cells(4, 21).Value, Range("N2:N" & Cells(Rows.Count, 14).End(xlUp).Row), 0), 1)
    
    
End Sub
