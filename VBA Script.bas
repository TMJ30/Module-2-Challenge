Attribute VB_Name = "Module1"
Option Explicit


Sub Ticker()

'Set up code to run through each worksheet

Dim ws As Worksheet
For Each ws In Worksheets
'Set column names
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Quarterly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    
    
    'Set dims
    Dim TickerName As String
    Dim OP As Double
    Dim CP As Double
    Dim PC As Double
    Dim QC As Double
    Dim TotalV As Double
    Dim i As Long
    Dim LR As Long
    Dim OpTbl As Integer
    
    
    
    'Assign values
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpTbl = 2
    TotalV = 0
    OP = ws.Cells(2, 3).Value 
    
    
    
    
    '----Conditions for totals and prints
    For i = 2 To LR
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Assign ticker name
            TickerName = ws.Cells(i, 1).Value
            'Print ticker name in excel (where to placee values)
            ws.Range("J" & OpTbl).Value = TickerName
            
            
            'Create and assign the sum volume per stock
            TotalV = TotalV + ws.Cells(i, 7).Value
            'Print sum volume per stock
            ws.Range("M" & OpTbl).Value = TotalV
           
            'Create closing price and assign quarterly change
            CP = ws.Cells(i, 6).Value
            QC = CP - OP
            'Print quartely change
             ws.Range("K" & OpTbl).Value = QC
            
            
            'Check for any 0/0 then calculate % change and assign per stock
            If OP = 0 Then 'Also note where I found the formula to handle non divisables by 0 because I forgot how it was handled in class
                PC = 0
            Else
                PC = QC / OP
            End If
            'Print % change
            ws.Range("L" & OpTbl).Value = PC
            
            'Format
            ws.Range("L" & OpTbl).NumberFormat = "0.00%"
           
            'Add one to summary table row...creates another row in summary table?/
             OpTbl = OpTbl + 1
            
            'Reset volume total
            TotalV = 0
            
            'Reset open price
            OP = ws.Cells(i + 1, 3)
            'If the next cell following a row is the same stock
        Else
            'Then add to the stock volume total
            TotalV = TotalV + ws.Cells(i, 7).Value
        End If
            
    Next i
    
        '---------------Find greatest % increase, decrease, and total volume
        'find the greatest percent change
        
        'Create row and column labels
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        'formula for greatest increase and decrease
        '1.) find the last row of name in new summary table
        Dim QCLR As Long
        Dim PCI As Double
        Dim PCD As Double
        Dim PCIName As String
        Dim PCDName As String
        Dim LargeTotal As Double
        Dim LargeTName As String
    
        'QCLR = Cells(Rows.Count, 10).End(xlUp).Row
        'finding max % of each name and return overall highest increase
        'print name matching highest % increase
        
        PCI = WorksheetFunction.Max(ws.Range("L:L")) 'finds greatest incease and prints %
        ws.Range("Q2").Value = PCI
        
        PCIName = Application.Match(ws.Range("Q2").Value, ws.Range("L:L"), 0)
        ws.Range("P2").Value = ws.Range("J" & PCIName)
        
        'print name and % decrease
        PCD = WorksheetFunction.Min(ws.Range("L:L")) 'finds greatest decrease and prints %
        ws.Range("Q3").Value = PCD
        
        PCDName = Application.Match(ws.Range("Q3").Value, ws.Range("L:L"), 0)
        ws.Range("P3").Value = ws.Range("J" & PCDName)
        
        'print name and highest total
        LargeTotal = WorksheetFunction.Max(ws.Range("M:M")) 'finds highest total $
        ws.Range("Q4").Value = LargeTotal
        
        LargeTName = Application.Match(ws.Range("Q4").Value, ws.Range("M:M"), 0)
        ws.Range("P4").Value = ws.Range("J" & LargeTName)
        
        
       '----------------Apply conditional formatting
       Dim CPColor As Double
       'Dim QCLR As Long
       Dim a As Long
       'grab last row of quarterly change
       QCLR = ws.Cells(Rows.Count, 11).End(xlUp).Row
       'format colos
       For a = 2 To QCLR
        
            If ws.Cells(a, 11).Value > 0 Then
                ws.Cells(a, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(a, 11).Value < 0 Then
                ws.Cells(a, 11).Interior.ColorIndex = 3
            ElseIf ws.Cells(a, 11).Value = 0 Then
                ws.Cells(a, 11).Interior.ColorIndex = 0
            End If
        Next a
    
Next ws

End Sub

