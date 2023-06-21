Attribute VB_Name = "Multi_Year_Stock"
Sub Stocks()



Dim i As Double
Dim Tic_name As String
Dim YC As Double
    YC = 0
Dim PC As Double
    PC = 0
Dim TSV As Double
    TSV = 0
Dim Summary_Table_Row As Integer
Dim ws As Worksheet
Dim first As Double
Dim change As Double
Dim GPI As Double
Dim GPD As Double
Dim GTV As Double



For Each ws In Sheets

    Summary_Table_Row = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Range("I1:L1").Columns.AutoFit

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Range("O4:Q1").Columns.AutoFit

'specialcase first row
first = ws.Cells(2, 3).Value


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
    
    'Loop
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Tic_name = ws.Cells(i, 1).Value
            YC = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
            
            TSV = TSV + ws.Cells(i, 7).Value
            change = ws.Cells(i, 6).Value - first
            PC = change / first
            
            
            
            'output results
            ws.Range("I" & Summary_Table_Row).Value = Tic_name
            ws.Range("J" & Summary_Table_Row).Value = change
            ws.Range("K" & Summary_Table_Row).Value = PC
            ws.Range("L" & Summary_Table_Row).Value = TSV
                
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset variables
            first = ws.Cells(i + 1, 3).Value
            TSV = 0
        Else
            TSV = TSV + ws.Cells(i, 7).Value
            
        End If
        
        'format cells
        If ws.Cells(i, 10) >= 0 Then
            ws.Cells(i, 10).Interior.Color = vbGreen
        Else
            ws.Cells(i, 10).Interior.Color = vbRed
        End If
        
    Next i
    
    'worksheetfunction.min
        ws.Cells(2, 17) = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        ws.Cells(3, 17) = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        ws.Cells(4, 17) = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
  
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
           
           
    
    'worksheetfuntion.match
        GPI = Application.WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2:K" & lastRow), 0)
        GPD = Application.WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2:K" & lastRow), 0)
        GTV = Application.WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2:L" & lastRow), 0)
    
    
    'print the ticker
        ws.Cells(2, 16).Value = ws.Cells(GPI + 1, 9).Value
        ws.Cells(3, 16).Value = ws.Cells(GPD + 1, 9).Value
        ws.Cells(4, 16).Value = ws.Cells(GTV + 1, 9).Value
    
    Next ws
    


End Sub

