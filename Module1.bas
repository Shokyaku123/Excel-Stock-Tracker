Sub StockAnalysis():

    'get variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim vol As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer
    Dim lastRow As Long
    Dim i As Long
    
  For Each ws In Worksheets
    
' Set ws = ThisWorkbook.Worksheets("2018")
 ' On Error Resume Next
  ' On Error GoTo 0
    lawRow = Cells(Rows.Count, 1).End(xlUp).Row
 'making code work for whole dataset
 
For i = 2 To lastRow
    'Check if the cell value can be converted to a double
    If IsNumeric(ws.Cells(i, 7).Value) Then
        vol = CDbl(ws.Cells(i, 7).Value)
    Else
        MsgBox "invalid value in cell (" & i & ", 7)"
    End If
    Next i
    
        
  'setting headers here
  
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


    
 'loop integers
Summary_Table_Row = 2
 year_open = ws.Cells(2, 3).Value
 'the loop itself
        For i = 2 To ws.UsedRange.Rows.Count
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        'need value
            ticker = ws.Cells(i, 1).Value
            vol = ws.Cells(i, 7).Value
            ' year_open = ws.Cells(i, 3).Value ' can't update year_open when new row is found
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close
        
        'insert values into the summary set
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
            year_open = ws.Cells(i, 3).Value
            
            vol = 0
        
        End If
    
    Next i
    
        ws.Columns("K").NumberFormat = "0.00%"
        
            'format with colors
            Dim rg As Range
            Dim g As Long
            Dim c As Long
            Dim color_cell As Range
            
      Set rg = ws.Range("J2:J" & Summary_Table_Row)
        c = rg.Cells.Count
        
        For g = 1 To c
        Set color_cell = rg(g)
        Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
        
        End With
    End Select
        Next g
        

    Next ws
    
            
End Sub
