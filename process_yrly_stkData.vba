Option Explicit


Sub loop_ticker()


    'Dim WS_Count As Integer
    Dim I As Long
    Dim Last_Row, lastSRow, FirstSRow As Long
    Dim lastrow As Long
    Dim Last_Col As Long
    Dim column As Long
    Dim rowCounter As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim perChange As Double
    Dim ticker_volume As Double
    Dim yrlChng As Double
    Dim eachCellValue As Long
    Dim myrange As Range
    Dim str As String
    Dim rangeAddress As Variant
    Dim grt_volume As Double
    Dim grt_per_dcr As Double
    Dim grt_per_inr As Double
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tkrRange As Range
    Dim Cell As Object
    Dim raArray As Variant
    Dim starting_ws As Worksheet
    Dim Rng2Compare As Range
    
Debug.Print Now

    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
         'If ws.Name = "2015" Then
        ticker_volume = 0
        rowCounter = 2
        column = 1
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Volume"
            
        'Find the last non-blank cell in column A
        
        Last_Row = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
        Debug.Print "RowCount for " & ws.Name & "  is  " & Last_Row
            
        'Find the last non-blank cell in row G
        Last_Col = ws.Cells(1, Columns.Count).End(xlToLeft).column
    
        For I = 2 To Last_Row

            ticker_volume = ticker_volume + ws.Cells(I, 7).Value
            'Debug.Print i, ticker_volume
        
            If ws.Cells(I + 1, column).Value <> ws.Cells(I, column).Value Then
                'Message Box the value of the current cell and value of the next cell
                str = ws.Cells(I, column).Value
                'Debug.Print str
                ws.Cells(rowCounter, 10) = str ' Assign the ticker from Row 1 to Ticker row
                ws.Cells(rowCounter, 13) = ticker_volume
                
                'Call SelectByValue(range("A2:G70926"), "A")
                rowCounter = rowCounter + 1
                ticker_volume = 0
                'Debug.Print openingPrice, closingPrice, closingPrice - openingPrice, perChange
                'Debug.Print ws.Name, str, I
                End If ' for cells
            Next I
            
            lastrow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
            Set Rng2Compare = Application.Range("A2:A" & Last_Row)

            For eachCellValue = 2 To lastrow
                str = Cells(eachCellValue, 10).Value
                
                '********************'
                Range("A:A").AutoFilter Field:=1, Criteria1:=str
                lastSRow = ActiveCell.SpecialCells(xlCellTypeLastCell).Row
                FirstSRow = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
                '********************'
        
                openingPrice = ws.Cells(FirstSRow, 3).Value
                closingPrice = ws.Cells(lastSRow, 6).Value
                
                If openingPrice = 0 Then
                    perChange = 0
                Else
                    perChange = (closingPrice - openingPrice) / openingPrice
                End If
                
                yrlChng = closingPrice - openingPrice
                ws.Cells(eachCellValue, 11).NumberFormat = "0.000000000"
                ws.Cells(eachCellValue, 11) = yrlChng
        
                If yrlChng > 0 Then
                    ws.Cells(eachCellValue, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(eachCellValue, 11).Interior.Color = vbRed
                End If
        
                ws.Cells(eachCellValue, 12) = perChange
                ws.Cells(eachCellValue, 12).NumberFormat = "0.00%" '(2 decimals)
                
            Next eachCellValue
        
            grt_per_inr = Application.WorksheetFunction.Max(Columns("M"))
            grt_per_dcr = Application.WorksheetFunction.Min(Columns("M"))
            grt_volume = Application.WorksheetFunction.Max(Columns("K"))
            
            ws.Range("O2").Value = "Greatest % increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest total volume"
            ws.Range("Q1").Value = "Value"
            ws.Range("P1").Value = "Ticker"
            
            ws.Range("O2:O4").ColumnWidth = 25
            ws.Range("Q2:Q4").ColumnWidth = 10
            
            ws.Range("Q2").Value = grt_per_inr
            ws.Range("Q3").Value = grt_per_dcr
            ws.Range("Q4").Value = grt_volume
            
            ActiveSheet.AutoFilterMode = False
            
            Call fin_max_min(ws)
            Debug.Print "Processed " & ws.Name
            'End If 'end ws.name loop
     
            
          ThisWorkbook.Save
        Next ws
        starting_ws.Activate
        Debug.Print Now
End Sub

Sub fin_max_min(ws As Worksheet):
    Dim grt_volume As Double
    Dim grt_per_dcr As Double
    Dim grt_per_inr As Variant
    Dim Test As Range
    'Dim ws As Worksheet
    Dim c As Range
    Dim firstAddress As String
    Dim rownumber As Double
    Dim rng As Range
    Dim rngCell As Range
    
    'Set ws = ThisWorkbook.Worksheets(1)
    
    ws.Activate
                    
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("Q1").Value = "Value"
    ws.Range("P1").Value = "Ticker"
    
    ws.Range("O2:O4").ColumnWidth = 25
    ws.Range("Q2:Q4").ColumnWidth = 10
    
    grt_per_inr = Application.WorksheetFunction.Max(Columns("L"))
    grt_per_dcr = Application.WorksheetFunction.Min(Columns("L"))
    grt_volume = Application.WorksheetFunction.Max(Columns("M"))
    
    ws.Range("Q2").Value = grt_per_inr
    ws.Range("Q3").Value = grt_per_dcr
    ws.Range("Q4").Value = grt_volume
    
    Set rng = ws.Range("L:L") 'percentage
            
    'greatest percentage increase
    For Each rngCell In rng
        If rngCell.Value = grt_per_inr Then
        Cells(2, 16).Value = Cells(rngCell.Row, 10).Value
            Debug.Print "Increase", Cells(rngCell.Row, 10).Value
        End If
    Next rngCell
            
    'greatest percentage decrease
    For Each rngCell In rng
        If rngCell.Value = grt_per_dcr Then
            Cells(3, 16).Value = Cells(rngCell.Row, 10).Value
            Debug.Print "Decrease", Cells(rngCell.Row, 10).Value
        End If
    Next rngCell
        
    Set rng = ws.Range("M:M") 'volume
                      
    'greatest volume
    For Each rngCell In rng
        If rngCell.Value = grt_volume Then
        Cells(4, 16).Value = Cells(rngCell.Row, 10).Value
            Debug.Print "Volume", Cells(rngCell.Row, 10).Value
        End If
    Next rngCell
End Sub
