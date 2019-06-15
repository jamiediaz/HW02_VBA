Attribute VB_Name = "Module1"
Sub LeelooDallasMultipass()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call totvol
    Next
    Application.ScreenUpdating = True
End Sub
Sub totvol()

    Dim Y As Long           'Row number
    Dim LAST As Long        'Last row in the worksheet
    Dim wkSHEET As String   'Worksheet name variable
    Dim LPER As Double      'Last row for percentage column
    Dim SRNG As Long        'First row of a ticker symbol
    Dim LRNG As Long        'Last row of a ticker symbol
    Dim LVOL As Long        'Last row on total volume
    Dim TOT As Long         'Total Stock Volume
    Dim MAXPERT As Double   'High Percentage
    Dim MINPERT As Double   'Low Percentage
    Dim MAXVOL As Long      'max volume
    
    'This finds the last row of the column
    LAST = Cells(Rows.Count, "A").End(xlUp).Row
    
    'This retrieves the name of the current worksheet
    wkSHEET = ActiveSheet.Name
    
    'Sorts the worksheet by ticker symbol and automatically by date.
    'This will get rid of blank cells in the middle of the array.
    ActiveWorkbook.Worksheets(wkSHEET).sort.SortFields.Clear
    ActiveWorkbook.Worksheets(wkSHEET).sort.SortFields.Add2 Key:=Range("A2:A" & LAST), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(wkSHEET).sort.SortFields.Add2 Key:=Range("B2:B" & LAST), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(wkSHEET).sort
        .SetRange Range("A1:G" & LAST)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
        
    'This Builds the headings for the data
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Value"
    
    
    
    'This loop will group and total the rows until it gets to the LAST variable
    TOT = 2
    SRNG = 2
    
    For Y = 2 To LAST
        
        'This will add the Volume
        If Cells(Y, "A").Value = Cells(Y + 1, "A").Value Then
            Cells(TOT, "I").Value = Cells(Y, "A").Value
            Cells(TOT, "L").Value = Cells(TOT, "L").Value + Cells(Y, "G").Value
           
        'When the ticker symbol changes
        ElseIf Cells(Y, "A").Value <> Cells(Y + 1, "A").Value Then
            LRNG = Y
                    
            'This will calculate the change between open and close
            Cells(TOT, "J").Value = Cells(LRNG, "F").Value - Cells(SRNG, "C").Value
                
                'This will change the color between green and red
                If Cells(TOT, "J") >= 0 Then
                Cells(TOT, "J").Interior.ColorIndex = 10
        
                ElseIf Cells(TOT, "J") < 0 Then
                Cells(TOT, "J").Interior.ColorIndex = 9
                End If
                
            'Percent change calculation
                
                If Cells(SRNG, "C").Value = 0 Then
                Cells(TOT, "K").Value = FormatPercent(0)
                Else
                Cells(TOT, "K").Value = FormatPercent(Cells(TOT, "J").Value / Cells(SRNG, "C").Value)
                End If
                
            SRNG = Y + 1
            TOT = TOT + 1
        
        End If
        
    Next Y
     
    'calculate last row of percentage row and total volume row
    LPER = Cells(Rows.Count, "K").End(xlUp).Row
    LVOL = Cells(Rows.Count, "L").End(xlUp).Row
    
    'calculate max and min percentage change and greatest volume
    'It then matches the values found with the cells from the ticker associated with the values
    Range("Q2").Value = FormatPercent(Application.WorksheetFunction.Max(Range(Cells(2, "K"), Cells(LPER, "K"))))
        MAXPERT = WorksheetFunction.Match(Range("Q2").Value, Range((Cells(2, "k")), (Cells(LPER, "K"))), 0)
        Range("P2").Value = Cells(MAXPERT + 1, "I")
    Range("Q3").Value = FormatPercent(Application.WorksheetFunction.Min(Range(Cells(2, "K"), Cells(LPER, "K"))))
        MINPERT = WorksheetFunction.Match(Range("Q3").Value, Range((Cells(2, "k")), (Cells(LPER, "K"))), 0)
        Range("P3").Value = Cells(MINPERT + 1, "I")
    Range("Q4").Value = Application.WorksheetFunction.Max(Range(Cells(2, "L"), Cells(LVOL, "L")))
        MAXVOL = WorksheetFunction.Match(Range("Q4").Value, Range((Cells(2, "L")), (Cells(LPER, "L"))), 0)
        Range("P4").Value = Cells(MAXVOL + 1, "I")
    
    Columns("A:Q").AutoFit
End Sub
