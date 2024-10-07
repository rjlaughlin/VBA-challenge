Sub Stocks()

'Dimension All Variables / Turning off Screen Updating
'------------------------------------------------------
    Dim rngTicker, rngPercentChange, rngQuarterlyChange, rngTotalVolume As Range
    Dim strTicker, strMin, strMax, strTickMin, strTickMax, strTickVolume As String
    Dim numOpen, numClose As Double
    Dim numMin, numMax As Variant
    Dim i, x, y, numVolume, clnOpen, clnClose, clnVol As Integer
    Dim blnStatus As Boolean

    Application.ScreenUpdating = False

' Initiating For Loop to loop through each sheet
'------------------------------------------------------

    For i = 1 To ThisWorkbook.Sheets.Count

        With ThisWorkbook.Sheets(i)

' Clear Formatting
'------------------------------------------------------
        .Cells.FormatConditions.Delete

        For x = 9 To 18
            .Columns(x).ClearContents
        Next x

' Repopulate designated cells with Text
'------------------------------------------------------
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Quarterly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(4, 15).Value = "Greatest Total Volume"
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"

' Define ranges and other variables
'------------------------------------------------------
        Set rngTicker = .Range(.Cells(2, 1), .Cells(.UsedRange.Rows.Count, 1))

        clnVol = .Cells.Find(What:="<vol>", After:=.Cells(2, 1), LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column

        clnOpen = .Cells.Find(What:="<open>", After:=.Cells(2, 1), LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column

        clnClose = .Cells.Find(What:="<close>", After:=.Cells(2, 1), LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column

        numVolume = 0
        blnStatus = False
        y = 2

' Initiate For Loop to go through each ticker and establish Open price
'------------------------------------------------------
        For Each Cell In rngTicker.Cells
            If blnStatus = False Then
                strTicker = Cell.Value
                numOpen = .Cells(Cell.Row, clnOpen).Value
                numVolume = .Cells(Cell.Row, clnVol).Value
            End If

' If Statement to determine stats and copy into designated cells
'------------------------------------------------------
            If Cell.Offset(1, 0).Value = strTicker Then
                numVolume = numVolume + .Cells(Cell.Offset(1, 0).Row, clnVol).Value
                numClose = .Cells(Cell.Offset(1, 0).Row, clnClose).Value
                blnStatus = True

            Else
                .Cells(y, 9).Value = strTicker
                .Cells(y, 10).Value = Format(numClose - numOpen, "0.00")
                .Cells(y, 11).Value = Format(numClose / numOpen - 1, "0.00%")
                .Cells(y, 12).Value = Format(numVolume, "0")
                y = y + 1
                blnStatus = False
            End If
        Next Cell

' Loop to go through Percent Change
'------------------------------------------------------
        Set rngQuarterlyChange = .Range(.Cells(2, 10), .Cells(.Cells(2, 10).End(xlDown).Row, 10))
        Set rngPercentChange = .Range(.Cells(2, 11), .Cells(.Cells(2, 11).End(xlDown).Row, 11))
        Set rngTotalVolume = .Range(.Cells(2, 12), .Cells(.Cells(2, 12).End(xlDown).Row, 12))

' Alternate method of collecting max/min values
'------------------------------------------------------
''numMax = Application.WorksheetFunction.Max(rngPercentChange)
''
''numMin = Application.WorksheetFunction.Min(rngPercentChange)
''
''numVolume = Application.WorksheetFunction.Max(rngTotalVolume)
''
''strTickMax = .Cells.Find(What:=numMax * 100, After:=.Cells(1, 11), LookIn:=xlFormulas2, LookAt _
''        :=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:= _
''        False, SearchFormat:=False).Offset(0, -2).Formula
''
''strTickMin = .Cells.Find(What:=numMin * 100, After:=.Cells(1, 11), LookIn:=xlFormulas2, LookAt _
''        :=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:= _
''        False, SearchFormat:=False).Offset(0, -2).Formula
''
''strTickVolume = .Cells.Find(What:=numVolume, After:=.Cells(1, 12), LookIn:=xlFormulas2, LookAt _
''        :=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:= _
''        False, SearchFormat:=False).Offset(0, -2).Formula
'
''Debug.Print numMax & "  " & strTickMax
''Debug.Print numMin & "  " & strTickMin & vbNewLine

' For loop through Percent Change values
'------------------------------------------------------
        For Each Cell In rngPercentChange
            If Cell.Row = 2 Then
                numMin = Cell.Value
                numMax = Cell.Value
                numVolume = .Cells(Cell.Row, 12).Value
                strTickMin = .Cells(Cell.Row, 9).Value
                strTickMax = .Cells(Cell.Row, 9).Value
            Else

            If Cell.Value > numMax Then
                numMax = Cell.Value
                strTickMax = .Cells(Cell.Row, 9).Value
            End If

            If Cell.Value < numMin Then
                numMin = Cell.Value
                strTickMin = .Cells(Cell.Row, 9).Value
            End If

            If .Cells(Cell.Row, 12).Value > numVolume Then
                numVolume = .Cells(Cell.Row, 12).Value
                strTickVolume = .Cells(Cell.Row, 9).Value
            End If
            End If

        Next Cell

        .Cells(2, 16).Value = strTickMax
        .Cells(2, 17).Value = Format(numMax, "0.00%")
        .Cells(3, 16).Value = strTickMin
        .Cells(3, 17).Value = Format(numMin, "0.00%")
        .Cells(4, 16).Value = strTickVolume
        .Cells(4, 17).Value = Format(numVolume, "0")


' Conditional formatting and additional formatting
'------------------------------------------------------
        .Range(.Cells(1, 10), .Cells(1, 15)).Columns.AutoFit

        rngQuarterlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"

        With rngQuarterlyChange.FormatConditions(rngQuarterlyChange.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 255, 0)
        .TintAndShade = 0
        End With


        rngQuarterlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"

        With rngQuarterlyChange.FormatConditions(rngQuarterlyChange.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        End With


        rngPercentChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"

        With rngPercentChange.FormatConditions(rngPercentChange.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 255, 0)
        .TintAndShade = 0
        End With


        rngPercentChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"

        With rngPercentChange.FormatConditions(rngPercentChange.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        End With

        .Columns(15).ColumnWidth = 19
        .Columns(17).ColumnWidth = 8.5
        .Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        .Cells(1, 1).Activate

        End With

' Next worksheet command (i)
'------------------------------------------------------
Next i

ThisWorkbook.Sheets(1).Activate
Application.ScreenUpdating = True

End Sub

