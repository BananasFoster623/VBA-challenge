Attribute VB_Name = "FinanceAnalysis"
Option Explicit

' Put all variable dimensioning here, except ReDim()
' a = dummy Range variable for counting the total number of rows
' b = dummy Range variable for counting the total number of rows
' count = used to store the number of rows that we will use to iterate
' countUnique = used to count the number of unique rows (i.e. ticker symbols)
' tickerIterator = iterator for setting values into the correct location in the arrays
' volumnTracker = used to track the total volume of a specific ticker
' temp = used for ticker comparison
' testTicker = used for ticker comparison

Dim a, b As Range
Dim i, j, k, count, countUnique, tickerIterator, volumeTracker, greatVolumeLoc As Long
Dim temp, testTicker As String
Dim ticker() As String
Dim startValue(), endValue(), totalVolume(), greatVolume As Double
Dim ws As Worksheet

Sub Main()

    For Each ws In Worksheets
        ws.Activate
        Call SetHeaders
        Call Work
        Call Formatting
    Next
    MsgBox ("Action Successful. No errors.")
    
End Sub

Sub Work()

    ' This section finds the length of the current region
    ' which is all the A values on Sheet "A"
    Range("A2").CurrentRegion.Select
    Set a = Selection
    count = 0
    For Each b In a.Rows
        count = count + 1
    Next
    
    ' This sections finds the number of unique ticker values by looping
    ' over all rows in column 1
    temp = ""
    countUnique = 0
    For i = 2 To count
        If Cells(i, 1) <> temp Then
            temp = Cells(i, 1)
            countUnique = countUnique + 1
        End If
    Next i
    
    ' ReDim the dynamic arrays to the size of countUnique so that they are the exact size for the number of tickers
    ReDim ticker(countUnique)
    ReDim startValue(countUnique)
    ReDim endValue(countUnique)
    ReDim totalVolume(countUnique)
    
    ' Initialize the dummy variables
    testTicker = ""
    tickerIterator = 0
    
    ' Loop down the first column to fill our arrays
    For i = 2 To count
        If i = 2 Then
            ticker(tickerIterator) = Cells(i, 1).Value
            startValue(tickerIterator) = Cells(i, 3).Value
            totalVolume(tickerIterator) = Cells(i, 7).Value
        ElseIf Cells(i + 1, 1).Value = "" Then
            GoTo Done
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            totalVolume(tickerIterator) = totalVolume(tickerIterator) + Cells(i, 7).Value
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            endValue(tickerIterator) = Cells(i, 6).Value
            totalVolume(tickerIterator) = totalVolume(tickerIterator) + Cells(i, 7).Value
            tickerIterator = tickerIterator + 1
            ticker(tickerIterator) = Cells(i + 1, 1).Value
            startValue(tickerIterator) = Cells(i + 1, 3).Value
            totalVolume(tickerIterator) = 0
        Else
            MsgBox ("Uh oh...something went wrong....")
        End If
    Next
Done:

    greatVolume = 0
    
    For i = 1 To tickerIterator
        If totalVolume(i) > greatVolume Then
            greatVolume = totalVolume(i)
            greatVolumeLoc = i
        End If
    Next
    
'    For i = LBound(totalVolume) To UBound(totalVolume)
'        If totalVolume(i) > greatVolume Then
'            greatVolume = totalVolume(i)
'        End If
'    Next i
    
    ' Print everything out to the appropriate cells
    Call PrintOut(ticker, startValue, endValue, totalVolume)

End Sub

Sub PrintOut(ticker, startValue, endValue, totalVolume)

    For i = 1 To countUnique
        Cells(i + 1, 9) = ticker(i - 1)
        Cells(i + 1, 10) = endValue(i - 1) - startValue(i - 1)
        If startValue(i - 1) = 0 Then
            Cells(i + 1, 11) = "NaN"
        Else
            Cells(i + 1, 11) = ((endValue(i - 1) - startValue(i - 1)) / startValue(i - 1))
        End If
        Cells(i + 1, 12) = totalVolume(i - 1)
    Next
    
    Cells(4, 16).Value = ticker(greatVolumeLoc)
    Cells(4, 17).Value = greatVolume

End Sub

Sub SetHeaders()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

End Sub

Sub Formatting()

    ' Autofit the column widths for the columns we added
    Columns("I:Q").Select
    Selection.Columns.AutoFit
    
    ' Use conditional formatting to format positive values as green and negative values as red
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5756247
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5197823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
        
    ' Format percent change as a percentage with 2 decimal places
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    ' Format Total Stock Volume with commas and no decimal places
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
End Sub
