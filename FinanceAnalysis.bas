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
Dim i, j, k, count, countUnique, tickerIterator, volumeTracker As Long
Dim temp, testTicker As String
Dim ticker() As String
Dim startValue(), endValue(), totalVolumn() As Double
Dim ws As Worksheet

Sub Main()

For Each ws In Worksheets
    ws.Activate
    Call SetHeaders
    Call Work
Next

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

' Print everything out to the appropriate cells
Call PrintOut(ticker, startValue, endValue, totalVolume)

End Sub

Sub PrintOut(ticker, startValue, endValue, totalVolume)

For i = 1 To countUnique
    Cells(i + 1, 9) = ticker(i - 1)
    Cells(i + 1, 10) = endValue(i - 1) - startValue(i - 1)
    Cells(i + 1, 11) = 100 * ((endValue(i - 1) - startValue(i - 1)) / startValue(i - 1))
    Cells(i + 1, 12) = totalVolume(i - 1)
Next

End Sub

Sub SetHeaders()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

End Sub
