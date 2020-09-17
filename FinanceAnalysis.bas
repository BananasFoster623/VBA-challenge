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
Dim ticker(), startValue(), endValue(), totalVolumn() As String

Sub Main()

' Check to see if the sheet exists already; a safety net in case someone has already run
' the macro in a workbook
If WkshtCheck(title:="Results") = "False" Then
    Sheets.Add(Before:=Sheets(1)).Name = "Results"
End If

Call Work

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

' Loop down the first column to fill our arrays

' Step 1 - Find the current row's ticker, startValue, totalVolume and write to array

' Step 2 - Check next row down to see if the ticker is the same as the current row

' Step 3 - If next row is same, add total volume and loop
'        - If next row is different, add total volume, write endValue to array
'        - Iterate tickerIterator and loop

' Things we need:
'   - count of total rows
'   - count of unique rows
'   -

' Initialize the dummy variables
testTicker = ""
tickerIterator = 0

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
' Print everything out the respective cells
For i = 1 To countUnique
    Cells(i + 1, 9) = ticker(i - 1)
    Cells(i + 1, 10) = startValue(i - 1)
    Cells(i + 1, 11) = endValue(i - 1)
    Cells(i + 1, 12) = startValue(i - 1) - endValue(i - 1)
    Cells(i + 1, 13) = totalVolume(i - 1)
Next

End Sub

Function WkshtCheck(title As String)
' This function is used to check if the Results sheet has already been added to make sure
' we don't get an error when adding a new sheet. Probably not going to use this.
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = title Then
            WkshtCheck = "True"
            Exit For
        Else
            WkshtCheck = "False"
        End If
    Next
End Function

Sub Old()

For i = 2 To count
    If Cells(i, 1) <> testTicker Then
        ticker(tickerIterator) = Cells(i, 1)
        startValue(tickerIterator) = Cells(i, 3)
        testTicker = Cells(i, 1)
        volumeTracker = volumeTracker + Cells(i, 7)
        ' Set to the next ticker iterator to indicate
        tickerIterator = tickerIterator + 1
    End If
    
    If Cells(i + 1, 1) <> testTicker Then
        endValue(tickerIterator - 1) = Cells(i, 6)
        totalVolume(tickerIterator - 1) = volumeTracker + Cells(i, 7)
        volumeTracker = 0
    End If
Next


End Sub
