Option Explicit 'Forces the code to declare everything explicitly
'Global Initializations - these will be accessible from any Function or Sub in this file
Private inputString As String, strPath As String, strLine As String
Private resultsCol As Integer, numberofOccurrences As Integer, searchTerms As Integer, intResult As 
Integer, FoundCount As Integer
Private wsTest As Worksheet
Private dataChart As Chart
Sub List_based_word_query()

Dim targetRng As Range
resultsCol = 2
searchTerms = 0

'Create new worksheet for ListResults
Set wsTest = Nothing
On Error Resume Next
Set wsTest = ActiveWorkbook.Sheets("ListResults")
On Error GoTo 0
If wsTest Is Nothing Then
    Worksheets.Add.Name = "ListResults"
End If

'Clear the ListResults Sheet
Worksheets("ListResults").Cells.Clear
On Error Resume Next
Worksheets("ListResults").ChartObjects.Delete
On Error GoTo 0

'Open a file of the user's choosing and activate it, reading it line by line
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intResult = Application.FileDialog(msoFileDialogOpen).Show
If intResult <> 0 Then
    strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Open strPath For Input As #1
    Line Input #1, strLine 'reads first line
    Do
        inputString = strLine  'inputString = current line
        Line Input #1, strLine 'next line
            numberofOccurrences = Single_word_occurrences(inputString, ColumnLetter(resultsCol))
        Sheets("ListResults").Cells(2, ColumnLetter(resultsCol)).Value = numberofOccurrences
            While strLine = "Or"
                Line Input #1, strLine 'next line
                numberofOccurrences = Separated_by_Or(strLine, ColumnLetter(resultsCol))
                Sheets("ListResults").Cells(2, ColumnLetter(resultsCol)).Value = FoundCount
                searchTerms = searchTerms + 1
                Line Input #1, strLine 'next line
            Wend
        resultsCol = resultsCol + 1
        searchTerms = searchTerms + 1
    Loop While EOF(1) = False 
    inputString = strLine 'the loop exits when at the last word, so we have to run one more time
    numberofOccurrences = Single_word_occurrences(inputString, ColumnLetter(resultsCol))
    Sheets("ListResults").Cells(2, ColumnLetter(resultsCol)).Value = numberofOccurrences
    resultsCol = resultsCol + 1
    searchTerms = searchTerms + 1
End If
Close #1

'Create a new chart.
Set dataChart = Charts.Add
Set dataChart = dataChart.Location(Where:=xlLocationAsObject, Name:="ListResults")
With dataChart
  .ChartType = xl3DBarStacked
  .SetSourceData Source:=Sheets("ListResults").Range(Cells(1, 2), Cells(2, resultsCol - 1)), PlotBy:=xlRows
  .HasTitle = True
  .ChartTitle.Text = "Occurrences of Keywords"
  .SetElement (msoElementDataLabelShow)
  With .Parent
    .Top = Range("C5").Top
    .Left = Range("C5").Left
    .Height = 75
    .Width = 100
    .Name = "Occurrences of Keywords"
  End With
End With
Sheets("ListResults").Cells(1, 1).Value = "Others"
Sheets("ListResults").Rows(1).Font.Bold = True
Sheets("ListResults").Rows.WrapText = True»     »       »       »       
Sheets("ListResults").Columns.ColumnWidth = 25
Sheets("ListResults").Cells.HorizontalAlignment = xlLeft
Sheets("ListResults").Cells.VerticalAlignment = xlTop
Exit Sub
End Sub

'This function takes the column number and converts it to a letter
Function ColumnLetter(ColumnNumber As Integer) As String
    Dim n As Integer
    Dim c As Byte
    Dim s As String
   
    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

'Run this function with all words after the first in a sequence with "or" separating the terms
Function Separated_by_Or(datatoFind As String, resultsCol As String) As Integer
    Dim strFirstAddress As String
    Dim foundRange As Range, LastAddress As Range
    Dim currentSheet As Integer, sheetCount As Integer, LastRow As Long, loopedOnce As Integer, numberofOthers As Integer

    loopedOnce = 0
    currentSheet = ActiveSheet.Index
    sheetCount = ActiveWorkbook.Sheets.Count
    Sheets("Sheet1").Activate
    Set foundRange = Range("F2:F30000").Find(What:=datatoFind, After:=Cells(2, 6), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    Sheets("ListResults").Cells(1, resultsCol).Value = Sheets("ListResults").Cells(1, resultsCol).Value & ", " & datatoFind
    If Not foundRange Is Nothing Then 'if datatoFind is found in search range
        strFirstAddress = foundRange.Address 'strFirstAddress = address of first occurrence of datatoFind
        Do 'Find next occurrence of datatoFind
            Set foundRange = Range("F2:F30000").Find(What:=datatoFind, After:=foundRange.Cells, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            'Place the value of this occurrence in the next cell down in the column that holds found values (resultsCol column of ListResults worksheet)
            LastRow = Sheets("ListResults").Range(resultsCol & Rows.Count).End(xlUp).Row + 1
            Sheets("ListResults").Range(resultsCol & LastRow).Value = foundRange.Address
            If loopedOnce = 1 Then
                FoundCount = FoundCount + 1
                If Sheets("ListResults").Cells(1000, "A").Value = "" Then 'This optimizes the Macro...without it the Others column will populate for a LONG time
                    numberofOthers = Others(LastAddress, foundRange) 'Currently not using this variable because I've limited the visible comments in Column A to 1000
                End If
            End If
            If loopedOnce = 0 Then
                loopedOnce = 1
            End If
            Set LastAddress = foundRange
            'The Loop ends on reaching the first occurrence of datatoFind
        Loop While foundRange.Address <> strFirstAddress And Not foundRange Is Nothing
    End If
    Separated_by_Or = FoundCount
    Application.ScreenUpdating = True
    Sheets(currentSheet).Activate
End Function

'This function populates the "Others" column (column A) with results that don't match any of the terms
Function Others(startCell As Range, endCell As Range) As Integer
    Dim i As Long, LastRow As Long
    Dim currentCell As Range
    i = 0
    For i = (startCell.Row + 1) To endCell.Row Step 1
        Set currentCell = startCell.Offset(i - startCell.Row, 0)
        If currentCell.Address = endCell.Address Then
            Others = i - startCell.Row
            Exit Function
        End If
        LastRow = Sheets("ListResults").Cells(Rows.Count, "A").End(xlUp).Row + 1
        Sheets("ListResults").Cells(LastRow, "A").Value = currentCell.Value
    Next
End Function

Function Single_word_occurrences(datatoFind As String, resultsCol As String) As Integer
    'Initializations
    Dim strFirstAddress As String
    Dim foundRange As Range, LastAddress As Range
    Dim currentSheet As Integer, sheetCount As Integer, LastRow As Long, loopedOnce As Integer, numberofOthers As Integer

    loopedOnce = 0
    FoundCount = 0
    currentSheet = ActiveSheet.Index
    sheetCount = ActiveWorkbook.Sheets.Count
    Sheets("Sheet1").Activate
    Set foundRange = Range("F2:F30000").Find(What:=datatoFind, After:=Cells(2, 6), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    Sheets("ListResults").Cells(1, resultsCol).Value = datatoFind
    If Not foundRange Is Nothing Then 'if datatoFind is found in search range
        strFirstAddress = foundRange.Address 'strFirstAddress = address of first occurrence of datatoFind
        Do 'Find next occurrence of datatoFind
            Set foundRange = Range("F2:F30000").Find(What:=datatoFind, After:=foundRange.Cells, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            'Place the value of this occurrence in the next cell down in the column that holds found values (resultsCol column of ListResults worksheet)
            LastRow = Sheets("ListResults").Range(resultsCol & Rows.Count).End(xlUp).Row + 1
            Sheets("ListResults").Range(resultsCol & LastRow).Value = foundRange.Address
            If loopedOnce = 1 Then
                FoundCount = FoundCount + 1
                If Sheets("ListResults").Cells(1000, "A").Value = "" Then 'This optimizes the Macro...without it the Others column will populate for a LONG time
                    numberofOthers = Others(LastAddress, foundRange) 'Currently not using this variable because I've limited the visible comments in Column A to 1000
                End If
            End If
            If loopedOnce = 0 Then
                loopedOnce = 1
            End If
            Set LastAddress = foundRange
            'The Loop ends on reaching the first occurrence of datatoFind
        Loop While foundRange.Address <> strFirstAddress And Not foundRange Is Nothing
    End If
    Single_word_occurrences = FoundCount
    Application.ScreenUpdating = True
    Sheets(currentSheet).Activate
End Function