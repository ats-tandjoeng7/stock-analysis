Attribute VB_Name = "VBA_Challenge"
'Completed by: Parto Tandjoeng
'GitHub repository: https://github.com/ats-tandjoeng7/stock-analysis
Option Explicit

Sub AllStocksAnalysis()
  Application.ScreenUpdating = False
  Dim yearvalue As String
  Dim startSS As Single, endSS As Single 'seconds
  yearvalue = InputBox("What year would you like to run the analysis on?", "Input for AllStocksAnalysis", "2018")
  If yearvalue = vbNullString Then 'when users press Cancel button
    Exit Sub
  End If
  startSS = Timer
  Worksheets("All Stocks Analysis").Activate
'  ClearWorksheet "All Stocks Analysis" 'clear cells
'Create a header row w/ Cells method
  Cells(1, 1).Value = "All Stocks (" + yearvalue + ")"
  Cells(3, 1).Value = "Ticker"
  Cells(3, 2).Value = "Total Daily Volume": Columns("B").AutoFit
  Cells(3, 3).Value = "Return"
  
  Dim id As Integer, i As Integer, rowEnd As Integer, colEnd As Integer
  Dim startingPrice As Double, endingPrice As Double
  Dim ticker As String
  'Initialize array of all tickers
  Dim tickers() As String
  tickers = Split("AY,CSIQ,DQ,ENPH,FSLR,HASI,JKS,RUN,SEDG,SPWR,TERP,VSLR", ",")
'  Dim tickers As Variant
'  tickers = Array("AY", "CSIQ", "DQ", "ENPH", "FSLR", "HASI", "JKS", "RUN", "SEDG", "SPWR", "TERP", "VSLR")
  Dim totalVolume As Long: totalVolume = 0 'reset totalVolume
  'establish the number of rows to loop over
  Dim rowStart As Integer: rowStart = 2
  Worksheets(yearvalue).Activate
  rowEnd = Cells(Rows.Count, 1).End(xlUp).Row 'last non-empty row
'  colEnd = Cells(1, Columns.Count).End(xlToLeft).Column 'last non-empty col
  For id = LBound(tickers) To UBound(tickers)
    ticker = tickers(id) ': Debug.Print ticker
    totalVolume = 0
    For i = rowStart To rowEnd 'loop over all the rows
      If Cells(i, 1).Value = ticker Then
        If Cells(i - 1, 1).Value <> ticker Then
          startingPrice = Cells(i, 6).Value
        ElseIf Cells(i + 1, 1).Value <> ticker Then
          endingPrice = Cells(i, 6).Value
        End If
        totalVolume = totalVolume + Cells(i, 8).Value 'increase totalVolume by the value in the current row
      End If
    Next i
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + id, 1).Value = ticker
    Cells(4 + id, 2).Value = totalVolume
    Cells(4 + id, 3).Value = endingPrice / startingPrice - 1
    Worksheets(yearvalue).Activate
  Next id
  formatAllStocksAnalysisTable
  endSS = Timer
  MsgBox "This code ran in " & (endSS - startSS) & " seconds for the year " & yearvalue
  Application.ScreenUpdating = True
End Sub

Sub DQAnalysis()
  Application.ScreenUpdating = False
  Worksheets("DQ Analysis").Activate
'  Range("A1:C3").Clear
''Create a header row w/ Range method
'  Range("A1").Value = "DAQO (Ticker: DQ)"
'  Range("A" & 3).Value = "Year"
'  Range("B" & 3).Value = "Total Daily Volume" ': Columns("B").AutoFit
'  Range("C" & 3).Value = "Return"
'  Range(Cells(1, 1), Cells(3, 3)).Clear
'Create a header row w/ Cells method
  Cells(1, 1).Value = "DAQO (Ticker: DQ)"
  Cells(3, 1).Value = "Year"
  Cells(3, 2).Value = "Total Daily Volume": Columns("B").AutoFit
  Cells(3, 3).Value = "Return"
  Cells(3, 4).Value = "Starting Price"
  Cells(3, 5).Value = "Ending Price"
  
  Worksheets("2018").Activate
  Dim i As Integer, rowEnd As Integer, colEnd As Integer
  Dim startingPrice As Double, endingPrice As Double
  Dim totalVolume As Long: totalVolume = 0 'reset totalVolume
  'establish the number of rows to loop over
  Dim rowStart As Integer: rowStart = 2
  rowEnd = Cells(Rows.Count, 1).End(xlUp).Row 'last non-empty row
  colEnd = Cells(1, Columns.Count).End(xlToLeft).Column 'last non-empty col
'  rowEnd = Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row 'last non-empty row
'  colEnd = Cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column 'last non-empty col
'  Debug.Print Rows.Count, Columns.Count, rowEnd, colEnd '1048576, 16384, 3013, 8
  For i = rowStart To rowEnd 'loop over all the rows
    If Cells(i, 1).Value = "DQ" Then
      If Cells(i - 1, 1).Value <> "DQ" Then
        startingPrice = Cells(i, colEnd - 2).Value
      ElseIf Cells(i + 1, 1).Value <> "DQ" Then
        endingPrice = Cells(i, colEnd - 2).Value
      End If
      totalVolume = totalVolume + Cells(i, colEnd).Value 'increase totalVolume by the value in the current row
    End If
  Next i
'  If (totalVolume = 107873900) Then MsgBox ("Numbers matched")
  Worksheets("DQ Analysis").Activate
  Cells(4, 1).Value = 2018
  Cells(4, 2).Value = totalVolume
  Cells(4, 3).Value = endingPrice / startingPrice - 1: Cells(4, 3).NumberFormat = "0.00%"
  Cells(4, 4).Value = startingPrice
  Cells(4, 5).Value = endingPrice
  Range(Cells(4, 4), Cells(4, 5)).NumberFormat = "$ #.##"
  Application.ScreenUpdating = True
End Sub

Sub chkbrd()
  Dim newWS As Worksheet
  Const myWS As String = "chkbrd"
  On Error GoTo ErrHandler
  Set newWS = Sheets.Add(After:=Worksheets(Worksheets.Count)): newWS.Name = myWS 'add/rename sheet
  On Error GoTo 0
  presetWS 'clear cells
  makechkbrd 2, 16, vbRed, vbGreen '2x16 checkerboard pattern w/ specified colors
  MsgBox ("Click OK to continue.")
  presetWS 'clear cells
  makechkbrd 8, 8 '8x8 checkerboard pattern with default colors
  MsgBox ("Click OK to continue.")
  presetWS 'clear cells
  makechkbrd 16, 16, vbMagenta, vbCyan '16x16 checkerboard pattern w/ specified colors
Exit Sub
ErrHandler:
  Application.DisplayAlerts = False
  newWS.Delete 'delete worksheet w/o prompt
  Application.DisplayAlerts = True
  Resume Next
End Sub

Sub listI2()
  Dim newWS As Worksheet
  Const myWS As String = "listI2"
  Dim i As Integer, j As Integer, rowNo As Integer, colNo As Integer
  On Error GoTo ErrHandler
  Set newWS = Sheets.Add(After:=Worksheets(Worksheets.Count)): newWS.Name = myWS 'add/rename sheet
  On Error GoTo 0
  presetWS myWS 'clear cells
  For i = 1 To 10
    Cells(1, i).Value = i * i
  Next i
  MsgBox ("Cell G1 equals " & Range("G1").Value & vbLf & "Click OK to continue.") '7*7=49
  
  '1 in each cell
  For i = 1 To 10
    For j = 1 To 10
      Cells(i, j).Value = 1
    Next j
  Next i
  MsgBox ("Click OK to continue.")
  presetWS myWS 'clear cells
  'sum of row number and column number in each cell
  For i = 1 To 10
    For j = 1 To 10
      rowNo = Cells(i, j).Row
      colNo = Cells(i, j).Column
      Cells(i, j).Value = rowNo + colNo
    Next j
  Next i
Exit Sub
ErrHandler:
  Application.DisplayAlerts = False
  newWS.Delete 'delete worksheet w/o prompt
  Application.DisplayAlerts = True
  Resume Next
End Sub

'Functions
Private Function makechkbrd(myRow As Integer, myCol As Integer, Optional color1 As String = vbBlack, Optional color2 As String = vbWhite)
  If myRow > 64 Or myCol > 64 Then 'max size is 64x64
    Exit Function
  Else
    If myRow Mod 2 <> 0 Then
      myRow = myRow + 1
    End If
    If myCol Mod 2 <> 0 Then
      myCol = myCol + 1
    End If
  End If
  'recycle the smallest unit of checkerboard patterns
  Dim myRC As Range
  Set myRC = Application.Union(Range("A1"), Range("B2")): myRC.Interior.Color = color1
  Set myRC = Application.Union(Range("A2"), Range("B1")): myRC.Interior.Color = color2
  ActiveSheet.Range("A1:B2").Copy ActiveSheet.Range(Cells(1, 1), Cells(myRow, myCol))
End Function

Private Function formatAllStocksAnalysisTable(Optional myWS As String = "All Stocks Analysis")
  Worksheets(myWS).Activate
  'Formatting
  Cells(1, 1).Font.FontStyle = "Bold"
  Range("A3:C3").Font.FontStyle = "Bold"
  Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
  Range("B4:B15").NumberFormat = "#,##0"
  Range("C4:C15").NumberFormat = "0.00%"
  Dim i As Integer, rowStart As Integer, rowEnd As Integer
  rowStart = 4
  rowEnd = Cells(Rows.Count, 1).End(xlUp).Row 'last non-empty row
  For i = rowStart To rowEnd
    If Cells(i, 3) > 0 Then
      'Color the cell green
      Cells(i, 3).Interior.Color = vbGreen
    ElseIf Cells(i, 3) < 0 Then
      If Cells(i, 3) <= -0.15 Then
        'Color the cell red
        Cells(i, 3).Interior.Color = vbRed
      Else
        'Color the cell yellow
        Cells(i, 3).Interior.Color = vbYellow
      End If
    Else
      'Clear the cell color
      Cells(i, 3).Interior.Color = xlNone
    End If
  Next i
End Function

Function ClearWorksheet(myWS As String)
  Worksheets(myWS).Activate
  Cells.Clear 'clear cells
End Function

Private Function presetWS(Optional myWS As String = "chkbrd")
  Const myRow = 64, myCol = 64
  Worksheets(myWS).Activate
  Cells.Clear 'clear cells
  Range(Cells(1, 1), Cells(myRow, myCol)).ColumnWidth = Range("A1").RowHeight / 6 'resize column width
End Function

