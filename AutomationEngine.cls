VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutomationEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private currentWorkBook As Workbook


Public Sub FormatColumnsAsText(ByVal sheetName As String, ByVal formattedColumns As String)
  Call FormatColumns(sheetName, formattedColumns, xlTextFormat)
End Sub

Public Sub FormatColumnsAsGeneral(ByVal sheetName As String, ByVal formattedColumns As String)
  Call FormatColumns(sheetName, formattedColumns, xlGeneralFormat)
End Sub

Public Sub FormatColumnsAsDate(ByVal sheetName As String, ByVal formattedColumns As String)
  Call FormatColumns(sheetName, formattedColumns, xlYMDFormat)
End Sub

Public Sub EnableEvents()
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.DisplayAlerts = True
  Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub DisableEvents()
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.DisplayStatusBar = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub RefreshPivotTables()
Dim pt As PivotTable
Dim ws As Worksheet
    
  For Each ws In currentWorkBook.Worksheets
    For Each pt In ws.PivotTables
      pt.RefreshTable
    Next pt
  Next ws
End Sub

Public Function GetLastRow(ByVal sheetName As String, ByVal searchColumn As Long) As Long
  GetLastRow = currentWorkBook.Sheets(sheetName).Cells(currentWorkBook.Sheets(sheetName).Rows.Count, searchColumn).End(xlUp).Row
End Function

Public Function GetLastColumn(ByVal sheetName As String, ByVal searchRow As Long) As Long
  GetLastColumn = currentWorkBook.Sheets(sheetName).Cells(searchRow, currentWorkBook.Sheets(sheetName).columns.Count).End(xlToLeft).Column
End Function

Public Sub AddR1C1Formula(ByVal sheetName As String, ByVal startCell As String, ByVal endCell As String, ByVal formulaText As String)
  Sheets(sheetName).Range(startCell & ":" & endCell).FormulaR1C1 = formulaText
End Sub

Public Sub DeleteEmptyRows(ByVal sheetName As String, ByVal columnRange As String)
  On Error Resume Next
    Sheets(sheetName).Select
    columns(columnRange).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
  On Error GoTo 0
End Sub

Public Sub ClearColumns(ByVal sheetName As String, ByVal rangeString As String)
  On Error Resume Next
    Sheets(sheetName).Select
    columns(rangeString).Select
    Selection.Clear
  On Error GoTo 0
End Sub

Public Sub AddBorders(ByVal sheetName As String, ByVal rangeString As String)
    Sheets(sheetName).Select
    
    Range(rangeString).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub SetColumnsColor(ByVal sheetName As String, ByVal columnsRange As String, ByVal foreColorIndex As Integer, ByVal backColorIndex As Integer)
    Sheets(sheetName).Select
    columns(columnsRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = backColorIndex
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = foreColorIndex
        .TintAndShade = 0
    End With
End Sub


Public Sub SaveFileToLocation(ByVal fileName As String, ByVal filePath As String, ByVal wantDatePostfix)
Dim FName As String
Dim FPath As String
    
  FName = fileName
  If (wantDatePostfix) Then
  FName = fileName & " " & Format(Now(), "yyyy-mm-dd hh mm")
  End If
  FPath = filePath
  currentWorkBook.SaveAs fileName:=FPath & "\" & FName, FileFormat:=xlExcel12, AccessMode:=xlExclusive
End Sub



Private Sub Class_Initialize()
  Set currentWorkBook = ActiveWorkbook
End Sub

Private Sub Class_Terminate()
  Call EnableEvents
End Sub

Private Sub FormatColumns(ByVal sheetName As String, ByVal formattedColumns As String, ByVal resultType As Long)
Dim i As Integer
Dim columns() As String

  If formattedColumns = "" Then
    MsgBox "Nie wskazano kolumn"
    Exit Sub
  End If
  
  If Not SheetExist(sheetName) Then
    MsgBox "Arkusz [" & sheetName & "] nie istnieje w pliku [" & currentWorkBook.Name & "]"
    Exit Sub
  End If
  
  columns = Split(formattedColumns, ";")
  
  For i = LBound(columns) To UBound(columns)
    On Error Resume Next
    currentWorkBook.Sheets(sheetName).Range(columns(i) & ":" & columns(i)).TextToColumns _
      Destination:=Range(columns(i) & "1"), _
      DataType:=xlFixedWidth, _
      FieldInfo:=Array(0, resultType), _
      TrailingMinusNumbers:=True
  Next i
End Sub

Private Function SheetExist(ByVal sheetName As String) As Boolean
Dim result As Boolean
Dim sheet As Worksheet

  result = False

  For Each sheet In currentWorkBook.Worksheets
    If UCase(sheetName) = UCase(sheet.Name) Then
      result = True
      Exit For
    End If
  Next

SheetExist = result
End Function




