VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISheetWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Abstract")
Option Explicit

Public Enum RangeDirection
  Horizontal
  Vertical
End Enum

Public Property Get OptimizePerformance() As Boolean: End Property
Public Property Get CurrentScreenUpdating() As Boolean: End Property
Public Property Get CurrentCalculation() As Excel.XlCalculation: End Property
Public Property Get CurrentEnableEvents() As Boolean: End Property

Public Function RowsColumnsDeltaByNamedCells(ByRef Sheet As Worksheet, _
ByVal FirstRowCellName As String, ByVal SecondRowCellName As String, _
ByVal FirstColumnCellName As String, ByVal SecondColumnCellName As String) As Long()
End Function

Public Function RowIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal RowOffset As Long, Optional ByRef Sheet As Worksheet) As Long
End Function

Public Function ColumnIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal ColumnOffset As Long, Optional ByRef Sheet As Worksheet) As Long
End Function

Public Function ArrayWithinTitlesWithColumnCriteria(ByRef Sheet As Worksheet, _
ByVal EntryArray As BetterArray, ByVal EntitledColumn As Long, _
ByVal Title As String, ByVal CriteriaColumn As Long, _
ByVal CriteriaText As String) As Variant()
End Function

Public Sub CopyDataToRangeByColumns(ByVal DataRange As Range, _
ByVal Data As BetterArray, ParamArray ColumIndexes() As Variant)
End Sub

Public Sub UpdateRowsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal RowOffset As Long, _
Optional ByVal RowsQuantity As Long = 1&)
End Sub

Public Sub UpdateColumnsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal ColOffset As Long, _
Optional ByVal ColsQuantity As Long = 1&)
End Sub

Public Sub CenterTextInRange(ByVal Text As String, ByRef TextRange As Range, _
Optional ByVal ColumnsNumber As Long = 1, Optional ByVal MergeAndCenter As Boolean = True, _
Optional ByVal BorderAround As Boolean = False, Optional ColumnsWidth As Double)
End Sub

Public Sub FillRangeWithSeries(ByRef SeriesRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
Optional ByVal Prefix As String, Optional ByVal Sufix As String)
End Sub

Public Sub RangeFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal StartingRef As Long, ByVal RefLenght As Long, _
Optional FunctionName As String = "SUM")
End Sub

Public Sub PartialSumFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal ColOrRowIndex As Long)
End Sub

