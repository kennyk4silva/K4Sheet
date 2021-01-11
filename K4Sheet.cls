VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Classes")
'@PredeclaredId
Option Explicit
Implements ISheetWorker

Private Type TSheetWorker
  OptimizePerformance As Boolean
  CurrentScreenUpdating As Boolean
  CurrentCalculation As Excel.XlCalculation
  CurrentEnableEvents As Boolean
End Type

Private This As TSheetWorker

Public Function Create(Optional ByVal Optimize As Boolean) As ISheetWorker
  With New SheetWorker
    .CurrentScreenUpdating = Application.ScreenUpdating
    .CurrentCalculation = Application.Calculation
    .CurrentEnableEvents = Application.EnableEvents
    .OptimizePerformance = Optimize
    If .OptimizePerformance Then
      Application.ScreenUpdating = False
      Application.Calculation = xlCalculationManual
      Application.EnableEvents = False
    End If
    Set Create = .Self 'Returns the newly created instance.
  End With
End Function

Public Property Get Self() As ISheetWorker
  Set Self = Me
End Property

Public Property Get OptimizePerformance() As Boolean
  OptimizePerformance = This.OptimizePerformance
End Property

Public Property Get CurrentScreenUpdating() As Boolean
  CurrentScreenUpdating = This.CurrentScreenUpdating
End Property

Public Property Let CurrentScreenUpdating(Value As Boolean)
  This.CurrentScreenUpdating = Value
End Property

Private Property Get ISheetWorker_CurrentScreenUpdating() As Boolean
  ISheetWorker_CurrentScreenUpdating = CurrentScreenUpdating
End Property

Public Property Get CurrentCalculation() As Excel.XlCalculation
  CurrentCalculation = This.CurrentCalculation
End Property

Public Property Let CurrentCalculation(Value As Excel.XlCalculation)
  This.CurrentCalculation = Value
End Property

Private Property Get ISheetWorker_CurrentCalculation() As XlCalculation
  ISheetWorker_CurrentCalculation = CurrentCalculation
End Property

Public Property Let CurrentEnableEvents(Value As Boolean)
  This.CurrentEnableEvents = Value
End Property

Public Property Get CurrentEnableEvents() As Boolean
  CurrentEnableEvents = This.CurrentEnableEvents
End Property

Private Property Get ISheetWorker_CurrentEnableEvents() As Boolean
  ISheetWorker_CurrentEnableEvents = CurrentEnableEvents
End Property

Public Property Let OptimizePerformance(Value As Boolean)
  This.OptimizePerformance = Value
End Property

Private Property Get ISheetWorker_OptimizePerformance() As Boolean
  ISheetWorker_OptimizePerformance = OptimizePerformance
End Property

Private Sub Class_Terminate()
  Application.ScreenUpdating = CurrentScreenUpdating
  Application.Calculation = CurrentCalculation
  Application.EnableEvents = CurrentEnableEvents
End Sub

Private Sub ISheetWorker_CopyDataToRangeByColumns(ByVal DestinyRange As Range, _
ByVal SourceData As BetterArray, ParamArray ColumnIndexes() As Variant)
  Dim ColRangeIndex As Long
  ColRangeIndex = 1
  Dim TempColumn As BetterArray
  Set TempColumn = New BetterArray
  Dim ColIndex As Long
  For ColIndex = 0 To UBound(ColumnIndexes)
    If ColumnIndexes(ColIndex) > 0 Then
      TempColumn.Items = SourceData.ExtractSegment(, ColumnIndexes(ColIndex))
      TempColumn.ToExcelRange DestinyRange.Cells(1, ColRangeIndex)
    End If
    ColRangeIndex = ColRangeIndex + 1
  Next ColIndex
End Sub

Public Function ISheetWorker_RowIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal RowOffset As Long, Optional ByRef Sheet As Worksheet) As Long
  If Sheet Is Nothing Then
    ISheetWorker_RowIndexFromNamedCell = ActiveSheet.Range(CellName).Row + RowOffset
  Else
    ISheetWorker_RowIndexFromNamedCell = Sheet.Range(CellName).Row + RowOffset
  End If
End Function

Private Function ISheetWorker_ColumnIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal ColumnOffset As Long, Optional Sheet As Worksheet) As Long
  If Sheet Is Nothing Then
    ISheetWorker_ColumnIndexFromNamedCell = ActiveSheet.Range(CellName).Column + _
    ColumnOffset
  Else
    ISheetWorker_ColumnIndexFromNamedCell = Sheet.Range(CellName).Column + _
    ColumnOffset
  End If
End Function

Private Function ISheetWorker_RowsColumnsDeltaByNamedCells(ByRef Sheet As Worksheet, _
ByVal FirstRowCellName As String, ByVal SecondRowCellName As String, _
ByVal FirstColumnCellName As String, ByVal SecondColumnCellName As String) As Long()
  On Error GoTo CleanFail
  
  Dim DeltaRowsColumns(1 To 2) As Long
  With Sheet
    DeltaRowsColumns(1) = Math.Abs(.Range(FirstRowCellName).Row - _
                                   .Range(SecondRowCellName).Row) - 1
    DeltaRowsColumns(2) = Math.Abs(.Range(FirstColumnCellName).Column - _
                                   .Range(SecondColumnCellName).Column) - 1
  End With
  ISheetWorker_RowsColumnsDeltaByNamedCells = DeltaRowsColumns

CleanFail:
End Function

Private Function ISheetWorker_ArrayWithinTitlesWithColumnCriteria( _
ByRef Sheet As Worksheet, ByVal EntryArray As BetterArray, _
ByVal EntitledColumn As Long, ByVal Title As String, ByVal CriteriaColumn As Long, _
ByVal CriteriaText As String) As Variant()
  On Error GoTo CleanFail
  
  Dim ColumnWithTitle As BetterArray
  Set ColumnWithTitle = New BetterArray
  ColumnWithTitle.Items = EntryArray.ExtractSegment(ColumnIndex:=EntitledColumn)
  Dim ColumnWithCriteria As BetterArray
  Set ColumnWithCriteria = New BetterArray
  ColumnWithCriteria.Items = EntryArray.ExtractSegment(ColumnIndex:=CriteriaColumn)
  Dim FirstRow As Long
  FirstRow = ColumnWithTitle.IndexOf(Title)
  Dim LastRow As Long
  LastRow = ColumnWithCriteria.IndexOf(CriteriaText, FirstRow + 1, _
  CompType:=CT_LIKENESS)
  ISheetWorker_ArrayWithinTitlesWithColumnCriteria = EntryArray.Slice(FirstRow + 1, _
  LastRow)
  
CleanFail:
End Function

Private Sub ISheetWorker_UpdateColumnsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal ColOffset As Long, _
Optional ByVal ColsQuantity As Long = 1&)
  With Sheet
    If ColsQuantity > 0 Then
      .Columns(.Range(CellName).Column + ColOffset).Resize(, ColsQuantity).Insert _
      xlShiftToRight
    ElseIf ColsQuantity < 0 Then
      .Columns(.Range(CellName).Column + ColOffset).Resize(, ColsQuantity).Delete _
      xlShiftToRight
    End If
  End With
End Sub

Private Sub ISheetWorker_UpdateRowsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal RowOffset As Long, _
Optional ByVal RowsQuantity As Long = 1&)
  With Sheet
    If RowsQuantity > 0 Then
      .Rows(.Range(CellName).Row + RowOffset).Resize(RowsQuantity).Insert
    ElseIf RowsQuantity < 0 Then
      .Rows(.Range(CellName).Row + RowOffset).Resize(RowsQuantity).Delete
    End If
  End With
End Sub

Private Sub ISheetWorker_CenterTextInRange(ByVal Text As String, _
ByRef TextRange As Range, Optional ByVal ColumnsNumber As Long = 1, _
Optional ByVal MergeAndCenter As Boolean = True, _
Optional ByVal OutsideBorder As Boolean = False, Optional ColumnsWidth As Double)
  On Error GoTo CleanFail
  Application.DisplayAlerts = False
  With TextRange.Resize(ColumnSize:=ColumnsNumber)
    .VerticalAlignment = Excel.xlCenter
    .Cells(1, 1) = Text
    If MergeAndCenter Then
      .Merge
      .HorizontalAlignment = Excel.xlHAlignCenter
    Else
      .UnMerge
      .HorizontalAlignment = xlCenterAcrossSelection
      .Borders(Excel.xlInsideVertical).LineStyle = Excel.xlLineStyleNone
    End If
    If OutsideBorder Then
      .BorderAround Excel.xlContinuous, Excel.xlThin
    Else
      .Borders(Excel.xlEdgeTop).LineStyle = Excel.xlLineStyleNone
      .Borders(Excel.xlEdgeBottom).LineStyle = Excel.xlLineStyleNone
      .Borders(Excel.xlEdgeLeft).LineStyle = Excel.xlLineStyleNone
      .Borders(Excel.xlEdgeRight).LineStyle = Excel.xlLineStyleNone
    End If
    If ColumnsWidth > 0 Then
      .ColumnWidth = ColumnsWidth
    End If
  End With
CleanFail:
  Application.DisplayAlerts = True
End Sub

Private Sub ISheetWorker_FillRangeWithSeries(ByRef SeriesRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
Optional ByVal Prefix As String, Optional ByVal Sufix As String)
  Dim SeriesArray() As Variant
  Dim Index As Long
  If Direction = Horizontal Then
    ReDim SeriesArray(1 To 1, 1 To CellsQuantity)
    For Index = 1 To UBound(SeriesArray, 2)
      SeriesArray(1, Index) = Prefix & Index & Sufix
    Next Index
    SeriesRange.Resize(1, CellsQuantity) = SeriesArray
  Else
    ReDim SeriesArray(1 To CellsQuantity, 1 To 1)
    For Index = 1 To UBound(SeriesArray, 1)
      SeriesArray(Index, 1) = Prefix & Index & Sufix
    Next Index
    SeriesRange.Resize(CellsQuantity, 1) = SeriesArray
  End If
End Sub

Private Sub ISheetWorker_RangeFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal StartingRef As Long, ByVal RefLenght As Long, _
Optional FunctionName As String = "SUM")
  Dim Formula As String
  Dim Index As Long
  If Direction = Horizontal Then
    Index = FormulaRange.Cells(1, 1).Row
    Formula = "=" & FunctionName & "(R[" & (StartingRef - Index) & "]C:R[" & _
    (StartingRef + RefLenght - Index - 1) & "]C)"
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = Formula
  Else
    Index = FormulaRange.Cells(1, 1).Column
    Formula = "=" & FunctionName & "(RC[" & (StartingRef - Index) & "]:RC[" & _
    (StartingRef + RefLenght - Index - 1) & "])"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
  End If
End Sub

Private Sub ISheetWorker_PartialSumFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal ColOrRowIndex As Long)
  Dim Offset As Long
  Offset = ColOrRowIndex - FormulaRange.Row
  If Direction = Horizontal Then
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = "=R[" & Offset & "]C+RC[-1]"
    FormulaRange.Cells(1, 1).FormulaR1C1 = "=R[" & Offset & "]C"
  Else
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = "=RC[" & Offset & "]+R[-1]C"
    FormulaRange.Cells(1, 1).FormulaR1C1 = "=RC[" & Offset & "]"
  End If
End Sub

Private Sub ISheetWorker_PercentFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByRef CellWithTotal As Range, _
ByVal Direction As RangeDirection, ByVal ColOrRowIndex As Long, _
Optional NumOfDecimals As Long = 2&)
  Dim TotalAbsoluteRef As String
  TotalAbsoluteRef = "R" & CellWithTotal.Row & "C" & CellWithTotal.Column
  Dim Formula As String
  If Direction = Horizontal Then
    Dim IndexRowOffset As Long
    IndexRowOffset = ColOrRowIndex - FormulaRange.Cells(1, 1).Row
    Formula = "=ROUND(R[" & IndexRowOffset & "]C/" & TotalAbsoluteRef & "," & _
    NumOfDecimals & ")"
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = Formula
  Else
    Dim IndexColOffset As Long
    IndexColOffset = ColOrRowIndex - FormulaRange.Cells(1, 1).Column
    Formula = "=ROUND(RC[" & IndexColOffset & "]/" & TotalAbsoluteRef & "," & _
    NumOfDecimals & ")"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
  End If
  FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "0.00%"
End Sub

Private Sub ISheetWorker_FactorFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal ColOrRowIndex As Long, ByVal Factor As Double, _
Optional NumOfDecimals As Long = 2&)
  Dim Offset As Long
  Dim Formula As String
  If Direction = Horizontal Then
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Row
    Formula = "=ROUND(" & Factor & "*R[" & Offset & "]C," & NumOfDecimals & ")"
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = Formula
  Else
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Column
    Formula = "=ROUND(" & Factor & "*RC[" & Offset & "]," & NumOfDecimals & ")"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
  End If
  FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
End Sub

Private Sub ISheetWorker_CellsSumFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ParamArray RowOrColIndexes() As Variant)
  Dim Formula As String
  Formula = "="
  Dim Offset As Long
  Dim Index As Long
  If Direction = Horizontal Then
    For Index = 0 To UBound(RowOrColIndexes)
      Offset = RowOrColIndexes(Index) - FormulaRange.Cells(1, 1).Row
      Formula = Formula & "R[" & Offset & "]C+"
    Next Index
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = Left(Formula, Len(Formula) - 1)
    FormulaRange.Resize(1, CellsQuantity).NumberFormat = "#,##0.00"
  Else
    For Index = 0 To UBound(RowOrColIndexes)
      Offset = RowOrColIndexes(Index) - FormulaRange.Cells(1, 1).Column
      Formula = Formula & "RC[" & Offset & "]+"
    Next Index
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Left(Formula, Len(Formula) - 1)
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
  End If
End Sub