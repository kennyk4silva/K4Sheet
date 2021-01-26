VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K4Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A class that provides features to facilitate common Worksheets programming tasks."
'@Folder("Classes")
'@ModuleDescription("A class that provides features to facilitate common Worksheets programming tasks.")
'@PredeclaredId

'MIT License
'
'Copyright (c) 2021 Kenny Silva
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Option Explicit
Implements IK4Sheet

'**********
'* Fields *
'**********
Private Type TK4Sheet
    OptimizationMode As k4OptimizationMode
    CurrentScreenUpdating As Boolean
    CurrentCalculation As Excel.XlCalculation
    CurrentEnableEvents As Boolean
End Type

Private This As TK4Sheet

'@Description("Constructor method.")
Public Function Create(Optional ByVal Optimization As k4OptimizationMode) As IK4Sheet
Attribute Create.VB_Description = "Constructor method."
    With New K4Sheet
        'Stores the current values of the Excel optimization properties.
        .CurrentScreenUpdating = Application.ScreenUpdating
        .CurrentCalculation = Application.Calculation
        .CurrentEnableEvents = Application.EnableEvents
        'Updates the optimization mode property with the function argument.
        .OptimizationMode = Optimization
        'Disables Excel optimization properties according to function argument.
        If Optimization <> k4OptimizationMode.k4NotOptimize Then
            Application.ScreenUpdating = False
            Application.Calculation = Excel.XlCalculation.xlCalculationManual
            If Optimization = k4OptimizationMode.k4WithoutEvents Then
                Application.EnableEvents = False
            End If
        End If
        'Returns the newly created instance.
        Set Create = .Self
    End With
End Function

'Description("Gets an instance of the current class.")
Public Property Get Self() As IK4Sheet
    Set Self = Me
End Property

'Description("Gets the current optimization mode.")
Public Property Get OptimizationMode() As k4OptimizationMode
    OptimizationMode = This.OptimizationMode
End Property

'Description("Gets the current optimization mode (interface).")
Private Property Get IK4Sheet_OptimizationMode() As k4OptimizationMode
    IK4Sheet_OptimizationMode = OptimizationMode
End Property

'Description("Sets the optimization mode.")
Public Property Let OptimizationMode(Value As k4OptimizationMode)
    This.OptimizationMode = Value
End Property

'Description("Gets the current Excel Application ScreenUpdating property.")
Public Property Get CurrentScreenUpdating() As Boolean
    CurrentScreenUpdating = This.CurrentScreenUpdating
End Property

'Description("Gets the current Excel Application ScreenUpdating property (interface).")
Private Property Get IK4Sheet_CurrentScreenUpdating() As Boolean
    IK4Sheet_CurrentScreenUpdating = CurrentScreenUpdating
End Property

'Description("Sets the Excel Application ScreenUpdating property.")
Public Property Let CurrentScreenUpdating(Value As Boolean)
    This.CurrentScreenUpdating = Value
End Property

'Description("Gets the current Excel Application Calculation property.")
Public Property Get CurrentCalculation() As Excel.XlCalculation
    CurrentCalculation = This.CurrentCalculation
End Property

'Description("Gets the current Excel Application Calculation property (interface).")
Private Property Get IK4Sheet_CurrentCalculation() As Excel.XlCalculation
    IK4Sheet_CurrentCalculation = CurrentCalculation
End Property

'Description("Sets the Excel Application Calculation property.")
Public Property Let CurrentCalculation(Value As Excel.XlCalculation)
    This.CurrentCalculation = Value
End Property

'Description("Gets the current Excel Application EnableEvents property.")
Public Property Let CurrentEnableEvents(Value As Boolean)
    This.CurrentEnableEvents = Value
End Property

'Description("Gets the current Excel Application EnableEvents property (interface).")
Private Property Get IK4Sheet_CurrentEnableEvents() As Boolean
    IK4Sheet_CurrentEnableEvents = CurrentEnableEvents
End Property

'Description("Sets the Excel Application EnableEvents property.")
Public Property Get CurrentEnableEvents() As Boolean
    CurrentEnableEvents = This.CurrentEnableEvents
End Property

'Description("Terminate event (after the object is unloaded).")
Private Sub Class_Terminate()
    'Restore the original Excel optimization properties before instanciating the class.
    Application.ScreenUpdating = CurrentScreenUpdating
    Application.Calculation = CurrentCalculation
    Application.EnableEvents = CurrentEnableEvents
End Sub

Private Sub IK4Sheet_CopyDataToRangeByColumns(ByRef DestinyRange As Range, _
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
      If TempColumn.EveryType("Double") Then
        DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length).NumberFormat = "#,##0.00"
        DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length).Value = _
        Application.Round(DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length), 2)
      End If
    End If
    ColRangeIndex = ColRangeIndex + 1
  Next ColIndex
End Sub

Private Sub IK4Sheet_CopyDataToRangeByColumnsNuevo(ByRef DestinyRange As Range, _
ByVal SourceData As BetterArray, ByVal ColumnIndexes As String)
  Dim ColRangeIndex As Long
  ColRangeIndex = 1
  Dim TempColumn As BetterArray
  Set TempColumn = New BetterArray
  Dim ColIndex As Long
  Dim Codigos() As String
  Codigos = Split(ColumnIndexes, ",")
  For ColIndex = LBound(Codigos) To UBound(Codigos)
    If Codigos(ColIndex) > 0 Then
      TempColumn.Items = SourceData.ExtractSegment(, Codigos(ColIndex))
      TempColumn.ToExcelRange DestinyRange.Cells(1, ColRangeIndex)
      If TempColumn.EveryType("Double") Then
        DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length).NumberFormat = "#,##0.00"
        DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length).Value = _
        Application.Round(DestinyRange.Cells(1, ColRangeIndex).Resize(TempColumn.Length), 2)
      End If
    End If
    ColRangeIndex = ColRangeIndex + 1
  Next ColIndex
End Sub

Public Function IK4Sheet_RowIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal RowOffset As Long, Optional ByRef Sheet As Worksheet) As Long
  If Sheet Is Nothing Then
    IK4Sheet_RowIndexFromNamedCell = ActiveSheet.Range(CellName).Row + RowOffset
  Else
    IK4Sheet_RowIndexFromNamedCell = Sheet.Range(CellName).Row + RowOffset
  End If
End Function

Private Function IK4Sheet_ColumnIndexFromNamedCell(ByVal CellName As String, _
Optional ByVal ColumnOffset As Long, Optional Sheet As Worksheet) As Long
  If Sheet Is Nothing Then
    IK4Sheet_ColumnIndexFromNamedCell = ActiveSheet.Range(CellName).Column + _
    ColumnOffset
  Else
    IK4Sheet_ColumnIndexFromNamedCell = Sheet.Range(CellName).Column + _
    ColumnOffset
  End If
End Function

Private Function IK4Sheet_RowsColumnsDeltaByNamedCells(ByRef Sheet As Worksheet, _
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
  IK4Sheet_RowsColumnsDeltaByNamedCells = DeltaRowsColumns

CleanFail:
End Function

Private Function IK4Sheet_ArrayWithinTitlesWithColumnCriteria( _
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
  LastRow = ColumnWithCriteria.IndexOf(CriteriaText, FirstRow + 1, CompType:=CT_LIKENESS)
  If LastRow = -9999 Then
    LastRow = ColumnWithCriteria.Length
  End If
  IK4Sheet_ArrayWithinTitlesWithColumnCriteria = EntryArray.Slice(FirstRow + 1, _
  LastRow)
  
CleanFail:
End Function

Private Sub IK4Sheet_UpdateColumnsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal ColOffset As Long, _
Optional ByVal ColsQuantity As Long = 1&)
  With Sheet
    If ColsQuantity > 0 Then
      .Columns(.Range(CellName).Column + ColOffset).Resize(, ColsQuantity).Insert _
      xlShiftToRight
    ElseIf ColsQuantity < 0 Then
      .Columns(.Range(CellName).Column + ColOffset).Resize(, Abs(ColsQuantity)).Delete _
      xlShiftToLeft
    End If
  End With
End Sub

Private Sub IK4Sheet_UpdateRowsQtyFromNamedCell(ByRef Sheet As Worksheet, _
ByVal CellName As String, Optional ByVal RowOffset As Long, _
Optional ByVal RowsQuantity As Long = 1&)
  With Sheet
    If RowsQuantity > 0 Then
      .Rows(.Range(CellName).Row + RowOffset).Resize(RowsQuantity).Insert
    ElseIf RowsQuantity < 0 Then
      .Rows(.Range(CellName).Row + RowsQuantity).Resize(Abs(RowsQuantity)).Delete
    End If
  End With
End Sub

Private Sub IK4Sheet_CenterTextInRange(ByVal Text As String, _
ByRef TextRange As Range, Optional ByVal ColumnsNumber As Long = 1, _
Optional ByVal MergeAndCenter As Boolean = True, _
Optional ByVal OutsideBorder As Boolean = False, Optional ColumnsWidth As Double)
  With TextRange.Resize(ColumnSize:=ColumnsNumber)
    .ClearContents
    .VerticalAlignment = Excel.xlCenter
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
    .Cells(1, 1) = Text
    If ColumnsWidth > 0 Then
      .ColumnWidth = ColumnsWidth
    End If
  End With
End Sub

Private Sub IK4Sheet_FillRangeWithSeries(ByRef SeriesRange As Range, _
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

Private Sub IK4Sheet_RangeFormulaR1C1(ByRef FormulaRange As Range, _
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
    FormulaRange.Resize(1, CellsQuantity).NumberFormat = "#,##0.00"
    FormulaRange.Resize(1, CellsQuantity).HorizontalAlignment = xlHAlignCenter
  Else
    Index = FormulaRange.Cells(1, 1).Column
    Formula = "=" & FunctionName & "(RC[" & (StartingRef - Index) & "]:RC[" & _
    (StartingRef + RefLenght - Index - 1) & "])"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
    FormulaRange.Resize(CellsQuantity, 1).HorizontalAlignment = xlHAlignCenter
  End If
End Sub

Private Sub IK4Sheet_PartialSumFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal ColOrRowIndex As Long)
  Dim Offset As Long
  If Direction = Horizontal Then
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Row
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = "=R[" & Offset & "]C+RC[-1]"
    FormulaRange.Resize(1, CellsQuantity).NumberFormat = "#,##0.00"
    FormulaRange.Resize(1, CellsQuantity).HorizontalAlignment = xlHAlignCenter
    FormulaRange.Cells(1, 1).FormulaR1C1 = "=R[" & Offset & "]C"
  Else
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Column
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = "=RC[" & Offset & "]+R[-1]C"
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
    FormulaRange.Resize(CellsQuantity, 1).HorizontalAlignment = xlHAlignCenter
    FormulaRange.Cells(1, 1).FormulaR1C1 = "=RC[" & Offset & "]"
  End If
End Sub

Private Sub IK4Sheet_PercentFormulaR1C1(ByRef FormulaRange As Range, _
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
    FormulaRange.Resize(1, CellsQuantity).NumberFormat = "0.00%"
    FormulaRange.Resize(1, CellsQuantity).HorizontalAlignment = xlHAlignCenter
  Else
    Dim IndexColOffset As Long
    IndexColOffset = ColOrRowIndex - FormulaRange.Cells(1, 1).Column
    Formula = "=ROUND(RC[" & IndexColOffset & "]/" & TotalAbsoluteRef & "," & _
    NumOfDecimals & ")"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "0.00%"
    FormulaRange.Resize(CellsQuantity, 1).HorizontalAlignment = xlHAlignCenter
  End If
End Sub

Private Sub IK4Sheet_FactorFormulaR1C1(ByRef FormulaRange As Range, _
ByVal CellsQuantity As Long, ByVal Direction As RangeDirection, _
ByVal ColOrRowIndex As Long, ByVal Factor As Double, _
Optional NumOfDecimals As Long = 2&)
  Dim Offset As Long
  Dim Formula As String
  If Direction = Horizontal Then
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Row
    Formula = "=ROUND(" & Replace(Factor, ",", ".") & "*R[" & Offset & "]C," & NumOfDecimals & ")"
    FormulaRange.Resize(1, CellsQuantity).FormulaR1C1 = Formula
    FormulaRange.Resize(1, CellsQuantity).NumberFormat = "#,##0.00"
    FormulaRange.Resize(1, CellsQuantity).HorizontalAlignment = xlHAlignCenter
  Else
    Offset = ColOrRowIndex - FormulaRange.Cells(1, 1).Column
    Formula = "=ROUND(" & Replace(Factor, ",", ".") & "*RC[" & Offset & "]," & NumOfDecimals & ")"
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Formula
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
    FormulaRange.Resize(CellsQuantity, 1).HorizontalAlignment = xlHAlignCenter
  End If
End Sub

Private Sub IK4Sheet_CellsSumFormulaR1C1(ByRef FormulaRange As Range, _
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
    FormulaRange.Resize(1, CellsQuantity).HorizontalAlignment = xlHAlignCenter
  Else
    For Index = 0 To UBound(RowOrColIndexes)
      Offset = RowOrColIndexes(Index) - FormulaRange.Cells(1, 1).Column
      Formula = Formula & "RC[" & Offset & "]+"
    Next Index
    FormulaRange.Resize(CellsQuantity, 1).FormulaR1C1 = Left(Formula, Len(Formula) - 1)
    FormulaRange.Resize(CellsQuantity, 1).NumberFormat = "#,##0.00"
    FormulaRange.Resize(CellsQuantity, 1).HorizontalAlignment = xlHAlignCenter
  End If
End Sub


Private Sub IK4Sheet_UpdatePieChartSerie(ByVal ChartName As String, _
ByVal XValuesRange As Range, ByVal ValuesRange As Range, ByVal ValuesQty As Long)
  Dim PieChart As ChartObject
  Set PieChart = ActiveSheet.ChartObjects(ChartName)
  Dim ChartSerie As Series
  Set ChartSerie = PieChart.Chart.SeriesCollection(1)
  ChartSerie.XValues = XValuesRange.Resize(ValuesQty)
  ChartSerie.Values = ValuesRange.Resize(ValuesQty)
End Sub

Private Sub IK4Sheet_UpdateClusteredColumnChart(ByVal ChartName As String, _
ByVal SeriesNames As Range, ByVal SeriesNamesQty As Long, ByVal SeriesValues As Range, _
ByVal SeriesValuesQty As Long, ByVal SeriesValuesDirection As RangeDirection, _
ByVal PrimaryAxisTitle As String, ByVal SecondaryAxisTitle As String, _
ParamArray MarkedLines() As Variant)

  Dim ChartObj As Chart
  Set ChartObj = ActiveSheet.ChartObjects(ChartName).Chart
  Dim SeriesObj As Series
  
  For Each SeriesObj In ChartObj.SeriesCollection
    If ChartObj.SeriesCollection.Count > 1 Then
      SeriesObj.Delete
    End If
  Next SeriesObj
  
  Dim Index As Long
  If SeriesValuesDirection = Horizontal Then
  
    ChartObj.SeriesCollection(1).name = "='" & ActiveSheet.name & "'!" & _
    SeriesNames.Cells(1, 1).Resize(1, SeriesNames.Columns.Count).Address
    ChartObj.SeriesCollection(1).Values = SeriesValues.Cells(1, 1).Resize(1, _
    SeriesValuesQty)
    ChartObj.SeriesCollection(1).XValues = _
    SeriesValues.Offset(-1, 0).Resize(1, SeriesValuesQty)
    ChartObj.SeriesCollection(1).ChartType = xlColumnClustered
    
    For Index = 2 To SeriesNamesQty
      Set SeriesObj = ChartObj.SeriesCollection.Add(SeriesValues.Cells(Index, 1).Resize(1, _
      SeriesValuesQty))
      SeriesObj.name = "='" & ActiveSheet.name & "'!" & _
      SeriesNames.Cells(Index, 1).Resize(1, SeriesNames.Columns.Count).Address
      SeriesObj.ChartType = xlColumnClustered
    Next Index
    
  Else
  
    ChartObj.SeriesCollection(1).name = "='" & ActiveSheet.name & "'!" & _
    SeriesNames.Cells(1, 1).Resize(SeriesNames.Rows.Count, 1).Address
    SeriesObj.ChartType = xlColumnClustered
    ChartObj.SeriesCollection(1).Values = SeriesValues.Cells(1, 1).Resize( _
    SeriesValuesQty, 1)
    ChartObj.SeriesCollection(1).XValues = _
    SeriesValues.Offset(0, -1).Resize(SeriesValuesQty, 1)
    ChartObj.SeriesCollection(1).ChartType = xlColumnClustered
    
    For Index = 2 To SeriesNamesQty
      Set SeriesObj = ChartObj.SeriesCollection.Add(SeriesValues.Cells(1, _
      Index).Resize(SeriesValuesQty, 1))
      SeriesObj.name = "='" & ActiveSheet.name & "'!" & _
      SeriesNames.Cells(1, Index).Resize(SeriesNames.Rows.Count, 1).Address
      SeriesObj.ChartType = xlColumnClustered
    Next Index
  End If
  
  For Index = 0 To UBound(MarkedLines)
    With ChartObj.SeriesCollection(MarkedLines(Index))
      .ChartType = xlLineMarkers
      .Smooth = True
      .AxisGroup = 2
      .MarkerStyle = xlMarkerStyleCircle
      If Index = 0 Then .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
      If Index = 1 Then .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
    End With
  Next Index
  
  With ChartObj
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Text = PrimaryAxisTitle
    .Axes(xlValue, xlSecondary).HasTitle = True
    .Axes(xlValue, xlSecondary).AxisTitle.Text = SecondaryAxisTitle
  End With
  
End Sub

Private Sub IK4Sheet_ExtractSeriesColorsToRange(ByVal ChartName As String, _
ByRef DestinyRange As Range, ByVal ExtractionDirection As RangeDirection, _
ByVal SeriesQty As Long)
  Dim ChartObj As Chart
  Set ChartObj = ActiveSheet.ChartObjects(ChartName).Chart
  
  Dim Index As Long
  Dim SeriesColors() As Long
  ReDim SeriesColors(1 To SeriesQty)
  
  For Index = 1 To UBound(SeriesColors)
    SeriesColors(Index) = ChartObj.FullSeriesCollection(Index).Format.Fill.ForeColor.RGB
    ChartObj.FullSeriesCollection(Index).Format.Line.Visible = msoFalse
  Next Index
  
  If ExtractionDirection = Horizontal Then
    For Index = 1 To SeriesQty
      DestinyRange.Offset(0, Index - 1).Interior.Color = SeriesColors(Index)
    Next Index
  Else
    For Index = 1 To SeriesQty
      DestinyRange.Offset(Index - 1, 0).Interior.Color = SeriesColors(Index)
    Next Index
  End If
End Sub

Private Sub IK4Sheet_NumericalCompensation(ByVal DescriptionRange As Range, _
ByVal RowQty As Long, ByRef DataRange As Range, ByVal ColumnQty As Long, _
ByVal DescriptionsToMatch As Variant, ByVal TotalsToMatch As Variant)
  Dim Descriptions As Variant
  Descriptions = DescriptionRange.Resize(RowQty, 1)
  
  Dim MeanDescriptions As BetterArray
  Set MeanDescriptions = New BetterArray
  MeanDescriptions.Items = DescriptionsToMatch
  MeanDescriptions.Transpose
  
  Dim MeanTotals As BetterArray
  Set MeanTotals = New BetterArray
  MeanTotals.Items = TotalsToMatch
  MeanTotals.Transpose

  Dim DataToAdjust As BetterArray
  Set DataToAdjust = New BetterArray
  DataToAdjust.Items = DataRange.Resize(RowQty, ColumnQty).Value
  
  Dim Index As Long
  Dim TotalToCompare As Variant
  Dim TempRow As Range
  Dim TotalByRow As Double
  Dim Delta As Double
  Dim MaxValue As Double
  Dim MaxValueIndex As Long
  For Index = 1 To UBound(Descriptions, 1)
    TotalToCompare = MeanTotals(MeanDescriptions.IndexOf(Descriptions(Index, 1)))
    If IsNumeric(TotalToCompare) Then
      Set TempRow = DataRange.Offset(Index - 1).Resize(, ColumnQty)
      TotalByRow = Application.Sum(TempRow)
      Delta = TotalToCompare - TotalByRow
      MaxValue = Application.Max(TempRow)
      MaxValueIndex = Application.Match(MaxValue, TempRow, 0)
      TempRow.Cells(1, MaxValueIndex) = TempRow.Cells(1, MaxValueIndex) + Delta
    End If
  Next Index
End Sub