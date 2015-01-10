VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} chartform 
   Caption         =   "XL Point Manipulator"
   ClientHeight    =   12270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3510
   OleObjectBlob   =   "chartform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "chartform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cb1Color_Click()
PassAction = 1
Call PerformAction
End Sub

Private Sub cb2ShowHide_Click()
PassAction = 2
Call PerformAction
End Sub

Private Sub cb3Axis_Click()
PassAction = 3
Call PerformAction
End Sub

Private Sub cb4LabelID_Click()
PassAction = 4
Call PerformAction
End Sub

Private Sub cb5LabelCustom_Click()
PassAction = 5
Call PerformAction
End Sub

Private Sub cb6Shape_Click()
PassAction = 6
Call PerformAction
End Sub

Private Sub cb7Grow_Click()
PassAction = 7
Call PerformAction
End Sub

Private Sub cb9Mark_Click()
PassAction = 9
Call PerformAction
End Sub

'Private Sub cbChartRestore_Click()
'
'Application.ScreenUpdating = False
'ThisWorkbook.Sheets("Hidden").Activate
'Dim tc2 As ChartObject
'For Each tc2 In ActiveSheet.ChartObjects
'tc2.Copy
'Next
'SavedChartWorkbook.Sheets(SavedChartSheetNum).Activate
'ActiveSheet.Paste
'Application.ScreenUpdating = True
'MsgBox "Saved chart has been pasted; if you don't see it, it might be in the upper left of the worksheet or elsewhere on the screen."
'
'End Sub
'
'Private Sub cbChartStore_Click()
'Set SavedChart = CurChart
'SavedChartSheetNum = ActiveSheet.Name
'Set SavedChartWorkbook = Application.ActiveWorkbook
'Application.ScreenUpdating = False
'SavedChart.ChartArea.Copy
'ThisWorkbook.Sheets("Hidden").Activate
'Dim tc1 As ChartObject
'For Each tc1 In ActiveSheet.ChartObjects
'tc1.Delete
'Next
'Range("A1").Select
'ActiveSheet.Paste
'SavedChartWorkbook.Sheets(SavedChartSheetNum).Activate
'Application.ScreenUpdating = True
'End Sub



Private Sub cbColorRand_Click()
TheColor(CurColor) = RGB(ArrayR(Int(402 * Rnd() + 1)), ArrayG(Int(402 * Rnd() + 1)), ArrayB(Int(402 * Rnd() + 1)))
If CurColor = 1 Then
    chartform.frCol1.BackColor = TheColor(CurColor)
    Else
    chartform.frCol2.BackColor = TheColor(CurColor)
End If
End Sub

Private Sub cbIDRangeSel_Click()
Set IDRange = Application.InputBox(Prompt:="Select the range containing Point IDs, WITHOUT A HEADER CELL. Unexpected things can happen if this range is not exactly parallel to that of your data.", Title:="Range Select", Type:=8)
If IDRange.Columns.Count > 1 Then MsgBox "You have selected a range more than one column wide. Only the leftmost selected column will be used."
chartform.capCPID.Caption = IDRange.Address
End Sub

 'Private Sub cbLabelCol_Click() '''''''''''''''''''''''''''''''''''''
 'If ActiveChart Is Nothing Then
 'MsgBox "Please select a chart and try again.", vbExclamation, "No Chart Selected"
 'Else
 'Dim TempRange As Range
 'Set TempRange = Application.InputBox(Prompt:="Please Select any cell in the column containing labels", Title:="Range Select", Type:=8)
 'LabelColumn = Evaluate("substitute(address(1, " & TempRange.Column & ", 4), ""1"", """")")
 'Dim Counter As Integer, ChartName As String, xVals As String, zVals As String
 'xVals = ActiveChart.SeriesCollection(1).Formula
 ''Parse the range for the data from xVals.
 'Dim TempText As String
 '' TempText = ActiveChart.SeriesCollection(1).Formula
 'Dim vArgs As Variant
 'vArgs = Split(xVals, ",")
 'xVals = Right(vArgs(1), Len(vArgs(1)) - InStr(vArgs(1), "!"))
 'Dim xCol As String
 'Dim ff As Integer, fff As Integer
 'ff = InStr(xVals, "$")
 'fff = InStr(ff + 1, xVals, "$")
 'xCol = Mid(xVals, ff, fff - ff + 1)
 'zVals = Replace(xVals, xCol, "$" & LabelColumn & "$")
 ''Add Labels from Column but hide them
 '      For Counter = 1 To Range(xVals).Cells.Count
 '      If ActiveChart.SeriesCollection(1).Points(Counter).HasDataLabel = False Or cbRew = True Then
 '      If Range(zVals).Cells(Counter, 1).Value <> "" Or cbRew = True Then
 '                 ActiveChart.SeriesCollection(1).Points(Counter).HasDataLabel = _
 '         True
 '      ActiveChart.SeriesCollection(1).Points(Counter).DataLabel.Text = _
 '         Range(zVals).Cells(Counter, 1).Value
 '
 '      End If
 '      End If
 '   Next Counter
 'End If
 'End Sub

Private Sub obclick_Click()
Call InitChartEvents
End Sub




Private Sub cbUndo_Click()
Call TheUndo
End Sub


Private Sub CheckBox1_Click()

End Sub

Private Sub chkWarn_Click()
TheWarning = Not TheWarning
End Sub

Private Sub CommandButton3_Click() 'Add new series
If TheWarning = True Then MsgBox "Note that error checking has been disabled for this new series procedure, so unexpected user input may result in strange things happening or nothing happening. This series addition can only be undone by hitting the Undo button right away before any other actions buttons are pressed."
On Error Resume Next
    With ActiveChart.SeriesCollection.NewSeries
        .XValues = Application.InputBox(Prompt:="Select the range containing X values, WITHOUT A HEADER CELL.", Title:="X Values", Type:=8)
        .Values = Application.InputBox(Prompt:="Select the range containing Y values, matching the size of the X value range.", Title:="Y Values", Type:=8)
        .Name = Application.InputBox(Prompt:="Select the range containing the series name. Normally this is the header cell of the Y values.", Title:="Y Values", Type:=8)
    End With
On Error GoTo 0
AddSeriesUndo = True
AddSeriesUndoNum = ActiveChart.SeriesCollection.Count
End Sub

Private Sub frCol1_Click()
    TheColor(1) = ChooseColorDialog(1)
    chartform.frCol1.BackColor = TheColor(1)
End Sub

Private Sub frCol2_Click()
    TheColor(2) = ChooseColorDialog(1)
    chartform.frCol2.BackColor = TheColor(2)
End Sub

Private Sub obCol1_Click()
CurColor = 1
End Sub

Private Sub obCol2_Click()
CurColor = 2
End Sub

Private Sub obcolor_Click()
ActionSelection = 1
End Sub

Private Sub obctrlq_Click()
Call CancelChartEvents
End Sub

Private Sub oblabel_Click()
ActionSelection = 2
End Sub

Private Sub obRem_Click()
ActionSelection = 3
End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub obSelect_Click()

End Sub

Private Sub vbEnd_Click()
ReDim clsEventCharts(1 To ActiveSheet.ChartObjects.Count)
Unload chartform
End Sub

