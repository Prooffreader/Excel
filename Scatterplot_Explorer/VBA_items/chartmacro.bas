Attribute VB_Name = "chartmacro"
Option Explicit
Public TheColor(1 To 2) As Double
Public CurColor As Integer
Public ActionSelection As Integer
Public xVals As String
Public MarkerArray(9) As Integer
Public CurChartNum As Integer
Public CurSeriesNum As Integer
Public CurPointNum As Integer
Public IDRange As Range
Public SavedChart As Chart
Public SavedChartSheetNum As String
Public SavedChartWorkbook As Workbook
Public clsEventCharts() As New clsEventChart
Public ArrayR(1 To 403) As Integer
Public ArrayG(1 To 403) As Integer
Public ArrayB(1 To 403) As Integer
Public CurSel As Integer '0=Nothing 1=Chart 2=Series 3=Point
Public UndoAction(1 To 21) As Integer
Public UndoSel(1 To 21) As Integer
Public UndoChartNum(1 To 21) As Integer
Public UndoSeriesNum(1 To 21) As Integer
Public UndoPointNum(1 To 21) As Integer
Public UndoWb(1 To 21) As String
Public UndoWs(1 To 21) As String
Public CurUndoLevel As Integer
Public UndoPointValues() As Variant
Public UndoPointsInThisSeries(1 To 21) As Variant
Public SavedHiddenMarkers(20, 20, 40, 400) As Byte
Public PassAction As Integer
Public TheWarning As Boolean
Public UndoMaxPointsPerSeries As Integer
Public wbname As String
Public wsname As String
Public TotalConsecutiveUndos As Integer
Public HiddenColor As Double
Public AddSeriesUndo As Boolean
Public AddSeriesUndoNum As Integer
Public LastWasMarker As Boolean
Public bStop As Boolean

Sub ActivateForm()
Attribute ActivateForm.VB_ProcData.VB_Invoke_Func = "w\n14"
Call RecolorButtons
chartform.cb1Color.BackColor = RGB(255, 148, 143)
PassAction = 1
UndoMaxPointsPerSeries = 1
Set IDRange = Nothing
    CurChartNum = 0
    CurSeriesNum = 0
    CurPointNum = -1
wbname = ActiveWorkbook.Name
wsname = ActiveSheet.Name
TheColor(1) = RGB(192, 0, 0)
TheColor(2) = RGB(0, 0, 192)
CurColor = 1
CurUndoLevel = 1 'The only time it will have this value, it's between 2 and 21 but gets incremented right away
TheWarning = True
Call InstantiateConstants
ActionSelection = 1
chartform.Show vbModeless
chartform.frCol1.BackColor = TheColor(1)
chartform.frCol2.BackColor = TheColor(2)
If ActiveSheet.ChartObjects.Count > 0 Then
        ReDim clsEventCharts(1 To ActiveSheet.ChartObjects.Count)
        ReDim AllCharts(1 To ActiveSheet.ChartObjects.Count)
        Dim chtObj As ChartObject
        Dim chtnum As Integer

        chtnum = 1
        For Each chtObj In ActiveSheet.ChartObjects
            Set clsEventCharts(chtnum).EvtChart = chtObj.Chart
            chtnum = chtnum + 1
        Next ' chtObj
End If

End Sub
Sub UpdateStatus() '
Dim t As String
t = ""
If CurPointNum > 0 Then
    t = t & "Point " & CurPointNum & ": " 'Don't need other identifier from Chart itself
    If IDRange Is Nothing Then
        t = t & "[No ID Range Selected]"
        Else
        t = t & IDRange(CurPointNum).Value
    End If
End If
t = t & vbNewLine
On Error GoTo EH
If CurSeriesNum > 0 Then
    t = t & "Series " & CurSeriesNum & ": " & ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Name & vbCrLf & "[contains " & ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points.Count & " points]"
    If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).AxisGroup = 1 Then
        t = t & vbCrLf & "[Primary Axis]"
        Else
        t = t & vbCrLf & "[Secondary Axis]"
End If
End If


t = t & vbNewLine & "Chart " & CurChartNum
On Error Resume Next
t = t & ": " & Replace(ActiveSheet.ChartObjects(CurChartNum).Chart.ChartTitle.Text, vbNewLine, "") & "[contains " & ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection.Count & " series]"
On Error GoTo 0
t = t & vbCrLf & wsname & vbCrLf & wbname

chartform.CurActObj.Caption = t
chartform.Repaint
Exit Sub

EH:
If CurSeriesNum > 0 Then t = t & "Series " & CurSeriesNum
Resume Next
End Sub

Sub PerformAction()
Attribute PerformAction.VB_ProcData.VB_Invoke_Func = "q\n14"
If PassAction <> 6 Then LastWasMarker = False

Call RecolorButtons
Select Case PassAction 'Color the active button red
Case 1: chartform.cb1Color.BackColor = RGB(255, 148, 143)
Case 2: chartform.cb2ShowHide.BackColor = RGB(255, 148, 143)
Case 3: chartform.cb3Axis.BackColor = RGB(255, 148, 143)
Case 4: chartform.cb4LabelID.BackColor = RGB(255, 148, 143)
Case 5: chartform.cb5LabelCustom.BackColor = RGB(255, 148, 143)
Case 6: chartform.cb6Shape.BackColor = RGB(255, 148, 143)
Case 7: chartform.cb7Grow.BackColor = RGB(255, 148, 143)
'Case 8: chartform.cb8Shrink.BackColor = RGB(255, 148, 143)
Case 9: chartform.cb9Mark.BackColor = RGB(255, 148, 143)
End Select
If CurSel = 1 Then Exit Sub 'Cannot perform an action on an entire chart
Dim CurMarkVal As Integer
Dim TempBoolGlow As Boolean

' ### FOR DEBUGGING ###
'Dim ExampleChart As Variant 'For Debugging
'Set ExampleChart = ActiveSheet.ChartObjects(CurChartNum) 'For Debugging
'Dim ExamplePoint As Point 'For Debugging
'Set ExamplePoint = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(CurPointNum) 'For Debugging
'Dim tdb As Integer
'tdb = CurPointNum
'If tdb < 1 Then tdb = 1
'Debug.Print ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(tdb).DataLabel.Text
' ### END ###

Dim jjj As Integer
Dim Text5 As String
Dim Num7 As Integer

' PART ONE: First some checks to exit sub before setting the undo, or to add user input needed everywhere:
If PassAction = 5 Then Text5 = InputBox("Please enter custom label:")
On Error GoTo ES
If PassAction = 7 Then Num7 = CInt(InputBox("Please enter marker size (usually 5 to 10):"))
On Error GoTo 0
If CurSel = 2 And PassAction = 6 Then 'Check if all markers in series are the same shape; this is needed because it should increment to a uniform value
    Dim TempBoolMk As Boolean
    For jjj = 2 To ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points.Count
        If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle <> ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj - 1).MarkerStyle Then TempBoolMk = True
    Next jjj
    If TempBoolMk = True Then
        Dim TempMkMs As Integer
        TempMkMs = MsgBox("This procedures makes all marker shapes in this series identical; at present, they are not identical. Do you wish to continue?", vbYesNo)
        If TempMkMs = vbNo Then Exit Sub
    End If
End If
If CurSel = 2 And PassAction = 2 Then 'Check if any markers in the series are hidden, if so will default to show
    Dim TempBoolShow As Boolean
    For jjj = 1 To ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points.Count
        If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = xlMarkerStyleNone Then TempBoolShow = True
    Next jjj
End If
If CurSel = 2 And PassAction = 9 Then 'Check if there are any marked points in series at all; if so, mark will default to unmark
    For jjj = 1 To ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points.Count
        If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius > 0 Then TempBoolGlow = True
    Next jjj
End If
If PassAction = 4 And IDRange Is Nothing Then
    MsgBox "No ID Range selected."
    Exit Sub
End If
If CurSel = 3 And PassAction = 3 And TheWarning = True Then
        Dim Tempxcvb As Integer
        Tempxcvb = MsgBox("Single point selected, but this action will change entire series. Continue?", vbYesNo)
        If Tempxcvb = vbNo Then Exit Sub
End If


'PART TWO: Set Undo Parameters
If LastWasMarker = False And PassAction <> 3 Then 'LastWasMarker is used to keep each individual increment from taking up an undo level
TotalConsecutiveUndos = 20
IncRoll CurUndoLevel, 2, 21
UndoAction(CurUndoLevel) = PassAction
UndoSel(CurUndoLevel) = CurSel
UndoWb(CurUndoLevel) = wbname
UndoWs(CurUndoLevel) = wsname
UndoChartNum(CurUndoLevel) = CurChartNum
UndoSeriesNum(CurUndoLevel) = CurSeriesNum
UndoPointNum(CurUndoLevel) = CurPointNum
End If

'PART THREE:
UndoPointsInThisSeries(CurUndoLevel) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points.Count
If UndoPointsInThisSeries(CurUndoLevel) > UndoMaxPointsPerSeries Then
    UndoMaxPointsPerSeries = UndoPointsInThisSeries(CurUndoLevel)
    ReDim Preserve UndoPointValues(1 To 21, UndoMaxPointsPerSeries)
End If
For jjj = 1 To UndoPointsInThisSeries(CurUndoLevel)
  bStop = False
  If chartform.cbLimit.Value = True And ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius = 0 Then bStop = True
    Select Case PassAction
        Case 1 'Color
            UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerBackgroundColor
            If CurSel = 2 Or CurPointNum = jjj Then
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerBackgroundColor = TheColor(CurColor)
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerForegroundColor = TheColor(CurColor)
            End If
        Case 2 'Show/Hide
            UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle
            If CurSel = 2 Or CurPointNum = jjj Then
                If TempBoolShow = True Or ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = xlMarkerStyleNone Then
                    On Error GoTo TooMuchData
                    If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = MarkerArray(SavedHiddenMarkers(ActiveSheet.Index, CurChartNum, CurSeriesNum, jjj))
                    On Error GoTo 0
                    Else
                    On Error Resume Next
                    SavedHiddenMarkers(ActiveSheet.Index, CurChartNum, CurSeriesNum, jjj) = CByte(RetMarkNum(ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle))
                    On Error GoTo 0
                    If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = xlMarkerStyleNone
                End If
            End If
        'Case 3 disabled since primary/secondary axis n/a to individual points, it appears right after end of jjj loop below
         Case 4 'ID LAbel
            If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).HasDataLabel = False Then
                UndoPointValues(CurUndoLevel, jjj) = ""
                Else
                UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).DataLabel.Text
            End If
            If CurSel = 2 Or CurPointNum = jjj Then
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).HasDataLabel = True
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).DataLabel.Text = IDRange(jjj).Value
            End If
        Case 5 'Custom Label
            If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).HasDataLabel = False Then
                UndoPointValues(CurUndoLevel, jjj) = ""
                Else
                UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).DataLabel.Text
            End If
            If CurSel = 2 Or CurPointNum = jjj Then
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).HasDataLabel = True
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).DataLabel.Text = Text5
            End If
        Case 6 'Shape
            'LastWasMarker is used to keep each individual increment from taking up an undo level
            If LastWasMarker = False Then UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle
            If CurSel = 2 Or CurPointNum = jjj Then
                CurMarkVal = RetMarkNum(ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle)
                CurMarkVal = CurMarkVal + 1
                If CurMarkVal = 10 Then CurMarkVal = 1
                If bStop = False Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = MarkerArray(CurMarkVal)
            End If
            LastWasMarker = True
        Case 7 'Size
           UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerSize
           If bStop = False Then
               If CurSel = 2 Or CurPointNum = jjj Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerSize = Num7
           End If
        'Case 8 is deprecated
        Case 9 'Mark Glow - obviously bStop is irrelevant here
           UndoPointValues(CurUndoLevel, jjj) = ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius
           If CurSel = 2 Or CurPointNum = jjj Then
                If TempBoolGlow = True Or ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius = 15 Then
                    ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius = 0
                    Else
                    ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Color.RGB = RGB(255, 0, 0)
                    ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Transparency = 0.45
                    ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).Format.Glow.Radius = 15
                End If
            End If
    End Select
Next

If PassAction = 3 Then
    If ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).AxisGroup = 2 Then ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).AxisGroup = 1 Else ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).AxisGroup = 2
End If

Exit Sub
ES:
MsgBox "Invalid marker size."
Exit Sub
TooMuchData:
ActiveSheet.ChartObjects(CurChartNum).Chart.SeriesCollection(CurSeriesNum).Points(jjj).MarkerStyle = 8
End Sub

Sub TheUndo()
If AddSeriesUndo = True Then
ActiveChart.SeriesCollection(AddSeriesUndoNum).Delete
AddSeriesUndo = False
Else
AddSeriesUndo = False
If TotalConsecutiveUndos > 1 Then
Dim iii As Integer
For iii = 1 To UndoPointsInThisSeries(CurUndoLevel)
If UndoSel(CurUndoLevel) = 2 Or UndoPointNum(CurUndoLevel) = iii Then
Select Case UndoAction(CurUndoLevel)
Case 1
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).MarkerBackgroundColor = UndoPointValues(CurUndoLevel, iii)
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).MarkerForegroundColor = UndoPointValues(CurUndoLevel, iii)
Case 2, 6
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).MarkerStyle = UndoPointValues(CurUndoLevel, iii)
Case 4, 5
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).DataLabel.Text = UndoPointValues(CurUndoLevel, iii)
Case 7
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).MarkerSize = UndoPointValues(CurUndoLevel, iii)
Case 9
    Workbooks(UndoWb(CurUndoLevel)).Sheets(UndoWs(CurUndoLevel)).ChartObjects(UndoChartNum(CurUndoLevel)).Chart.SeriesCollection(UndoSeriesNum(CurUndoLevel)).Points(iii).Format.Glow.Radius = UndoPointValues(CurUndoLevel, iii)
End Select
End If
Next iii

CurUndoLevel = CurUndoLevel - 1
If CurUndoLevel = 1 Then CurUndoLevel = 21
TotalConsecutiveUndos = TotalConsecutiveUndos - 1
End If
End If 'For AddSeriesUndo
End Sub

Sub RecolorButtons()
'Colors them orange
chartform.cb1Color.BackColor = RGB(255, 208, 143)
chartform.cb2ShowHide.BackColor = RGB(255, 208, 143)
chartform.cb3Axis.BackColor = RGB(255, 208, 143)
chartform.cb4LabelID.BackColor = RGB(255, 208, 143)
chartform.cb5LabelCustom.BackColor = RGB(255, 208, 143)
chartform.cb6Shape.BackColor = RGB(255, 208, 143)
chartform.cb7Grow.BackColor = RGB(255, 208, 143)
'chartform.cb8Shrink.BackColor = RGB(255, 208, 143)
chartform.cb9Mark.BackColor = RGB(255, 208, 143)
End Sub

Function RetMarkNum(i5 As Integer) As Integer
Select Case i5
Case 1: RetMarkNum = 1
Case 2: RetMarkNum = 2
Case 3: RetMarkNum = 3
Case -4168: RetMarkNum = 4
Case 5: RetMarkNum = 5
Case -4118: RetMarkNum = 6
Case -4115: RetMarkNum = 7
Case 3: RetMarkNum = 8
Case 3: RetMarkNum = 9
Case Else: RetMarkNum = 1
End Select
End Function
