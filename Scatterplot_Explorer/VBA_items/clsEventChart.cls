VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Declare object of type "Chart" with events
Public WithEvents EvtChart As Chart
Attribute EvtChart.VB_VarHelpID = -1

Private Sub EvtChart_Activate()
EvtChart.ChartObjects msoBringToFront
'CurChartNum = EvtChart.Parent.Index 'THE REST OF IT SLOWS EVERYTHING DOWN!
'CurSel = 1
'Call UpdateStatus
End Sub

Private Sub EvtChart_Select(ByVal ElementId As Long, ByVal arg1 As Long, ByVal arg2 As Long)
LastWasMarker = False 'New selection means new undo level
CurChartNum = ActiveChart.Parent.Index
If chartform.cbFreeze = False Then
wbname = ActiveWorkbook.Name
wsname = ActiveSheet.Name
'CurChartNum = 0
'If ElementID > 0 Then CurChartNum = EvtChart.Parent.Index
Application.StatusBar = CurChartNum
CurSeriesNum = arg1
CurPointNum = arg2
CurSel = 1
If arg1 > 0 Then CurSel = 2
If arg2 > 0 Then CurSel = 3
 'Application.StatusBar = ElementID & Arg1 & Arg2
 Call UpdateStatus
 End If
 If arg2 > 0 And chartform.cbSelect.Value = True Then Call PerformAction
End Sub

Private Sub EvtChart_MouseMove(ByVal Button As Long, ByVal Shift As Long, _
ByVal x As Long, ByVal y As Long)
Dim ElementId As Long
Dim arg1 As Long, arg2 As Long
On Error Resume Next
EvtChart.GetChartElement x, y, ElementId, arg1, arg2

Dim t As String
t = ""
If arg2 > 0 Then
    If IDRange Is Nothing Then
        t = t & "[NoID]"
        Else
        t = t & IDRange(arg2).Value
    End If
t = t & " : pt." & arg2
End If
If arg1 > 0 Then
    t = t & " S" & arg1 & ":" & EvtChart.ChartObjects(arg1).Chart.SeriesCollection(arg1).Name
    'If EvtChart.ChartObjects(arg1).Chart.SeriesCollection(arg1).AxisGroup = 2 Then t = t & "[2°]"
End If

Application.StatusBar = t
End Sub
