Attribute VB_Name = "DTHandy"
' #################################
' #                               #
' # DAVID TAYLOR'S HANDY VBA CODE #
' #                               #
' #################################

Sub Toggle(ByRef TheInput As Integer, Val1 As Integer, Val2 As Integer)
    If TheInput = Val1 Then
        TheInput = Val2
    ElseIf TheInput = Val2 Then
        TheInput = Val1
    End If
End Sub

Sub Increment(ByRef TheInput As Integer)
    TheInput = TheInput + 1
End Sub

Sub Decrement(ByRef TheInput As Integer)
    TheInput = TheInput + 1
End Sub

Sub IncRoll(ByRef TheInput As Integer, MinVal As Integer, MaxVal As Integer)
    TheInput = TheInput + 1
    If TheInput > MaxVal Then TheInput = MinVal
End Sub

Sub DecRoll(TheInput As Integer, MinVal As Integer, MaxVal As Integer)
    TheInput = TheInput - 1
    If TheInput < MinVal Then TheInput = MaxVal
End Sub
