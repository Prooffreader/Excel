Attribute VB_Name = "colormacros"
Option Explicit

'================================================================================
'
' Attribution to software creator must remain in situ:
'
' Author:                   Mark Kubiszyn
' Website:                  http://www.kubiszyn.co.uk
' Email, Comments / Bugs:   mark@kubiszyn.co.uk
'
'================================================================================

' used for the Colour Picker Dialog
Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
Private Declare PtrSafe Function ChooseColorDlg Lib "comdlg32.dll" Alias "ChooseColorA" ( _
    pChoosecolor As CHOOSECOLOR) As LongPtr

Private Type CHOOSECOLOR
    lStructSize As LongPtr
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As LongPtr
    lpCustColors As LongPtr
    flags As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&

Private dwCustClrs(0 To 15) As LongPtr

' # ADJUST this for an alternative Cell
Private Const CellToColour = "B11"


'================================================================================
' Colour Picker, Subroutine
'================================================================================
Public Sub PickColour()
 Dim c As LongPtr
 c = ChooseColorDialog(1)
 If c <> -1 Then
  Sheet1.Range(CellToColour).Interior.Color = c
  Sheet1.Range(CellToColour).Value = CStr(HexRGB(c))
 End If
End Sub




'================================================================================
' UserForm Colour Picker, Subroutine
'================================================================================
Public Sub UserFormColourPicker()
 UserForm.Show
End Sub




'================================================================================
' ChooseColorDialog, Colour Dialog Picker Function
'================================================================================
Public Function ChooseColorDialog(DefaultColor As LongPtr) As LongPtr
 Dim lpChoosecolor As CHOOSECOLOR
  With lpChoosecolor
   .lStructSize = Len(lpChoosecolor)
   .hwndOwner = GetActiveWindow
   .rgbResult = DefaultColor
   .lpCustColors = VarPtr(dwCustClrs(0))
   .flags = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN
  End With
  If ChooseColorDlg(lpChoosecolor) Then
   ChooseColorDialog = lpChoosecolor.rgbResult
  Else
   ChooseColorDialog = -1
  End If
End Function




'================================================================================
' HexRGB, Function
'================================================================================
Public Function HexRGB(ByVal lCdlColor As LongPtr) As String
 Dim lCol As LongPtr
 Dim iRed, iGreen, iBlue As Integer
 Dim vHexR, vHexG, vHexB As Variant
 lCol = lCdlColor
 iRed = lCol Mod 4096
 lCol = lCol \ 4096
 iGreen = lCol Mod 4096
 lCol = lCol \ 4096
 iBlue = CInt(lCol / 4096)
 'Determine Red Hex
 vHexR = Hex(iRed)
 If Len(vHexR) < 2 Then
  vHexR = "0" & vHexR
 End If
 'Determine Green Hex
 vHexG = Hex(iGreen)
 If Len(vHexG) < 2 Then
  vHexG = "0" & iGreen
 End If
 'Determine Blue Hex
 vHexB = Hex(iBlue)
 If Len(vHexB) < 2 Then
  vHexB = "0" & vHexB
 End If
 'Add it up, return the function value
 HexRGB = vHexR & vHexG & vHexB
End Function



