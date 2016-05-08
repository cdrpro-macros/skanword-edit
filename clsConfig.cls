VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public clrFontFill As New Color
Public clrObjFill As New Color
Public clrObjOutl As New Color

Private Sub Class_Initialize()
    clrFontFill.CMYKAssign 0, 0, 0, 100
    clrObjFill.CMYKAssign 0, 0, 0, 100
    clrObjOutl.CMYKAssign 0, 0, 0, 100
End Sub

Public Sub Load()
    Dim s$, nScope&
    Const sDefColor$ = "138A0000000000000064000000000000"
    s = GetSetting(appName, appOpt, "Font Color", sDefColor)
    If Len(s) = Len(sDefColor) Then DecodeColor clrFontFill, s
    s = GetSetting(appName, appOpt, "Fill Color", sDefColor)
    If Len(s) = Len(sDefColor) Then DecodeColor clrObjFill, s
    s = GetSetting(appName, appOpt, "Outline Color", sDefColor)
    If Len(s) = Len(sDefColor) Then DecodeColor clrObjOutl, s
End Sub
