VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkanWord_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub startEdit()
    If ActiveDocument Is Nothing Then Beep: Exit Sub
    myForm.Show 0
End Sub


