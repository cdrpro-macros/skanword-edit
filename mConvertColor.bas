Attribute VB_Name = "mConvertColor"
Public cRGB As New Color
Public Const appName$ = "SkanWordEdit"
Public Const appOpt$ = "Settings"

' Conver Colors:

Public Function EncodeColor(c As Color) As String
        Dim md As Long, c1 As Long, c2 As Long, c3 As Long
        Dim c4 As Long, c5 As Long, c6 As Long, c7 As Long
        c.CorelScriptGetComponent md, c1, c2, c3, c4, c5, c6, c7
        EncodeColor = DecToHex(md) & DecToHex(c1) & DecToHex(c2) & DecToHex(c3) & DecToHex(c4) & DecToHex(c5) & DecToHex(c6) & DecToHex(c7)
End Function

Public Function DecodeColor(c As Color, s As String)
        Dim md As Long, c1 As Long, c2 As Long, c3 As Long
        Dim c4 As Long, c5 As Long, c6 As Long, c7 As Long
        md = HexToDec(Mid$(s, 1, 4))
        If Application.VersionMajor < 11 And (md Mod 1000) > 30 Then
           c.CMYKAssign 0, 0, 0, 100
        Else
           c1 = HexToDec(Mid$(s, 5, 4))
           c2 = HexToDec(Mid$(s, 9, 4))
           c3 = HexToDec(Mid$(s, 13, 4))
           c4 = HexToDec(Mid$(s, 17, 4))
           c5 = HexToDec(Mid$(s, 21, 4))
           c6 = HexToDec(Mid$(s, 25, 4))
           c7 = HexToDec(Mid$(s, 29, 4))
           c.CorelScriptAssign md, c1, c2, c3, c4, c5, c6, c7
        End If
End Function

Private Function DecToHex(v As Long, Optional Length As Long = 4) As String
        DecToHex = Right$(String(Length, "0") & Hex$(v), Length)
End Function

Private Function HexToDec(s As String) As Long
        Dim n As Long
        n = Val("&h" & s)
        If n < 0 Then n = n + 65536
        HexToDec = n
End Function
