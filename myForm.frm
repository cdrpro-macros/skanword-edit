VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} myForm 
   Caption         =   "SkanWord Edit 2.2"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   OleObjectBlob   =   "myForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cfg As New clsConfig

Private Sub UserForm_Initialize()
    CopLabel.Caption = "Copyright " & Chr(169) & _
    " 2010" & Chr(10) & "Sancho," & Chr(10) & "www.cdrpro.ru"
    fontName.Text = GetSetting(appName, appOpt, "Font Name", "ER_helv")
    fontSize.Value = GetSetting(appName, appOpt, "Font Size", "6,5")
    fontLine.Value = GetSetting(appName, appOpt, "Font Line", "100")
    oWidth.Value = GetSetting(appName, appOpt, "Outline Width", "2")
    cfg.Load
    UpdateColor1 cfg.clrFontFill
    UpdateColor2 cfg.clrObjFill
    UpdateColor3 cfg.clrObjOutl
    myLoadPresets
End Sub

Private Sub fontName_Change(): SaveSetting appName, appOpt, "Font Name", fontName.Text: End Sub
Private Sub fontSize_Change(): SaveSetting appName, appOpt, "Font Size", fontSize.Value: End Sub
Private Sub fontLine_Change(): SaveSetting appName, appOpt, "Font Line", fontLine.Value: End Sub
'Private Sub oRoundness_Change(): SaveSetting appName, appOpt, "Roundness", oRoundness.Value: End Sub
Private Sub oWidth_Change(): SaveSetting appName, appOpt, "Outline Width", oWidth.Value: End Sub

Private Sub cmTextColor_Click()
    If cfg.clrFontFill.UserAssignEx() = False Then Exit Sub
    UpdateColor1 cfg.clrFontFill
End Sub

Private Sub UpdateColor1(c As Color)
    myConvertToRGB c
    cmTextColor.BackColor = RGB(cRGB.RGBRed, cRGB.RGBGreen, cRGB.RGBBlue)
    SaveSetting appName, appOpt, "Font Color", EncodeColor(c)
End Sub

Private Sub cm_fillColor_Click()
    If cfg.clrObjFill.UserAssignEx() = False Then Exit Sub
    UpdateColor2 cfg.clrObjFill
End Sub

Private Sub UpdateColor2(c As Color)
    myConvertToRGB c
    cm_fillColor.BackColor = RGB(cRGB.RGBRed, cRGB.RGBGreen, cRGB.RGBBlue)
    SaveSetting appName, appOpt, "Fill Color", EncodeColor(c)
End Sub

Private Sub cm_oColor_Click()
    If cfg.clrObjOutl.UserAssignEx() = False Then Exit Sub
    UpdateColor3 cfg.clrObjOutl
End Sub

Private Sub UpdateColor3(c As Color)
    myConvertToRGB c
    cm_oColor.BackColor = RGB(cRGB.RGBRed, cRGB.RGBGreen, cRGB.RGBBlue)
    SaveSetting appName, appOpt, "Outline Color", EncodeColor(c)
End Sub

Private Sub cm_Start_Click()
    ActiveDocument.Unit = ActiveDocument.Rulers.HUnits
    ActiveDocument.ReferencePoint = cdrCenter

    ActiveDocument.BeginCommandGroup "SkanWorkEdit"
    Optimization = True

    Dim st As Shape, p As Page

    For Each p In ActiveDocument.Pages
      p.Activate
      For Each st In p.FindShapes(, cdrTextShape)
          If st.ParentGroup Is Nothing Then
              If st.Text.Type = cdrArtisticText Or st.Text.Type = cdrArtisticFittedText Then Call DoEdit(st)
          End If
      Next
    Next

    Optimization = False
    ActiveDocument.EndCommandGroup
    Application.CorelScript.RedrawScreen

    MsgBox "Готово!", vbInformation, appName
End Sub

Private Sub DoEdit(st As Shape)
    If cb_text.Value = True Then
        st.Fill.UniformColor.CopyAssign cfg.clrFontFill
        With st.Text.Story
            .Font = fontName.Text
            .Size = CSng(Replace(fontSize.Value, ".", ","))
            .LineSpacing = CSng(Replace(fontLine.Value, ".", ","))
        End With
        'new 1 =============================
        If cb_Perenos.Value = True Then
          st.Text.Story.Text = Replace(st.Text.Story.Text, "*", "-" & vbCrLf, , , vbTextCompare)
        End If
        'new 1 =============================
    End If

    Dim s As Shape
    Set s = ActivePage.SelectShapesAtPoint(st.PositionX, st.PositionY, False)
    Dim sf As Shape

    For Each sf In s.Shapes
        If sf.Type = cdrRectangleShape Then
            If cb_block.Value = True Then
                sf.Fill.UniformColor.CopyAssign cfg.clrObjFill
                sf.Outline.Width = CDbl(Replace(oWidth.Text, ".", ","))
                sf.Outline.Color.CopyAssign cfg.clrObjOutl
            End If

            sf.Rectangle.CornerLowerLeft = CLng(cLL.Value)
            sf.Rectangle.CornerLowerRight = CLng(cLR.Value)
            sf.Rectangle.CornerUpperLeft = CLng(cUL.Value)
            sf.Rectangle.CornerUpperRight = CLng(cUR.Value)

            ActiveDocument.ClearSelection
            st.CreateSelection: sf.AddToSelection

            'Function AlignObjects(HorizontalAlignment As Long, VerticalAlignment As Long) As Long
            Dim h&, v&
            If aTL.Value Then
                h = 2: v = 1
            ElseIf aCL.Value Then
                h = 2: v = 3
            ElseIf aBL.Value Then
                h = 2: v = 2
            ElseIf aTC.Value Then
                h = 3: v = 1
            ElseIf aCC.Value Then
                h = 3: v = 3
            ElseIf aBC.Value Then
                h = 3: v = 2
            ElseIf aTR.Value Then
                h = 1: v = 1
            ElseIf aCR.Value Then
                h = 1: v = 3
            ElseIf aBR.Value Then
                h = 1: v = 2
            End If
            CorelScript.AlignObjects h, v    '1 - right/top   2 - left/bt    3 - center
            'st.AlignToShape cdrAlignHCenter + cdrAlignVCenter, sf
            Call sf.OrderToFront
            Call st.OrderToFront
        End If
    Next
End Sub

' Convert color to RGB for display on the form
Private Sub myConvertToRGB(c As Color)
        cRGB.CopyAssign c
        cRGB.ConvertToRGB
End Sub
 
' Load Preset
Private Sub myLoadPresets()
        Dim c&, i&, presName$
        c = GetSetting(appName, appOpt, "PresetsCount", 0)
        For i = 1 To c
            presName = GetSetting(appName, appOpt, "Presets" & i & "Name")
            If presName <> "" Then cb_presList.AddItem i & "| " & presName
        Next i
End Sub

' Change Preset
Private Sub cb_presList_Change()
    Dim i&, a$(), c1&, a2$()
    If cb_presList.SelLength = 0 Then Exit Sub

    For c1 = 0 To cb_presList.ListCount - 1 Step 1
        If cb_presList.SelText = cb_presList.List(c1) Then
            a2 = Split(cb_presList.SelText, "|")
            i = CLng(a2(0))
            Exit For
        End If
    Next c1

    a = Split(GetSetting(appName, appOpt, "Presets" & i), "|")
    Dim sAlign$
    sAlign = a(0)

    Select Case sAlign
        Case "TL": aTL.Value = True
        Case "CL": aCL.Value = True
        Case "BL": aBL.Value = True
        Case "TC": aTC.Value = True
        Case "CC": aCC.Value = True
        Case "BC": aBC.Value = True
        Case "TR": aTR.Value = True
        Case "CR": aCR.Value = True
        Case "BR": aBR.Value = True
    End Select

    fontName.Text = a(1)
    fontSize.Text = a(2)
    DecodeColor cfg.clrFontFill, a(3)
    UpdateColor1 cfg.clrFontFill
    DecodeColor cfg.clrObjFill, a(4)
    UpdateColor2 cfg.clrObjFill
    DecodeColor cfg.clrObjOutl, a(5)
    UpdateColor3 cfg.clrObjOutl
    oWidth.Text = a(6)
    cLL.Value = a(7)
    cLR.Value = a(8)
    cUL.Value = a(9)
    cUR.Value = a(10)

    If UBound(a) = 13 Then
      cb_text.Value = a(11)
      cb_block.Value = a(12)
      fontLine.Text = a(13)
    End If

    'new 1 =============================
    If UBound(a) = 14 Then
      cb_text.Value = a(11)
      cb_block.Value = a(12)
      fontLine.Text = a(13)
      cb_Perenos.Value = a(14)
    End If
End Sub

' Add Preset
Private Sub cm_presAdd_Click()
    Dim strPres$, strPresN$, c&
    c = GetSetting(appName, appOpt, "PresetsCount", 0)
    c = c + 1

    strPresN = InputBox("Name for Preset", "Name...")
    If strPresN = "" Then Exit Sub

    Dim sAlign$
    If aTL.Value Then
        sAlign = "TL"
    ElseIf aCL.Value Then
        sAlign = "CL"
    ElseIf aBL.Value Then
        sAlign = "BL"
    ElseIf aTC.Value Then
        sAlign = "TC"
    ElseIf aCC.Value Then
        sAlign = "CC"
    ElseIf aBC.Value Then
        sAlign = "BC"
    ElseIf aTR.Value Then
        sAlign = "TR"
    ElseIf aCR.Value Then
        sAlign = "CR"
    ElseIf aBR.Value Then
        sAlign = "BR"
    End If

    strPres = sAlign & "|" & fontName.Text & "|" & fontSize.Text & "|" & EncodeColor(cfg.clrFontFill) & "|" & _
    EncodeColor(cfg.clrObjFill) & "|" & EncodeColor(cfg.clrObjOutl) & "|" & oWidth.Text & "|" & _
    cLL.Value & "|" & cLR.Value & "|" & cUL.Value & "|" & cUR.Value & "|" & _
    cb_text.Value & "|" & cb_block.Value & "|" & fontLine.Text & "|" & cb_Perenos.Value
    
    SaveSetting appName, appOpt, "Presets" & c, strPres
    SaveSetting appName, appOpt, "Presets" & c & "Name", strPresN

    cb_presList.AddItem c & "| " & strPresN
    cb_presList.Text = c & "| " & strPresN
    SaveSetting appName, appOpt, "PresetsCount", c
End Sub

' Delete Preset
Private Sub cm_presDel_Click()
        Dim i&, c&, i2&, a$(), c1&

        If cb_presList.SelLength = 0 Then Exit Sub
        For c1 = 0 To cb_presList.ListCount - 1 Step 1
            If cb_presList.SelText = cb_presList.List(c1) Then
                a = Split(cb_presList.SelText, "|")
                i = CLng(a(0))
                Exit For
            End If
        Next c1

        cb_presList.Clear
        c = CLng(GetSetting(appName, appOpt, "PresetsCount", 0))

        If i < c Then
            For i2 = i + 1 To c Step 1
                SaveSetting appName, appOpt, "Presets" & i, _
                GetSetting(appName, appOpt, "Presets" & i2)
                SaveSetting appName, appOpt, "Presets" & i & "Name", _
                GetSetting(appName, appOpt, "Presets" & i2 & "Name")
                i = i + 1
            Next i2
            DeleteSetting appName, appOpt, "Presets" & c
            DeleteSetting appName, appOpt, "Presets" & c & "Name"
        Else
            DeleteSetting appName, appOpt, "Presets" & i
            DeleteSetting appName, appOpt, "Presets" & i & "Name"
        End If

        SaveSetting appName, appOpt, "PresetsCount", c - 1
        myLoadPresets
End Sub


