Attribute VB_Name = "modTextBoxFormatter"
Option Explicit

Private Const SETTINGS_APP As String = "CodexExcelTextBoxFormatter"
Private Const SETTINGS_SECTION As String = "Format"
Private Const DEFAULT_FONT_SIZE As Double = 14
Private Const DEFAULT_FONT_HEX As String = "#0033CC"
Private Const DEFAULT_FILL_HEX As String = "#FFFFFF"
Private Const DEFAULT_BOLD As Boolean = True

Public Sub FormatSelectedTextBoxes()
    Dim selectedShapes As ShapeRange
    Dim currentShape As Shape
    Dim fontColor As Long
    Dim fillColor As Long
    Dim fontSize As Double
    Dim useBold As Boolean

    On Error GoTo SafeExit
    Set selectedShapes = Selection.ShapeRange

    fontSize = GetConfiguredFontSize()
    fontColor = ColorFromHexOrDefault(GetConfiguredFontHex(), DEFAULT_FONT_HEX)
    fillColor = ColorFromHexOrDefault(GetConfiguredFillHex(), DEFAULT_FILL_HEX)
    useBold = GetConfiguredBold()

    For Each currentShape In selectedShapes
        ApplyFormatToShape currentShape, fontSize, fontColor, fillColor, useBold
    Next currentShape

SafeExit:
    ' Deliberately silent: unsupported selections must never interrupt work.
End Sub

Public Sub OpenTextBoxFormatSettings()
    On Error GoTo SafeExit
    frmTextBoxSettings.Show
SafeExit:
End Sub

Public Sub FormatSelectedTextBoxesFromRibbon(ByVal control As IRibbonControl)
    FormatSelectedTextBoxes
End Sub

Public Sub OpenTextBoxFormatSettingsFromRibbon(ByVal control As IRibbonControl)
    OpenTextBoxFormatSettings
End Sub

Private Sub ApplyFormatToShape(ByVal targetShape As Shape, _
                               ByVal fontSize As Double, _
                               ByVal fontColor As Long, _
                               ByVal fillColor As Long, _
                               ByVal useBold As Boolean)
    Dim itemIndex As Long

    On Error GoTo SkipShape

    If targetShape.Type = msoGroup Then
        For itemIndex = 1 To targetShape.GroupItems.Count
            ApplyFormatToShape targetShape.GroupItems.Item(itemIndex), fontSize, fontColor, fillColor, useBold
        Next itemIndex
        Exit Sub
    End If

    If targetShape.Type <> msoTextBox Then Exit Sub

    With targetShape.TextFrame2.TextRange.Font
        .Size = fontSize
        If useBold Then
            .Bold = msoTrue
        Else
            .Bold = msoFalse
        End If
        With .Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = fontColor
            .Transparency = 0
        End With
    End With

    With targetShape.TextFrame.Characters.Font
        .Size = fontSize
        .Bold = useBold
        .Color = fontColor
    End With

    With targetShape.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
        .Transparency = 0
    End With

SkipShape:
End Sub

Public Function GetConfiguredFontSize() As Double
    Dim storedValue As String
    storedValue = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "FontSize", CStr(DEFAULT_FONT_SIZE))

    If IsNumeric(storedValue) Then
        GetConfiguredFontSize = CDbl(storedValue)
        If GetConfiguredFontSize >= 1 And GetConfiguredFontSize <= 409 Then Exit Function
    End If

    GetConfiguredFontSize = DEFAULT_FONT_SIZE
End Function

Public Function GetConfiguredFontHex() As String
    Dim storedValue As String
    Dim parsedColor As Long
    storedValue = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "FontColor", DEFAULT_FONT_HEX)

    If TryParseHexColor(storedValue, parsedColor) Then
        GetConfiguredFontHex = NormalizeHex(storedValue)
    Else
        GetConfiguredFontHex = DEFAULT_FONT_HEX
    End If
End Function

Public Function GetConfiguredFillHex() As String
    Dim storedValue As String
    Dim parsedColor As Long
    storedValue = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "FillColor", DEFAULT_FILL_HEX)

    If TryParseHexColor(storedValue, parsedColor) Then
        GetConfiguredFillHex = NormalizeHex(storedValue)
    Else
        GetConfiguredFillHex = DEFAULT_FILL_HEX
    End If
End Function

Public Function GetConfiguredBold() As Boolean
    Dim storedValue As String
    storedValue = LCase$(Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Bold", CStr(DEFAULT_BOLD))))
    GetConfiguredBold = (storedValue = "true" Or storedValue = "1" Or storedValue = "yes")
End Function

Public Function IsValidFontSize(ByVal value As Variant) As Boolean
    If Not IsNumeric(value) Then Exit Function
    IsValidFontSize = (CDbl(value) >= 1 And CDbl(value) <= 409)
End Function

Public Function TryParseHexColor(ByVal value As String, ByRef parsedColor As Long) As Boolean
    Dim normalized As String
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    On Error GoTo InvalidColor
    normalized = NormalizeHex(value)
    If Len(normalized) <> 7 Then GoTo InvalidColor

    redValue = CLng("&H" & Mid$(normalized, 2, 2))
    greenValue = CLng("&H" & Mid$(normalized, 4, 2))
    blueValue = CLng("&H" & Mid$(normalized, 6, 2))
    parsedColor = RGB(redValue, greenValue, blueValue)
    TryParseHexColor = True
    Exit Function

InvalidColor:
    TryParseHexColor = False
End Function

Public Function NormalizeHex(ByVal value As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(value))
    If Left$(normalized, 1) <> "#" Then normalized = "#" & normalized
    NormalizeHex = normalized
End Function

Public Sub SaveFormatterSettings(ByVal fontSize As Double, _
                                 ByVal fontHex As String, _
                                 ByVal fillHex As String, _
                                 ByVal useBold As Boolean)
    Dim parsedColor As Long

    If Not IsValidFontSize(fontSize) Then Exit Sub
    If Not TryParseHexColor(fontHex, parsedColor) Then Exit Sub
    If Not TryParseHexColor(fillHex, parsedColor) Then Exit Sub

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "FontSize", CStr(fontSize)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "FontColor", NormalizeHex(fontHex)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "FillColor", NormalizeHex(fillHex)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Bold", CStr(useBold)
End Sub

Public Sub ResetFormatterSettings()
    SaveFormatterSettings DEFAULT_FONT_SIZE, DEFAULT_FONT_HEX, DEFAULT_FILL_HEX, DEFAULT_BOLD
End Sub

Public Function SelfTestSettingsForm() As Boolean
    On Error GoTo Failed
    Load frmTextBoxSettings
    If CStr(frmTextBoxSettings.Controls("txtFontSize").Value) = vbNullString Then GoTo Failed
    If CStr(frmTextBoxSettings.Controls("txtFontColor").Value) = vbNullString Then GoTo Failed
    If CStr(frmTextBoxSettings.Controls("txtFillColor").Value) = vbNullString Then GoTo Failed
    Unload frmTextBoxSettings
    SelfTestSettingsForm = True
    Exit Function

Failed:
    On Error Resume Next
    Unload frmTextBoxSettings
    SelfTestSettingsForm = False
End Function

Public Function DefaultFontSize() As Double
    DefaultFontSize = DEFAULT_FONT_SIZE
End Function

Public Function DefaultFontHex() As String
    DefaultFontHex = DEFAULT_FONT_HEX
End Function

Public Function DefaultFillHex() As String
    DefaultFillHex = DEFAULT_FILL_HEX
End Function

Public Function DefaultBold() As Boolean
    DefaultBold = DEFAULT_BOLD
End Function

Private Function ColorFromHexOrDefault(ByVal value As String, ByVal defaultValue As String) As Long
    If TryParseHexColor(value, ColorFromHexOrDefault) Then Exit Function
    Call TryParseHexColor(defaultValue, ColorFromHexOrDefault)
End Function
