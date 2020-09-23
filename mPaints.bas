Attribute VB_Name = "mPaints"
Option Explicit

Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private RGBs() As String
Private Const HSLMAX As Integer = 240
Private Const RGBMAX As Integer = 255
Private Const UNDEFINED As Integer = (HSLMAX * 2 \ 3)

Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)

Public Function TranslateColor(lColor As Long) As Long
   OleTranslateColor lColor, 0, TranslateColor
End Function

Public Function getDarkColor(ByVal Colour As Long, ByVal d As Byte) As Long
    Dim R As Byte, G As Byte, B As Byte
    Dim CR As Byte, CG As Byte, CB As Byte
    
    CR = (Colour And &HFF&)
    CG = (Colour And &HFF00&) / &H100
    CB = (Colour And &HFF0000) / &H10000
    
    If (CR > d) Then R = (CR - d)
    If (CG > d) Then G = (CG - d)
    If (CB > d) Then B = (CB - d)

    getDarkColor = RGB(R, G, B)
End Function

Public Function getLightColor(ByVal Colour As Long, ByVal d As Byte) As Long
    Dim R As Byte, G As Byte, B As Byte
    Dim CR As Byte, CG As Byte, CB As Byte
    R = 255
    G = 255
    B = 255 '
    
    CR = (Colour And &HFF&)
    CG = (Colour And &HFF00&) / &H100
    CB = (Colour And &HFF0000) / &H10000
    
    If (CInt(CR) + CInt(d) <= 255) Then R = (CR + d)
    If (CInt(CG) + CInt(d) <= 255) Then G = (CG + d)
    If (CInt(CB) + CInt(d) <= 255) Then B = (CB + d) '

    getLightColor = RGB(R, G, B)
End Function

Public Function BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
      
End Function

Public Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

    'Creates an array of colors blending from
    'Color1 to Color2 in lSteps number of steps.
    'Returns the count and fills the laRetColors() array.

    Dim lIdx    As Long
    Dim lRed    As Long
    Dim lGrn    As Long
    Dim lBlu    As Long
    Dim fRedStp As Single
    Dim fGrnStp As Single
    Dim fBluStp As Single

    'Stop possible error
    If lSteps < 2 Then lSteps = 2

    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000

    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))

    'Create the colors
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1                               'First Color
    laRetColors(lSteps - 1) = lColor2                      'Last Color
    For lIdx = 1 To lSteps - 2                             'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
                (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
                (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next lIdx

    'Return number of colors in array
    BlendColors = lSteps

End Function

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double

    'Divides dNumer by dDenom if dDenom <> 0
    'Eliminates 'Division By Zero' error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

' ======================================================================== '
Public Function LongToRGB(Value As Long, Optional ChooseDelimiter As String = ",") As String
    On Error Resume Next

    Dim Blue As Double, Green As Double, Red As Double
    Dim BlueS As Double, GreenS As Double, RGBs As String

    Blue = Abs(Fix((Value / 256) / 256))
    BlueS = (Blue * 256) * 256
    Green = Abs(Fix((Value - BlueS) / 256))
    GreenS = Green * 256
    Red = Abs(Fix(Value - BlueS - GreenS))
    RGBs = (Red & ChooseDelimiter & Green & ChooseDelimiter & Blue)
    LongToRGB = RGBs
End Function

' ======================================================================== '

Public Function RGBToHex2(R As Long, G As Long, B As Long) As String
    On Error Resume Next

    Dim sR As String, sG As String, sB As String

    sR = Hex$(R)
    sG = Hex$(G)
    sB = Hex$(B)

    If sR = "0" Then sR = "00"
    If sG = "0" Then sG = "00"
    If sB = "0" Then sB = "00"

    RGBToHex2 = sR & sG & sB
End Function

' ======================================================================== '

Public Function RGBToHex(RGBValue As String, Optional delimiter As String = ",") As String
    On Error Resume Next

    Dim PartHexValue(2) As String, FullHexValue As String, i As Long

    RGBs() = Split(RGBValue, delimiter)

    Do While i <= 2
        PartHexValue(i) = Hex$(Trim$(RGBs(i)))
        If PartHexValue(i) = "0" Then PartHexValue(i) = "00"
        If Len(PartHexValue(i)) <> 2 Then PartHexValue(i) = "0" & PartHexValue(i)
        FullHexValue = FullHexValue & PartHexValue(i)

        i = i + 1
    Loop

    RGBToHex = FullHexValue
End Function

' ======================================================================== '

Public Function HexToRGB(ByVal HexValue As String) As String
    On Error Resume Next

    If AscW(HexValue) = 35 Then HexValue = Right$(HexValue, Len(HexValue) - 1)
'    If Left$(HexValue, 1) = "#" Then
    Dim RGBValue(2) As Long

    RGBValue(0) = CLng((GetHexValue(Mid$(HexValue, 1, 1))) * 16 + (GetHexValue(Mid$(HexValue, 2, 1))))
    RGBValue(1) = CLng((GetHexValue(Mid$(HexValue, 3, 1))) * 16 + (GetHexValue(Mid$(HexValue, 4, 1))))
    RGBValue(2) = CLng((GetHexValue(Mid$(HexValue, 5, 1))) * 16 + (GetHexValue(Mid$(HexValue, 6, 1))))

    HexToRGB = "RGB(" & RGBValue(0) & ", " & RGBValue(1) & ", " & RGBValue(2) & ")"
End Function

' ======================================================================== '

Public Function ShortHexToLong(ByVal HexValue As String) As Long
    On Error Resume Next

    If AscW(HexValue) = 35 Then HexValue = Right$(HexValue, Len(HexValue) - 1)
'    If Left$(HexValue, 1) = "#" Then
    Dim RGBValue(2) As Long

    RGBValue(0) = CLng((GetHexValue(Mid$(HexValue, 1, 1))) * 16)    ' + (GetHexValue(Mid$(HexValue, 2, 1))))
    RGBValue(1) = CLng((GetHexValue(Mid$(HexValue, 2, 1))) * 16)    ' + (GetHexValue(Mid$(HexValue, 4, 1))))
    RGBValue(2) = CLng((GetHexValue(Mid$(HexValue, 3, 1))) * 16)    ' + (GetHexValue(Mid$(HexValue, 6, 1))))

    ShortHexToLong = RGB(RGBValue(0), RGBValue(1), RGBValue(2))
End Function

' ======================================================================== '

Public Function GetHexValue(HexChar As String) As String
    On Error Resume Next

    Dim a As Integer
    a = AscW(UCase$(HexChar))

    Select Case a
        Case 65 To 70       ' "A" To "F"
            GetHexValue = CStr(10 + (a - 65))
        Case 48 To 57       ' 0 To 9
            GetHexValue = HexChar
        Case Else
            GetHexValue = 0
    End Select
End Function

' ======================================================================== '

Public Function LongToHex(LongValue As Long, Optional ByVal bIncludeSign As Boolean = True) As String
    On Error Resume Next

    Dim Pad As Integer, ColorCode As String

    ColorCode = Hex$(LongValue)    ' convert to hex

    Pad = 6 - Len(ColorCode)    ' determine how many zeros to pad in front of converted value

    If Pad > 0 Then ColorCode = String$(Pad, "0") & ColorCode
    ColorCode = Right$(ColorCode, 2) & Mid$(ColorCode, 3, 2) & Left$(ColorCode, 2)      ' convert to hex

    If bIncludeSign Then LongToHex = "#" & ColorCode Else LongToHex = ColorCode
End Function

' ======================================================================== '

Public Function iMax(a As Integer, B As Integer) As Integer
    ' Return the Larger of two values
    If a > B Then iMax = a Else iMax = B
End Function

' ======================================================================== '

Public Function iMin(a As Integer, B As Integer) As Integer
    ' Return the smaller of two values
    If a < B Then iMin = a Else iMin = B
End Function

' ======================================================================== '

Public Function RgbToHSL(ByVal R As Long, ByVal G As Long, ByVal B As Long) As String
    On Error Resume Next

    Dim cMax As Integer, cMin As Integer
    Dim RDelta As Double, GDelta As Double, BDelta As Double
    Dim H As Double, S As Double, L As Double
    Dim cMinus As Long, cPlus As Long

    cMax = iMax(iMax(CInt(R), CInt(G)), CInt(B))   'Highest and lowest
    cMin = iMin(iMin(CInt(R), CInt(G)), CInt(B)) 'color values

    cMinus = cMax - cMin ' Used To simplify the
    cPlus = cMax + cMin ' calculations somewhat.

    ' Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)

    If cMax = cMin Then ' achromatic (r=g=b, greyscale)
        S = 0 ' Saturation 0 For greyscale
        H = UNDEFINED ' Hue undefined For greyscale
    Else
        ' Calculate color saturation
        If L <= (HSLMAX / 2) Then
            S = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If

        ' Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus


        Select Case cMax
            Case CLng(R)
                H = BDelta - GDelta
            Case CLng(G)
                H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select

        If H < 0 Then H = H + HSLMAX
    End If

    RgbToHSL = H & " " & S & " " & L
End Function

' ======================================================================== '

Public Function HSLtoRGB(ByRef H As Double, ByRef S As Double, ByRef L As Double) As String
    On Error Resume Next

    Dim R As Double, G As Double, B As Double
    Dim Magic1 As Double, Magic2 As Double

    If CInt(S) = 0 Then 'Greyscale
        R = (L * RGBMAX) / HSLMAX 'luminescence,
                'converted to the proper range
        G = R 'All RGB values same in greyscale
        B = R
'        If CInt(H) <> UNDEFINED Then
'            'This is technically an error.
'            'The RGBtoHSL routine will always return
'            'Hue = UNDEFINED (160 when HSLMAX is 240)
'            'when Sat = 0.
'            'if you are writing a color mixer and
'            'letting the user input color values,
'            'you may want to set Hue = UNDEFINED
'            'in this case.
'        End If
    Else
        'Get the "Magic Numbers"
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + S) + 0.5) / HSLMAX
        Else
            Magic2 = L + S - ((L * S) + 0.5) / HSLMAX
        End If

        Magic1 = 2 * L - Magic2

        ' Get R, G, B; change units from HSLMAX range
        ' to RGBMAX range
        R = (HueToRGB(Magic1, Magic2, H + (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
        G = (HueToRGB(Magic1, Magic2, H) * RGBMAX + 0.5) / HSLMAX
        B = (HueToRGB(Magic1, Magic2, H - (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
    End If

    H = R: S = G: L = B
    HSLtoRGB = "RGB(" & CInt(R) & ", " & CInt(G) & ", " & CInt(B) & ")"
End Function

' ======================================================================== '

Private Function HueToRGB(Mag1 As Double, Mag2 As Double, ByVal Hue As Double) As Double
    On Error Resume Next

    ' Utility function for HSLtoRGB

    ' Range check
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If

    'Return r, g, or b value from parameters
    Select Case Hue 'Values get progressively larger.
                'Only the first true condition will execute
        Case Is < (HSLMAX / 6)
            HueToRGB = (Mag1 + (((Mag2 - Mag1) * Hue + (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HueToRGB = Mag2
        Case Is < (HSLMAX * 2 / 3)
            HueToRGB = (Mag1 + (((Mag2 - Mag1) * ((HSLMAX * 2 / 3) - Hue) + (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HueToRGB = Mag1
    End Select
End Function

' ======================================================================== '

Public Function Brighten(ByRef R As Long, ByRef G As Long, ByRef B As Long, Optional Percent As Single = 0.07)
    On Error Resume Next

    ' Lightens the color by a specifie percent, given as a Single (10% = .10)
    Dim H As Double, S As Double, L As Double, HSL As String, sTmp() As String

    If Percent <= 0 Then Exit Function

    HSL = RgbToHSL(R, G, B)
    sTmp = Split(HSL)
    H = CDbl(sTmp(0))
    S = CDbl(sTmp(1))
    L = CDbl(sTmp(2)) + (HSLMAX * Percent)
    If L > HSLMAX Then L = HSLMAX
    Call HSLtoRGB(H, S, L)
    R = H: G = S: B = L
End Function

' ======================================================================== '

Public Function Darken(ByRef R As Long, ByRef G As Long, ByRef B As Long, Optional Percent As Single = 0.07)
    ' Darkens the color by a specifie percent, given as a Single
    On Error Resume Next
    Dim H As Double, S As Double, L As Double, HSL As String, sArray() As String

    If Percent <= 0 Then Exit Function

    HSL = RgbToHSL(R, G, B)
    sArray = Split(HSL)
    H = CDbl(sArray(0))
    S = CDbl(sArray(1))
    L = CDbl(sArray(2)) - (HSLMAX * Percent)
    If L < 0 Then L = 0
    Call HSLtoRGB(H, S, L)
    R = H: G = S: B = L
End Function

' ======================================================================== '

Public Function Darken2(ByRef Colour As Long, Optional Percent As Single = 0.07)
    ' Darkens the color by a specifie percent, given as a Single
    On Error Resume Next
    Dim H As Double, S As Double, L As Double, HSL As String, sArray() As String
    Dim R As Integer, G As Integer, B As Integer
    
    If Percent <= 0 Then Exit Function
    
    R = (Colour And &HFF&)
    G = (Colour And &HFF00&) / &H100
    B = (Colour And &HFF0000) / &H10000
    
    HSL = RgbToHSL(R, G, B)
    sArray = Split(HSL)
    H = CDbl(sArray(0))
    S = CDbl(sArray(1))
    L = CDbl(sArray(2)) - (HSLMAX * Percent)
    If L < 0 Then L = 0
    Call HSLtoRGB(H, S, L)
    Colour = RGB(CLng(H), CLng(S), CLng(L))
End Function

' ======================================================================== '

Public Function Brighten2(ByRef Colour As Long, Optional Percent As Single = 0.07)
    ' Darkens the color by a specifie percent, given as a Single
    On Error Resume Next
    Dim H As Double, S As Double, L As Double, HSL As String, sArray() As String
    Dim R As Integer, G As Integer, B As Integer
    
    If Percent <= 0 Then Exit Function
    
    R = (Colour And &HFF&)
    G = (Colour And &HFF00&) / &H100
    B = (Colour And &HFF0000) / &H10000
    
    HSL = RgbToHSL(R, G, B)
    sArray = Split(HSL)
    H = CDbl(sArray(0))
    S = CDbl(sArray(1))
    L = CDbl(sArray(2)) + (HSLMAX * Percent)
    If L < 0 Then L = 0
    Call HSLtoRGB(H, S, L)
    Colour = RGB(CLng(H), CLng(S), CLng(L))
End Function


