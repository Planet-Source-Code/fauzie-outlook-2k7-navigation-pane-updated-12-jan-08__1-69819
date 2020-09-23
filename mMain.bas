Attribute VB_Name = "mMain"
Option Explicit
DefInt A-Z

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Private Const HWND_TOP& = 0
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOACTIVATE& = &H10
Private Const SWP_NOSIZE& = &H1
Private Const SWP_SHOWWINDOW& = &H40

'===============================================================================
' Open file dialog APIs
'===============================================================================
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST

Type OPENFILENAME
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 lpstrFilter        As String
 lpstrCustomFilter  As String
 nMaxCustFilter     As Long
 nFilterIndex       As Long
 lpstrFile          As String
 nMaxFile           As Long
 lpstrFileTitle     As String
 nMaxFileTitle      As Long
 lpstrInitialDir    As String
 lpstrTitle         As String
 Flags              As Long
 nFileOffset        As Integer
 nFileExtension     As Integer
 lpstrDefExt        As String
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As String
End Type

Public OFN As OPENFILENAME

Public Declare Function CommDlgExtendedError Lib "COMDLG32.DLL" () As Long
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Enum FileDlgModes
 fdmOpenFile = 1
 fdmSaveFile
 fdmSaveFileNoConfirm
 fdmOpenFileOrPrompt
End Enum

Public Const PIC_FILTER1$ = "Pictures (*.bmp;*.dib;*.ico;*.gif;*.jpg;*.rle)|*.bmp;*.dib;*.ico;*.gif;*.jpg;*.rle|Bitmaps (*.bmp;*.dib;*.rle)|*.bmp;*.dib;*.rle|Icons (*.ico)|*.ico|Internet Images (*.gif;*.jpg)|*.gif;*.jpg"
'===============================================================================
'Choose color dialog APIs
'===============================================================================
Private Type TCHOOSECOLOR
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 rgbResult          As Long
 lpCustColors       As Long
 Flags              As Long
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public CustomColors(0 To 15) As Long
'===============================================================================

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4

Public PE As clsPaintEffects

Public Const FSMAIL$ = "fauzie811@yahoo.com"

Public Const INTERR$ = "An unexpected application error has occured!"
Public Const ERRTEXT$ = "If this problem continues, please contact me, at " + FSMAIL$ + ", quoting the above information."

Public Sub InitPaintEffects()
  If PE Is Nothing Then
    Set PE = New clsPaintEffects
  End If
End Sub

Public Sub Highlight(C As Control)
  With C
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Public Function IsInControl(ByVal hWnd As Long) As Boolean
  Dim P As POINTAPI
  GetCursorPos P
  If hWnd = WindowFromPoint(P.x, P.y) Then IsInControl = -1
End Function

Public Function SelectFile$(OwnerHWnd As Long, Optional Title$ = "", Optional Filter$ = "All Files (*.*)|*.*", Optional FilterIDX As Long = 0, Optional DefFile$, Optional DefPath$, Optional DefExt$, Optional ByVal FileMode As FileDlgModes = fdmOpenFile)
 Dim R As Long, SP As Long, ShortSize As Long, Z As Long
 With OFN
  .lStructSize = Len(OFN)
  .hWndOwner = OwnerHWnd
  .hInstance = App.hInstance
  .lpstrFilter = Replace$(Filter$, "|", Chr$(0)) & Chr$(0)
  .nFilterIndex = FilterIDX
  .lpstrFile = DefFile$ & String$(257 - Len(DefFile$), 0)
  .nMaxFile = Len(.lpstrFile) - 1
  .lpstrFileTitle = .lpstrFile
  .nMaxFileTitle = .nMaxFile
  .lpstrDefExt = DefExt$ & Chr$(0)
  .lpstrInitialDir = IIf(Len(DefPath$), DefPath$, CurDir$) & Chr$(0)
  .lpstrTitle = Title$ & Chr$(0)
  If FileMode = fdmSaveFile Or FileMode = fdmSaveFileNoConfirm Then
   .Flags = OFS_FILE_SAVE_FLAGS
   If FileMode = fdmSaveFile Then .Flags = .Flags Or OFN_OVERWRITEPROMPT
   R = GetSaveFileName(OFN)
  Else
   .Flags = OFS_FILE_OPEN_FLAGS
   If FileMode = fdmOpenFileOrPrompt Then .Flags = .Flags Or OFN_CREATEPROMPT
   R = GetOpenFileName(OFN)
  End If
  If R Then
   SP = InStr(.lpstrFile, Chr$(0))
   If SP Then .lpstrFile = Left$(.lpstrFile, SP - 1)
   SelectFile$ = Trim$(Replace$(.lpstrFile, Chr$(0), ""))
  Else
   Z = CommDlgExtendedError()
   If Z Then MsgBox "Unable to get filename(s)." & vbCr & vbCr & "CommDlgExtendedError returned " & Z, vbCritical
  End If
 End With
End Function

Public Function SelectColor(hWndParent As Long, DefColor As Long, Optional ShowExpDlg As Boolean = 0, Optional InitCustomColours As Boolean = -1) As Long
 Dim i
 Dim C As Long
 Dim CC As TCHOOSECOLOR
 Dim CT$
 'Initialise Custom Colours
 If InitCustomColours Then
  For i = 0 To 15
   CT$ = GetSetting$("Fauzie's Software", "CustomColours", CStr(i))
   CustomColors(i) = IIf(Len(CT$), Val(CT$), QBColor(15))
  Next
 End If
 'Show Dialog
 With CC
  .rgbResult = DefColor
  .hWndOwner = hWndParent
  .lpCustColors = VarPtr(CustomColors(0))
  .Flags = &H101
  If ShowExpDlg Then .Flags = .Flags Or &H2
  .lStructSize = Len(CC)
  C = ChooseColor(CC)
  If C Then
   SelectColor = .rgbResult
  Else
   SelectColor = -1
  End If
 End With
End Function

'======================================================================
'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
'Public Sub DrawGrad(lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hdc As Long, Optional bH As Boolean)
Public Sub DrawGrad(lEndColor As Long, lStartcolor As Long, Rct As RECT, ByVal hdc As Long, Optional bH As Boolean)
    On Error Resume Next
    
    ''Draw a Vertical Gradient in the current HDC
    Dim x As Long, y As Long, X2 As Long, Y2 As Long
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    x = Rct.Left: y = Rct.Top
    X2 = Rct.Right - Rct.Left: Y2 = Rct.Bottom - Rct.Top
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)
    
        
    For ni = 0 To IIf(bH, X2, Y2)
        
        If bH Then
            DrawLine x + ni, y, x + ni, Y2, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine x, y + ni, X2, y + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
        
    Next ni
End Sub
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function
'======================================================================

'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Public Sub DrawLine( _
           ByVal x As Long, _
           ByVal y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim POS     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
        MoveToEx cHdc, x, y, POS
        LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub
'======================================================================

Public Sub HideTip()
 On Error Resume Next
  Unload frmTooltip
 On Error GoTo 0
End Sub

Public Function ShowTip(ByVal Tip$, ByVal hWnd As Long, Optional ByVal Font As StdFont) As Boolean
 Const DX = -2   ' Offset from the mouse position.
 Const DY = 18
 Dim x As Long, y As Long
 Dim PT As POINTAPI
 On Error Resume Next
  GetCursorPos PT
  x = PT.x
  y = PT.y
  HideTip
  With frmTooltip
   If Not Font Is Nothing Then
    Set .lblTip.Font = Font
    Set .Font = Font
   End If
   .lblTip.Width = .TextWidth(Tip$)
   .lblTip.Caption = Tip$
   .lblTip.Refresh
   .CtlHWnd = hWnd
   .Move (x + DX) * Screen.TwipsPerPixelX, (y + DY) * Screen.TwipsPerPixelY, .lblTip.Width + (8 * Screen.TwipsPerPixelX), .lblTip.Height + (5 * Screen.TwipsPerPixelY)
   .tmrTip.Enabled = 0
   .tmrTip.Enabled = -1
   If .Left + .Width > Screen.Width Then .Left = Screen.Width - .Width
   If .Top + .Height > Screen.Height Then .Top = Screen.Height - .Height
   SetWindowPos .hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
  End With
  ShowTip = -1
 On Error GoTo 0
End Function

