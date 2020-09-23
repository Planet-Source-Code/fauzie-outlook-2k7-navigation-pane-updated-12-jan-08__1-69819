VERSION 5.00
Begin VB.UserControl NavPane 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   PropertyPages   =   "NavPane.ctx":0000
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   ToolboxBitmap   =   "NavPane.ctx":0037
   Begin VB.Timer tmrTip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   1800
   End
   Begin VB.Menu mPop 
      Caption         =   "Pop"
      Begin VB.Menu mMore 
         Caption         =   "Show &More Buttons"
      End
      Begin VB.Menu mFewer 
         Caption         =   "Show Fe&wer Buttons"
      End
      Begin VB.Menu mAddRem 
         Caption         =   "&Add or Remove Buttons"
         Begin VB.Menu mButts 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "NavPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' used to convert icons/bitmaps to stdPicture objects
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hpal As Long
End Type

' used to load the current hand cursor theme
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Const IDC_HAND As Long = 32649
Private myHand_handle As Long

Public Enum Themes
  [System Color] = 0
  [Blue] = 1
  [Silver] = 2
  [Black] = 3
  [Custom] = 99
End Enum

Private Const CaptionHeight = 26
Private Const ButtonHeight = 31

Public Redraw As Integer
Attribute Redraw.VB_VarMemberFlags = "40"

Dim bPic(1 To 4) As New cMemDC
Dim PicCapt As New cMemDC, PicGrip As New cMemDC
Dim PNT As POINTAPI, FirstButton As Integer, VisibleButton As Integer
Dim Resizing As Boolean, fX As Single, fY As Single
Dim RctCapt As RECT, RctBtm As RECT, RctGrip As RECT, RctConfig As RECT
Dim RctButton() As RECT
Dim RctButtonCapt() As RECT
Dim Tmp As Integer
Dim mBorderColor As OLE_COLOR
Dim bDown As Boolean, bDownButton As Integer, bCDown As Boolean

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Theme = 0 '99
Const m_def_CaptionForeColor = 0
Const m_def_ExpandedButtons = 3
Const m_def_ConfigButtonToolTip = "Configure buttons"
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_ToolTipFont As Font
Dim m_CaptionFont As Font
Dim m_Theme As Themes
Dim m_CaptionForeColor As OLE_COLOR
Dim m_ButtonCount As Integer
Dim m_ExpandedButtons As Integer
Dim m_ActiveButton As Integer
Dim m_Buttons() As New ButtonItem
Dim m_BaseColor As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_ButtonColorOver As OLE_COLOR
Dim m_ButtonColorDown As OLE_COLOR
Dim m_ShowConfigButton As Boolean
Dim m_ConfigButtonToolTip As String

Public ClientLeft As Single, ClientTop As Single, ClientWidth As Single, ClientHeight As Single
Attribute ClientLeft.VB_VarMemberFlags = "400"
Attribute ClientTop.VB_VarMemberFlags = "400"
Attribute ClientWidth.VB_VarMemberFlags = "400"
Attribute ClientHeight.VB_VarMemberFlags = "400"

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event ButtonChanged(Button As ButtonItem)
Event ButtonCountChanged(NewCount As Integer)
Event ThemeChanged(NewTheme As Themes)
Event Resize()

Private Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture
' Convert an icon/bitmap handle to a Picture object

On Error GoTo ExitRoutine

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture

ExitRoutine:
End Function

Public Function AboutBox()
Attribute AboutBox.VB_UserMemId = -552
  frmAbout.Show vbModal
End Function

Private Sub DrawAll()
  If Redraw = 0 Then Exit Sub
  InitPaintEffects
  Dim i As Integer
  Cls
  Line (0, 0)-(ScaleWidth, 0), m_BorderColor
  Line (0, 0)-(0, ScaleHeight), m_BorderColor
  Line (0, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1), m_BorderColor
  Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight), m_BorderColor
  
  BitBlt UserControl.hdc, 1, 1, 2, CaptionHeight, PicCapt.hdc, 0, 0, vbSrcCopy
  StretchBlt UserControl.hdc, 3, 1, ScaleWidth - 4, CaptionHeight, PicCapt.hdc, 2, 0, 35, 26, vbSrcCopy
  Line (0, CaptionHeight + 1)-(ScaleWidth, CaptionHeight + 1), m_BorderColor
  
  ' Button caption
  Set UserControl.Font = m_CaptionFont
  UserControl.ForeColor = m_CaptionForeColor
  SetRect RctCapt, 7, 1, ScaleWidth - 8, CaptionHeight
  
  ' Bottom Button
  SetRect RctBtm, 1, ScaleHeight - ButtonHeight, ScaleWidth - 1, ScaleHeight - 1
  StretchBlt UserControl.hdc, RctBtm.Left, RctBtm.Top, RctBtm.Right - RctBtm.Left, RctBtm.Bottom - RctBtm.Top, bPic(1).hdc, 0, 0, 40, 31, vbSrcCopy
  Line (0, RctBtm.Top - 1)-(ScaleWidth, RctBtm.Top - 1), m_BorderColor
  
  ' Config Button
  SetRect RctConfig, ScaleWidth - 19, ScaleHeight - ButtonHeight, ScaleWidth - 1, ScaleHeight - 1
  DrawConfigButton
  
  If m_ButtonCount <> 0 Then
    ReDim Preserve RctButton(m_ButtonCount) As RECT
    ReDim Preserve RctButtonCapt(m_ButtonCount) As RECT
    Call SetRects
    For i = m_ButtonCount To 1 Step -1
'      If m_Buttons(i).Visible Then VisibleButton = VisibleButton + 1
      DrawItem i
    Next
    For i = 1 To m_ButtonCount
      If m_Buttons(i).Visible Then FirstButton = i: Exit For
    Next
    
    ' Grip
    SetRect RctGrip, 1, RctButton(FirstButton).Top - 8, ScaleWidth - 1, 0
    RctGrip.Bottom = RctGrip.Top + 7
    BitBlt UserControl.hdc, (RctGrip.Right - 40) / 2, RctGrip.Top, 40, RctGrip.Bottom - RctGrip.Top, PicGrip.hdc, 0, 0, vbSrcCopy
    StretchBlt UserControl.hdc, 1, RctGrip.Top, (RctGrip.Right - 40) / 2, RctGrip.Bottom - RctGrip.Top, PicGrip.hdc, 0, 0, 2, 7, vbSrcCopy
    StretchBlt UserControl.hdc, (RctGrip.Right / 2) + 20, RctGrip.Top, (RctGrip.Right - 40) / 2, RctGrip.Bottom - RctGrip.Top, PicGrip.hdc, 0, 0, 2, 7, vbSrcCopy
    Line (0, RctGrip.Top - 1)-(ScaleWidth, RctGrip.Top - 1), m_BorderColor
  
    ' Draw the caption
    Set UserControl.Font = m_CaptionFont
    UserControl.ForeColor = m_CaptionForeColor
    DrawText hdc, m_Buttons(ActiveButton).Caption, Len(m_Buttons(ActiveButton).Caption), RctCapt, DT_SINGLELINE Or DT_VCENTER
  
    ' Set the client area
    ClientLeft = 1: ClientTop = RctCapt.Bottom + 2
    ClientWidth = ScaleWidth - 2
    ClientHeight = RctGrip.Top - RctCapt.Bottom - 3
  End If
End Sub

Private Sub DrawConfigButton(Optional State As Integer)
  If Not m_ShowConfigButton Then Exit Sub
  Dim C As Long, PT As POINTAPI
  With RctConfig
    SetRect RctConfig, ScaleWidth - 19, ScaleHeight - ButtonHeight, ScaleWidth - 1, ScaleHeight - 1
    StretchBlt UserControl.hdc, RctConfig.Left, RctConfig.Top, RctConfig.Right - RctConfig.Left, RctConfig.Bottom - RctConfig.Top, bPic(State + 1).hdc, 0, 0, 40, 31, vbSrcCopy
    C = .Top + (.Bottom - .Top) \ 2 - 2
    
    UserControl.ForeColor = getDarkColor(TranslateColor(m_BorderColor), 30)
    Call MoveToEx(UserControl.hdc, .Left + 7, C, PT)      'Top line, left
    Call LineTo(UserControl.hdc, .Left + 12, C)           'Top right
    Call MoveToEx(UserControl.hdc, .Left + 8, C + 1, PT)    'Mdl left
    Call LineTo(UserControl.hdc, .Left + 11, C + 1)         'Mdl right
    Call MoveToEx(UserControl.hdc, .Left + 9, C + 2, PT)    'Bot left
    Call LineTo(UserControl.hdc, .Left + 10, C + 2)          'Bot right
  End With
End Sub

Private Sub DrawItem(ButtonIndex As Integer, Optional State As Integer = 0)
  Dim n As Integer
  Dim PX As Single, PY As Single
  Dim PW As Single, PH As Single
  With m_Buttons(ButtonIndex)
    n = IIf(m_ExpandedButtons <= m_ButtonCount, m_ExpandedButtons, m_ButtonCount)
    Select Case State
    Case 0 ' Normal state
      If ActiveButton = ButtonIndex Then
        StretchBlt UserControl.hdc, RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Top, RctButton(ButtonIndex).Right - RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Bottom - RctButton(ButtonIndex).Top, bPic(3).hdc, 0, 0, 40, 31, vbSrcCopy
      Else
        StretchBlt UserControl.hdc, RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Top, RctButton(ButtonIndex).Right - RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Bottom - RctButton(ButtonIndex).Top, bPic(1).hdc, 0, 0, 40, 31, vbSrcCopy
      End If
    Case 1 ' Hover state
      If ActiveButton = ButtonIndex Then
        StretchBlt UserControl.hdc, RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Top, RctButton(ButtonIndex).Right - RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Bottom - RctButton(ButtonIndex).Top, bPic(4).hdc, 0, 0, 40, 31, vbSrcCopy
      Else
        StretchBlt UserControl.hdc, RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Top, RctButton(ButtonIndex).Right - RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Bottom - RctButton(ButtonIndex).Top, bPic(2).hdc, 0, 0, 40, 31, vbSrcCopy
      End If
    Case 2 ' Down state
      StretchBlt UserControl.hdc, RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Top, RctButton(ButtonIndex).Right - RctButton(ButtonIndex).Left, RctButton(ButtonIndex).Bottom - RctButton(ButtonIndex).Top, bPic(4).hdc, 0, 0, 40, 31, vbSrcCopy
    End Select
    
    ' Draw text
    If .Expanded Then
      Set UserControl.Font = m_Font
      UserControl.ForeColor = m_ForeColor
      DrawText hdc, .Caption, Len(.Caption), RctButtonCapt(ButtonIndex), DT_SINGLELINE Or DT_VCENTER
    End If
    
    Line (0, RctButton(ButtonIndex).Top - 1)-(ScaleWidth, RctButton(ButtonIndex).Top - 1), m_BorderColor
    
    ' Draw icon
    If Not .Icon Is Nothing And .Visible Then
      If .Expanded Then
        PW = ScaleX(.Icon.Width, vbHimetric, vbPixels)
        PH = ScaleY(.Icon.Height, vbHimetric, vbPixels)
        PX = RctButton(ButtonIndex).Left + (30 - PW) / 2
        PY = RctButton(ButtonIndex).Top + (ButtonHeight - PH) / 2
        If .Icon.Type = vbPicTypeIcon Then
          'DrawTransparentBitmap doesn't support icons
          PE.PaintStandardPicture hdc, .Icon, PX, PY, PW, PH, 0, 0
        Else
          If m_Buttons(ButtonIndex).UseMaskColor Then
             PE.PaintTransparentPicture hdc, .Icon, PX, PY, PW, PH, 0, 0, m_Buttons(ButtonIndex).MaskColor
          Else
             PE.PaintStandardPicture hdc, .Icon, PX, PY, PW, PH, 0, 0
          End If
        End If
      Else
        If Not .SmallIcon Is Nothing Then
          PW = ScaleX(.SmallIcon.Width, vbHimetric, vbPixels)
          PH = ScaleY(.SmallIcon.Height, vbHimetric, vbPixels)
          PX = RctButton(ButtonIndex).Left + (24 - PW) / 2
          PY = RctButton(ButtonIndex).Top + (ButtonHeight - PH) / 2
          If .SmallIcon.Type = vbPicTypeIcon Then
            'DrawTransparentBitmap doesn't support icons
            PE.PaintStandardPicture hdc, .SmallIcon, PX, PY, PW, PH, 0, 0
          Else
            If m_Buttons(ButtonIndex).UseMaskColor Then
               PE.PaintTransparentPicture hdc, .SmallIcon, PX, PY, PW, PH, 0, 0, m_Buttons(ButtonIndex).MaskColor
            Else
               PE.PaintStandardPicture hdc, .SmallIcon, PX, PY, PW, PH, 0, 0
            End If
          End If
        End If
      End If
    End If
  End With
End Sub

Private Sub SetRects()
  Dim n As Integer, i As Integer, li As Integer, j As Integer
  n = IIf(m_ExpandedButtons <= m_ButtonCount, m_ExpandedButtons, m_ButtonCount)
  j = 1
'  VisibleButton = 0
  For i = 1 To m_ButtonCount
    With m_Buttons(i)
      If .Visible Then
'        VisibleButton = VisibleButton + 1
        If j <= n Then
          SetRect RctButton(i), 1, ScaleHeight - ((RctBtm.Bottom - RctBtm.Top) + 1) - ((n - j + 1) * (ButtonHeight + 1)), ScaleWidth - 1, 0
          RctButton(i).Bottom = RctButton(i).Top + ButtonHeight
          SetRect RctButtonCapt(i), RctButton(i).Left + 33, RctButton(i).Top, RctButton(i).Right, RctButton(i).Bottom
          .Expanded = True
        Else
          SetRect RctButton(i), ScaleWidth - 45 - (-(j - VisibleButton) * 23) + 3, RctBtm.Top, 0, RctBtm.Bottom
          RctButton(i).Right = RctButton(i).Left + 23
          SetRect RctButtonCapt(i), 0, 0, 0, 0
          .Expanded = False
        End If
        li = i
        j = j + 1
      Else
        SetRect RctButton(i), 0, 0, 0, 0
        SetRect RctButtonCapt(i), 0, 0, 0, 0
        .Expanded = False
      End If
    End With
  Next
End Sub

Private Sub mButts_Click(Index As Integer)
  mButts(Index).Checked = Not mButts(Index).Checked
'  m_Buttons(Index).Visible = mButts(Index).Checked
  ButtonVisible(Index) = mButts(Index).Checked
'  DrawAll
End Sub

Private Sub mFewer_Click()
  ExpandedButtons = ExpandedButtons - 1
End Sub

Private Sub mMore_Click()
  ExpandedButtons = ExpandedButtons + 1
End Sub

Private Sub Timer1_Timer()
  GetCursorPos PNT
  If WindowFromPoint(PNT.x, PNT.y) <> UserControl.hWnd Then ResetTip: Tmp = -1: DrawAll: Timer1 = False
End Sub

Private Sub ShowCtlTip(Tip$, Optional Force As Boolean)
 On Error Resume Next
  If Tip$ = "" Then
   HideTip
   tmrTip.Enabled = 0
   Extender.ToolTipText = ""
  Else
   tmrTip.Enabled = Ambient.UserMode
   tmrTip.Tag = Tip$
   If Force Then Call tmrTip_Timer
  End If
 On Error GoTo 0
End Sub

Private Sub ResetTip()
 On Error Resume Next
  tmrTip.Enabled = 0
  HideTip
  Extender.ToolTipText = ""
 On Error GoTo 0
End Sub

Private Sub tmrTip_Timer()
 On Error Resume Next
  ResetTip
  If IsInControl(hWnd) Then
   If ShowTip(tmrTip.Tag, GetActiveWindow(), m_ToolTipFont) = 0 Then
    Extender.ToolTipText = tmrTip.Tag
   End If
  End If
 On Error GoTo 0
End Sub

Private Sub UserControl_Initialize()
'  Set WProc = New cSubclass
  
  myHand_handle = LoadCursor(0, IDC_HAND)
  UserControl.MouseIcon = HandleToPicture(myHand_handle, False)
  Dim i As Integer
  For i = 1 To 4
    bPic(i).Create 40, 31
  Next
  PicCapt.Create 40, 26
  PicGrip.Create 40, 7
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  ResetTip
  If Button = 1 Then
    i = MatchRect(x, y)
    If i = -1 Then
      GoTo tEnd
    ElseIf i = -2 Then
      Resizing = True
      fX = x: fY = RctButton(FirstButton).Top
    ElseIf i = -3 Then
      bCDown = True
      Tmp = 0
      DrawConfigButton 3
      UserControl.Refresh
    Else
      bDown = True
      bDownButton = i
      DrawItem i, 2
    End If
  End If
tEnd:
  
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, PW As Single, PH As Single
  Timer1 = True
  If Resizing Then
    If y - fY > 30 Then
      ExpandedButtons = ExpandedButtons - 1
      fY = RctButton(FirstButton).Top
    ElseIf y - fY < -30 Then
      ExpandedButtons = ExpandedButtons + 1
      fY = RctButton(FirstButton).Top
    End If
    If ExpandedButtons > VisibleButton Then ExpandedButtons = VisibleButton
    If ExpandedButtons < 0 Then ExpandedButtons = 0
    UserControl.Refresh
    RaiseEvent Resize
  Else
    i = MatchRect(x, y)
    If i <> -2 And Button = 0 Then MousePointer = vbArrow
    If i = -1 Then
      ResetTip
      DrawAll
      Tmp = 0
      GoTo tEnd
    ElseIf i = -2 Then
      ResetTip
      If Button = 0 Then MousePointer = vbSizeNS
      Tmp = 0
    ElseIf i = -3 Then
      If Tmp <> -3 Then
        Debug.Print Tmp
        If Tmp > 0 And Tmp <= m_ButtonCount Then
          DrawItem Tmp
        End If
        If Button = 0 Then ShowCtlTip m_ConfigButtonToolTip, frmTooltip.Visible
        DrawConfigButton IIf(bCDown, 3, 1)
        UserControl.Refresh
        
        Tmp = -3
      End If
    Else
      DrawConfigButton
      UserControl.Refresh
'      If Tmp <> i Then UserControl.MousePointer = vbCustom
      If Button = 0 Then
        If Tmp <> i Then
'          ResetTip
          Set UserControl.Font = m_Font
          UserControl.ForeColor = m_ForeColor
          If Tmp > 0 And Tmp <= m_ButtonCount Then
            DrawItem Tmp
          End If
          
          DrawItem i, 1
          ShowCtlTip IIf(m_Buttons(i).ToolTipText <> "", m_Buttons(i).ToolTipText, IIf(m_Buttons(i).Expanded, "", m_Buttons(i).Caption)), frmTooltip.Visible
'          Extender.ToolTipText = IIf(m_Buttons(i).ToolTipText <> "", m_Buttons(i).ToolTipText, m_Buttons(i).Caption)
          Tmp = i
        End If
      Else
        Tmp = i
      End If
      If bDown Then
        If bDownButton <> i Then
          DrawItem bDownButton
        Else
          DrawItem bDownButton, 2
        End If
      End If
    End If
  End If
  
tEnd:
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  Tmp = 0
  Resizing = False
  DrawConfigButton
  i = MatchRect(x, y)
  If bDown Then
    If bDownButton = i Then
      ActiveButton = i
      DrawItem m_ActiveButton, 2
      RaiseEvent ButtonChanged(m_Buttons(i))
    End If
    bDown = False
  ElseIf bCDown And i = -3 Then
    SetUpMenu
    PopupMenu mPop, , UserControl.ScaleWidth, RctConfig.Top + (RctConfig.Bottom - RctConfig.Top) \ 2 - 2
    bCDown = False
  End If
End Sub

Private Sub UserControl_Resize()
  DrawAll
  RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
  hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color of the buttons."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  Redraw = 1
  DrawAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets a Font of the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set m_Font = New_Font
  PropertyChanged "Font"
  Redraw = 1
  DrawAll
End Property

Public Property Get ToolTipFont() As Font
Attribute ToolTipFont.VB_Description = "Returns/sets a Font of the button tooltip."
  Set ToolTipFont = m_ToolTipFont
End Property

Public Property Set ToolTipFont(ByVal New_ToolTipFont As Font)
  Set m_ToolTipFont = New_ToolTipFont
  PropertyChanged "ToolTipFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
  Redraw = 1
  DrawAll
End Sub

Public Property Get ButtonCount() As Integer
Attribute ButtonCount.VB_Description = "Returns the count of the buttons on the Navigation Pane"
Attribute ButtonCount.VB_MemberFlags = "400"
  ButtonCount = m_ButtonCount
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get CaptionFont() As Font
Attribute CaptionFont.VB_ProcData.VB_Invoke_Property = ";Font"
  Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
  Set m_CaptionFont = New_CaptionFont
  PropertyChanged "CaptionFont"
  Redraw = 1
  DrawAll
End Property

Public Property Get ButtonCaption(ByVal Index As Integer) As String
Attribute ButtonCaption.VB_MemberFlags = "400"
  ButtonCaption = m_Buttons(Index).Caption
End Property

Public Property Let ButtonCaption(ByVal Index As Integer, ByVal ButtonCaption As String)
  m_Buttons(Index).Caption = ButtonCaption
  Refresh
  PropertyChanged "ButtonCaption"
End Property

Public Property Get ButtonDescription(ByVal Index As Integer) As String
Attribute ButtonDescription.VB_MemberFlags = "400"
  ButtonDescription = m_Buttons(Index).Description
End Property

Public Property Let ButtonDescription(ByVal Index As Integer, ByVal ButtonDescription As String)
  m_Buttons(Index).Description = ButtonDescription
  PropertyChanged "ButtonDescription"
End Property

Public Property Get ButtonToolTipText(ByVal Index As Integer) As String
Attribute ButtonToolTipText.VB_MemberFlags = "400"
  ButtonToolTipText = m_Buttons(Index).ToolTipText
End Property

Public Property Let ButtonToolTipText(ByVal Index As Integer, ByVal NewStr As String)
 m_Buttons(Index).ToolTipText = NewStr
 PropertyChanged "ButtonToolTipText"
End Property

Public Property Get ButtonKey(ByVal Index As Integer) As String
Attribute ButtonKey.VB_MemberFlags = "400"
 ButtonKey = m_Buttons(Index).Key
End Property

Public Property Let ButtonKey(ByVal Index As Integer, ByVal ButtonKey As String)
 m_Buttons(Index).Key = ButtonKey
 PropertyChanged "ButtonKey"
End Property

Public Property Get ButtonUseMaskColor(ByVal Index As Integer) As Boolean
Attribute ButtonUseMaskColor.VB_MemberFlags = "400"
  ButtonUseMaskColor = m_Buttons(Index).UseMaskColor
End Property

Public Property Let ButtonUseMaskColor(ByVal Index As Integer, ByVal State As Boolean)
  If m_Buttons(Index).UseMaskColor <> State Then
    m_Buttons(Index).UseMaskColor = State
    Refresh
    PropertyChanged "ButtonUseMaskColor"
  End If
End Property

Public Property Get ButtonMaskColor(ByVal Index As Integer) As OLE_COLOR
Attribute ButtonMaskColor.VB_MemberFlags = "400"
  ButtonMaskColor = m_Buttons(Index).MaskColor
End Property

Public Property Let ButtonMaskColor(ByVal Index As Integer, ByVal ButtonMaskColor As OLE_COLOR)
  If m_Buttons(Index).MaskColor <> ButtonMaskColor Then
    m_Buttons(Index).MaskColor = ButtonMaskColor
    Refresh
    PropertyChanged "ButtonMaskColor"
  End If
End Property

Public Property Get ButtonIcon(ByVal Index As Integer) As StdPicture
Attribute ButtonIcon.VB_MemberFlags = "400"
 Set ButtonIcon = m_Buttons(Index).Icon
End Property

Public Property Set ButtonIcon(ByVal Index As Integer, ByVal ButtonIcon As StdPicture)
  Set m_Buttons(Index).Icon = ButtonIcon
  Refresh
  PropertyChanged "ButtonIcon"
End Property

Public Property Get ButtonSmallIcon(ByVal Index As Integer) As StdPicture
Attribute ButtonSmallIcon.VB_MemberFlags = "400"
 Set ButtonSmallIcon = m_Buttons(Index).SmallIcon
End Property

Public Property Set ButtonSmallIcon(ByVal Index As Integer, ByVal ButtonSmallIcon As StdPicture)
  Set m_Buttons(Index).SmallIcon = ButtonSmallIcon
  Refresh
  PropertyChanged "ButtonSmallIcon"
End Property

Public Property Get ButtonVisible(ByVal Index As Integer) As Boolean
Attribute ButtonVisible.VB_MemberFlags = "400"
  ButtonVisible = m_Buttons(Index).Visible
End Property

Public Property Let ButtonVisible(ByVal Index As Integer, ByVal State As Boolean)
  If m_Buttons(Index).Visible <> State Then
    m_Buttons(Index).Visible = State
    If State Then VisibleButton = VisibleButton + 1 Else VisibleButton = VisibleButton - 1
    If m_ExpandedButtons > VisibleButton Then m_ExpandedButtons = VisibleButton
    Refresh
    PropertyChanged "ButtonVisible"
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Theme() As Themes
Attribute Theme.VB_Description = "Returns/sets the color preset for the Navigation Pane."
Attribute Theme.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As Themes)
  m_Theme = New_Theme
  PropertyChanged "Theme"
'  SetBorder
  SetColorScheme m_Theme
  RaiseEvent ThemeChanged(m_Theme)
  Redraw = 1
  DrawAll
End Property

Public Property Get ExpandedButtons() As Integer
Attribute ExpandedButtons.VB_Description = "Returns/sets a value that determines how many buttons are expanded."
Attribute ExpandedButtons.VB_ProcData.VB_Invoke_Property = "General"
  ExpandedButtons = m_ExpandedButtons
End Property

Public Property Let ExpandedButtons(ByVal New_ExpandedButtons As Integer)
  m_ExpandedButtons = New_ExpandedButtons
  PropertyChanged "ExpandedButtons"
  Redraw = 1
  DrawAll
End Property

Public Property Get ActiveButton() As Integer
Attribute ActiveButton.VB_Description = "Returns/sets the current active button."
Attribute ActiveButton.VB_ProcData.VB_Invoke_Property = "General"
Attribute ActiveButton.VB_MemberFlags = "200"
  ActiveButton = m_ActiveButton
End Property

Public Property Let ActiveButton(ByVal New_ActiveButton As Integer)
  m_ActiveButton = New_ActiveButton
  PropertyChanged "ActiveButton"
  Redraw = 1
  DrawAll
  RaiseEvent Resize
  If m_ActiveButton > 0 Then RaiseEvent ButtonChanged(m_Buttons(m_ActiveButton))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionForeColor() As OLE_COLOR
Attribute CaptionForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  CaptionForeColor = m_CaptionForeColor
End Property

Public Property Let CaptionForeColor(ByVal New_CaptionForeColor As OLE_COLOR)
  m_CaptionForeColor = New_CaptionForeColor
  PropertyChanged "CaptionForeColor"
  Redraw = 1
  DrawAll
End Property

Public Property Get ShowConfigButton() As Boolean
  ShowConfigButton = m_ShowConfigButton
End Property

Public Property Let ShowConfigButton(ByVal New_ShowConfigButton As Boolean)
  m_ShowConfigButton = New_ShowConfigButton
  PropertyChanged "ShowConfigButton"
  Redraw = 1
  DrawAll
End Property

Public Property Get ConfigButtonToolTip() As String
Attribute ConfigButtonToolTip.VB_Description = "Returns/sets the text displayed when the mouse is paused over the Config Button."
  ConfigButtonToolTip = m_ConfigButtonToolTip
End Property

Public Property Let ConfigButtonToolTip(ByVal New_ConfigButtonToolTip As String)
  m_ConfigButtonToolTip = New_ConfigButtonToolTip
  PropertyChanged "ConfigButtonToolTip"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddButton() As Integer
  AddButton = AddButtonEx()
  Redraw = 1
  DrawAll
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ButtonColorOver = RGB(255, 215, 103)
  m_ButtonColorDown = RGB(255, 171, 63)
  m_BaseColor = vb3DLight
  m_ForeColor = m_def_ForeColor
  Set m_ToolTipFont = Ambient.Font
  Set m_Font = Ambient.Font
  m_Font.Bold = True
  Set m_CaptionFont = Ambient.Font
  m_CaptionFont.Bold = True
  m_CaptionFont.Size = 12
  m_Theme = m_def_Theme
  m_ExpandedButtons = m_def_ExpandedButtons
  m_ActiveButton = 0
  m_CaptionForeColor = m_def_CaptionForeColor
  m_ShowConfigButton = True
  m_ConfigButtonToolTip = m_def_ConfigButtonToolTip
  SetColorScheme m_Theme
  Redraw = 1
  DrawAll
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim i As Integer
  m_ButtonColorOver = PropBag.ReadProperty("ButtonColorOver", RGB(255, 215, 103))
  m_ButtonColorDown = PropBag.ReadProperty("ButtonColorDown", RGB(255, 171, 63))
  m_BorderColor = PropBag.ReadProperty("BorderColor", vb3DLight)
  m_BaseColor = PropBag.ReadProperty("BaseColor", vb3DLight)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set m_ToolTipFont = PropBag.ReadProperty("ToolTipFont", Ambient.Font)
  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
  m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
  m_ExpandedButtons = PropBag.ReadProperty("ExpandedButtons", m_def_ExpandedButtons)
  m_ActiveButton = PropBag.ReadProperty("ActiveButton", 0)
  m_CaptionForeColor = PropBag.ReadProperty("CaptionForeColor", m_def_CaptionForeColor)
  m_ShowConfigButton = PropBag.ReadProperty("ShowConfigButton", True)
  m_ConfigButtonToolTip = PropBag.ReadProperty("ConfigButtonToolTip", m_def_ConfigButtonToolTip)
  m_ButtonCount = PropBag.ReadProperty("ButtonCount", 0)
  ReDim m_Buttons(m_ButtonCount) As New ButtonItem
  For i = 1 To m_ButtonCount
    'Load Buttons
    With m_Buttons(i)
      .Caption = PropBag.ReadProperty("ButtonCaption" & i, "")
      .Description = PropBag.ReadProperty("ButtonDescription" & i, "")
      .Key = PropBag.ReadProperty("ButtonKey" & i, "")
      .UseMaskColor = PropBag.ReadProperty("ButtonUseMaskColor" & i, -1)
      .Visible = PropBag.ReadProperty("ButtonVisible" & i, -1)
      .MaskColor = PropBag.ReadProperty("ButtonMaskColor" & i, QBColor(13))
      Set .Icon = PropBag.ReadProperty("ButtonIcon" & i, Nothing)
      Set .SmallIcon = PropBag.ReadProperty("ButtonSmallIcon" & i, Nothing)
      .ToolTipText = PropBag.ReadProperty("ButtonToolTipText" & i, "")
      If .Visible Then VisibleButton = VisibleButton + 1
    End With
  Next
  SetColorScheme m_Theme
  Redraw = 1
  DrawAll
  
'  If Ambient.UserMode Then
'    With WProc
'      .Start UserControl.hwnd
'      .AttachAfterMSG WM_MOUSELEAVE
'    End With
'  End If
End Sub

Private Sub UserControl_Terminate()
  Dim i%
  For i = 1 To 4
    Set bPic(i) = Nothing
  Next
'  WProc.UnSubclass UserControl.hwnd
'  Set WProc = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Dim i As Integer
  Call PropBag.WriteProperty("ButtonColorOver", m_ButtonColorOver, RGB(255, 215, 103))
  Call PropBag.WriteProperty("ButtonColorDown", m_ButtonColorDown, RGB(255, 171, 63))
  Call PropBag.WriteProperty("BorderColor", m_BorderColor, vb3DLight)
  Call PropBag.WriteProperty("BaseColor", m_BaseColor, vb3DLight)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("ToolTipFont", m_ToolTipFont, Ambient.Font)
  Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
  Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
  Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
  Call PropBag.WriteProperty("ExpandedButtons", m_ExpandedButtons, m_def_ExpandedButtons)
  Call PropBag.WriteProperty("ActiveButton", m_ActiveButton, 0)
  Call PropBag.WriteProperty("CaptionForeColor", m_CaptionForeColor, m_def_CaptionForeColor)
  Call PropBag.WriteProperty("ShowConfigButton", m_ShowConfigButton, True)
  Call PropBag.WriteProperty("ConfigButtonToolTip", m_ConfigButtonToolTip, m_def_ConfigButtonToolTip)
  Call PropBag.WriteProperty("ButtonCount", m_ButtonCount)
  For i = 1 To m_ButtonCount
    With m_Buttons(i)
      PropBag.WriteProperty "ButtonCaption" & i, .Caption, ""
      PropBag.WriteProperty "ButtonDescription" & i, .Description, ""
      PropBag.WriteProperty "ButtonKey" & i, .Key, ""
      PropBag.WriteProperty "ButtonUseMaskColor" & i, .UseMaskColor, -1
      PropBag.WriteProperty "ButtonVisible" & i, .Visible, -1
      PropBag.WriteProperty "ButtonMaskColor" & i, .MaskColor, QBColor(13)
      PropBag.WriteProperty "ButtonIcon" & i, .Icon, Nothing
      PropBag.WriteProperty "ButtonSmallIcon" & i, .SmallIcon, Nothing
      PropBag.WriteProperty "ButtonToolTipText" & i, .ToolTipText, ""
    End With
  Next
End Sub

Private Sub RaiseErrorEx(ByVal ProcName$, ByVal ErrNum As Long, Optional ByVal ErrMsg$ = "")
  If Ambient.UserMode Then
    '"Runtime" - raise error
    If Len(ErrMsg$) Then
      Err.Raise ErrNum, App.EXEName & "." & TypeName(Me) & ":" & ProcName$, ErrMsg$
    Else
      Err.Raise ErrNum, App.EXEName & "." & TypeName(Me) & ":" & ProcName$
    End If
  Else
    '"Design time" - display error
    If Len(ErrMsg$) = 0 Then
      On Error Resume Next
      Error ErrNum
      ErrMsg$ = Err.Description
      On Error GoTo 0
    End If
    VBA.MsgBox INTERR$ & vbCr & vbCr & ErrMsg$ & " (" & ErrNum & ")" & vbCr & vbCr & ERRTEXT$, vbCritical, App.EXEName & "." & TypeName(Me)
  End If
End Sub

Public Function RemoveButton(ButtonIndex As Integer) As Integer
  Dim i As Integer
  If ButtonIndex < m_ButtonCount Then
    If m_Buttons(ButtonIndex).Visible Then VisibleButton = VisibleButton - 1
    For i = ButtonIndex To m_ButtonCount - 1
      m_Buttons(i).Key = m_Buttons(i + 1).Key
      Set m_Buttons(i).Icon = m_Buttons(i + 1).Icon
      m_Buttons(i).ToolTipText = m_Buttons(i + 1).ToolTipText
      m_Buttons(i).MaskColor = m_Buttons(i + 1).MaskColor
      m_Buttons(i).UseMaskColor = m_Buttons(i + 1).UseMaskColor
      m_Buttons(i).Caption = m_Buttons(i + 1).Caption
    Next
  ElseIf ButtonIndex < 1 Or ButtonIndex > ButtonCount Then
    RaiseErrorEx "RemoveButton", 380
  Else
      
  End If
  If ActiveButton = ButtonIndex Then ActiveButton = 1
  m_ButtonCount = m_ButtonCount - 1
  ReDim Preserve m_Buttons(m_ButtonCount) As New ButtonItem
  If ExpandedButtons > m_ButtonCount Then ExpandedButtons = m_ButtonCount
  PropertyChanged "ButtonCount"
  RaiseEvent ButtonCountChanged(m_ButtonCount)
  RemoveButton = m_ButtonCount
  Redraw = 1
  DrawAll
End Function

Public Function AddButtonEx(Optional Caption$ = "", Optional Key$ = "", Optional ToolTipText$ = "", Optional MaskColor As OLE_COLOR = 16711935, Optional UseMaskColor As Boolean = -1, Optional Visible As Boolean = -1, Optional Icon As StdPicture = Nothing, Optional SmallIcon As StdPicture = Nothing) As Integer
  m_ButtonCount = m_ButtonCount + 1
  ReDim Preserve m_Buttons(m_ButtonCount) As New ButtonItem
  With m_Buttons(m_ButtonCount)
    .Caption = Caption$
    .Key = Key$
    .ToolTipText = ToolTipText$
    .MaskColor = MaskColor
    .UseMaskColor = UseMaskColor
    .Visible = Visible
    Set .Icon = Icon
    Set .SmallIcon = SmallIcon
  End With
  If Visible Then VisibleButton = VisibleButton + 1
  Redraw = 1
  DrawAll
  PropertyChanged "ButtonCount"
  RaiseEvent ButtonCountChanged(m_ButtonCount)
  AddButtonEx = m_ButtonCount
End Function

Public Function SwapButton(ByVal CurIndex As Integer, ByVal NewIndex As Integer) As Integer
  Dim CI, ni, i, S, pActive As Integer
  Dim T As New ButtonItem
  CI = CurIndex
  ni = NewIndex
  If CI < 1 Or CI > m_ButtonCount Or ni < 1 Or ni > m_ButtonCount Then
    RaiseErrorEx "SwapButton", 380
  Else
    If ni > CI Then S = 1 Else S = -1
    For i = CI To ni - S Step S
      Set T = m_Buttons(i)
      If m_ActiveButton - 1 = i Then pActive = i
      Set m_Buttons(i) = m_Buttons(i + S)
      Set m_Buttons(i + S) = T
    Next
    If m_ActiveButton = CI Then m_ActiveButton = ni Else m_ActiveButton = pActive
    Refresh
    SwapButton = NewIndex
  End If
  Set T = Nothing
End Function

Private Function MatchRect(ByVal x As Long, ByVal y As Long) As Integer
  Dim Rct As RECT
  Dim i, Ok As Long
  
  Ok = GetClientRect(hWnd, Rct)
  If PtInRect(Rct, x, y) <> 0 Then
    If m_ButtonCount <> 0 Then
      If PtInRect(RctGrip, x, y) Then MatchRect = -2: Exit Function
      If PtInRect(RctConfig, x, y) And m_ShowConfigButton Then MatchRect = -3: Exit Function
      For i = m_ButtonCount To 1 Step -1
        If PtInRect(RctButton(i), x, y) Then
          MatchRect = i
          Exit Function
        End If
      Next
    End If
  
  End If
  
  MatchRect = -1
End Function

Private Sub SetPics()
  Dim Rct As RECT
  Dim clr As Long, hclr As Long, dclr As Long, hdclr As Long
  clr = m_BaseColor
  hclr = m_ButtonColorOver
  dclr = m_ButtonColorDown
  hdclr = ShiftColor(dclr, -50, True)
  hdclr = ShiftColorOXP(hdclr, -90)
  hdclr = getLightColor(hdclr, 70)
'  hdclr = RGB(251, 140, 60)
  
  Rct.Top = 0: Rct.Bottom = 12
  Rct.Right = 40
  DrawGrad ShiftColorOXP(clr, 170), ShiftColorOXP(clr, 72), Rct, bPic(1).hdc
  DrawGrad ShiftColor(hclr, 122, False), ShiftColorOXP(hclr, 109), Rct, bPic(2).hdc
  DrawGrad ShiftColorOXP(dclr, 142), ShiftColorOXP(dclr, 58), Rct, bPic(3).hdc
  DrawGrad ShiftColor(hdclr, 50, True), ShiftColor(hdclr, 30, True), Rct, bPic(4).hdc
  
  Rct.Top = 12: Rct.Bottom = 31
  DrawGrad clr, ShiftColorOXP(clr, 60), Rct, bPic(1).hdc
  DrawGrad hclr, ShiftColorOXP(hclr, 90), Rct, bPic(2).hdc
  DrawGrad dclr, ShiftColor(dclr, 57, False), Rct, bPic(3).hdc
  DrawGrad hdclr, ShiftColor(hdclr, 72, True), Rct, bPic(4).hdc

  ' Caption
  Rct.Top = 1: Rct.Bottom = 26
  Rct.Left = 1: Rct.Right = 40
  DrawGrad ShiftColorOXP(clr, 171), ShiftColorOXP(clr, 7), Rct, PicCapt.hdc
  DrawLine 0, 0, 40, 0, PicCapt.hdc, vbWhite
  DrawLine 0, 0, 0, 26, PicCapt.hdc, vbWhite
  
  'Grip
  Rct.Top = 1: Rct.Bottom = 7
  Rct.Left = 0
  DrawLine 0, 0, 40, 0, PicGrip.hdc, vbWhite
  
  DrawGrad ShiftColorOXP(clr, 171), ShiftColorOXP(clr, 7), Rct, PicGrip.hdc
  
  DrawLine 10, 2, 12, 2, PicGrip.hdc, m_BorderColor
  DrawLine 10, 2, 10, 4, PicGrip.hdc, m_BorderColor
  
  DrawLine 14, 2, 16, 2, PicGrip.hdc, m_BorderColor
  DrawLine 14, 2, 14, 4, PicGrip.hdc, m_BorderColor
  
  DrawLine 18, 2, 20, 2, PicGrip.hdc, m_BorderColor
  DrawLine 18, 2, 18, 4, PicGrip.hdc, m_BorderColor
  
  DrawLine 22, 2, 24, 2, PicGrip.hdc, m_BorderColor
  DrawLine 22, 2, 22, 4, PicGrip.hdc, m_BorderColor
  
  DrawLine 26, 2, 28, 2, PicGrip.hdc, m_BorderColor
  DrawLine 26, 2, 26, 4, PicGrip.hdc, m_BorderColor
  
  DrawLine 12, 4, 12, 2, PicGrip.hdc, vbWhite
  DrawLine 12, 4, 10, 4, PicGrip.hdc, vbWhite
  
  DrawLine 16, 4, 16, 2, PicGrip.hdc, vbWhite
  DrawLine 16, 4, 14, 4, PicGrip.hdc, vbWhite
  
  DrawLine 20, 4, 20, 2, PicGrip.hdc, vbWhite
  DrawLine 20, 4, 18, 4, PicGrip.hdc, vbWhite
  
  DrawLine 24, 4, 24, 2, PicGrip.hdc, vbWhite
  DrawLine 24, 4, 22, 4, PicGrip.hdc, vbWhite
  
  DrawLine 28, 4, 28, 2, PicGrip.hdc, vbWhite
  DrawLine 28, 4, 26, 4, PicGrip.hdc, vbWhite
End Sub

Private Function ShiftColor(ByVal lColor As Long, ByVal Value As Long, Optional isXP As Boolean = False, Optional isSoft As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result

Dim Color As Long
Dim Red As Long, Blue As Long, Green As Long

Color = TranslateColor(lColor)

'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
If isSoft Then Value = Value \ 2

If Not isXP Then 'for XP button i use a work-aroud that works fine
    Blue = ((Color \ &H10000) Mod &H100) + Value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value

'a bit of optimization done here, values will overflow a
' byte only in one direction... eg: if we added 32 to our
' color, then only a > 255 overflow can occurr.
If Value > 0 Then
    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255
ElseIf Value < 0 Then
    If Red < 0 Then Red = 0
    If Green < 0 Then Green = 0
    If Blue < 0 Then Blue = 0
End If

'more optimization by replacing the RGB function by its correspondent calculation
ShiftColor = Red + 256& * Green + 65536 * Blue
End Function

Private Function ShiftColorOXP(ByVal lColor As Long, Optional ByVal Base As Long = &HB0) As Long
Dim Red As Long, Blue As Long, Green As Long
Dim Delta As Long, theColor As Long

theColor = TranslateColor(lColor)

Blue = ((theColor \ &H10000) Mod &H100)
Green = ((theColor \ &H100) Mod &H100)
Red = (theColor And &HFF)
Delta = &HFF - Base

Blue = Base + Blue * Delta \ &HFF
Green = Base + Green * Delta \ &HFF
Red = Base + Red * Delta \ &HFF

If Red > 255 Then Red = 255
If Green > 255 Then Green = 255
If Blue > 255 Then Blue = 255

If Red < 0 Then Red = 0
If Green < 0 Then Green = 0
If Blue < 0 Then Blue = 0

ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Public Property Get BaseColor() As OLE_COLOR
Attribute BaseColor.VB_Description = "Returns/sets the base color of the Navigation Pane."
Attribute BaseColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BaseColor.VB_UserMemId = -501
  BaseColor = m_BaseColor
End Property

Public Property Let BaseColor(ByVal New_BaseColor As OLE_COLOR)
  m_BaseColor = New_BaseColor
  PropertyChanged "BaseColor"
  m_Theme = Custom
  SetPics
  Redraw = 1
  DrawAll
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the border color of the Navigation Pane."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  PropertyChanged "BorderColor"
  m_Theme = Custom
  SetPics
  Redraw = 1
  DrawAll
End Property

Public Property Get ButtonColorOver() As OLE_COLOR
Attribute ButtonColorOver.VB_Description = "Returns/sets the button color when the mouse is over it."
Attribute ButtonColorOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ButtonColorOver = m_ButtonColorOver
End Property

Public Property Let ButtonColorOver(ByVal New_ButtonColorOver As OLE_COLOR)
  m_ButtonColorOver = New_ButtonColorOver
  PropertyChanged "ButtonColorOver"
  m_Theme = Custom
  SetPics
  Redraw = 1
  DrawAll
End Property

Public Property Get ButtonColorDown() As OLE_COLOR
  ButtonColorDown = m_ButtonColorDown
End Property

Public Property Let ButtonColorDown(ByVal New_ButtonColorDown As OLE_COLOR)
  m_ButtonColorDown = New_ButtonColorDown
  PropertyChanged "ButtonColorDown"
  m_Theme = Custom
  SetPics
  Redraw = 1
  DrawAll
End Property

Private Sub SetColorScheme(cTheme As Themes)
  Select Case cTheme
  Case [System Color]
    m_BaseColor = vb3DLight
    m_BorderColor = vbButtonShadow
    m_ButtonColorOver = RGB(255, 215, 103) 'vb3DHighlight 'vbButtonFace
    m_ButtonColorDown = RGB(255, 171, 63)
  Case Blue
    m_BaseColor = RGB(173, 209, 255)
    m_BorderColor = RGB(101, 147, 207)
    m_ButtonColorOver = RGB(255, 215, 103)
    m_ButtonColorDown = RGB(255, 171, 63)
  Case Silver
    m_BaseColor = RGB(197, 199, 209)
    m_BorderColor = RGB(111, 112, 116)
    m_ButtonColorOver = RGB(255, 215, 103)
    m_ButtonColorDown = RGB(255, 171, 63)
  Case Black
    m_BaseColor = RGB(199, 203, 209)
    m_BorderColor = RGB(76, 83, 92)
    m_ButtonColorOver = RGB(255, 215, 103)
    m_ButtonColorDown = RGB(255, 171, 63)
  Case Custom
  
  End Select
  SetPics
End Sub

Private Sub SetUpMenu()
  mButts(0).Visible = True
  If m_ButtonCount > 0 Then
    Dim i As Integer
    If mButts.UBound > 1 Then
      For i = 1 To mButts.UBound
        Unload mButts(i)
      Next
    End If
    For i = 1 To m_ButtonCount
      Load mButts(i)
      With mButts(i)
        .Caption = m_Buttons(i).Caption
        .Checked = m_Buttons(i).Visible
        .Enabled = True
        .Visible = True
      End With
    Next
    mButts(0).Visible = False
  Else
  
  End If
  mMore.Enabled = m_ExpandedButtons < VisibleButton And VisibleButton <> 0
  mFewer.Enabled = m_ExpandedButtons > 0 And VisibleButton <> 0
End Sub

'Private Sub WProc_WinProcs(pHwnd As Long, uMSG As Long, wParam As Long, lParam As Long)
'  If pHwnd = UserControl.hwnd Then
'    Select Case uMSG
'    Case WM_MOUSELEAVE
'      Debug.Print "WM_MOUSELEAVE"
'      ResetTip
'      Tmp = -1
'      DrawAll
'    End Select
'  End If
'End Sub
