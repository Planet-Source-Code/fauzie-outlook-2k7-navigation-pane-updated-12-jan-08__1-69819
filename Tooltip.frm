VERSION 5.00
Begin VB.Form frmTooltip 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   2955
   ClientTop       =   3195
   ClientWidth     =   1560
   ControlBox      =   0   'False
   Icon            =   "Tooltip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTip 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   -45
   End
   Begin VB.Label lblTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "TipLabel "
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   660
   End
End
Attribute VB_Name = "frmTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Module     : frmTooltip
'Description: ToolTip Form
'Version    : 1.00 Dec 2007
'Release    : VB6
'Copyright  : © 2007 by Fauzie's Software. All rights reserved
'E-Mail     : fauzie811@yahoo.com
'--------------------------------------------------------------------

Option Explicit
DefInt A-Z

Dim clr() As Long

Const m_BorderColor = &H767676

Public CtlHWnd As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Sub Form_Load()
 AutoRedraw = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Unload Me
End Sub

Private Sub Form_Resize()
 Dim i, j As OLE_COLOR
 Cls
 ScaleMode = 3
 BlendColors TranslateColor(vbWhite), TranslateColor(RGB(203, 218, 239)), ScaleHeight + 1, clr
 For i = 0 To ScaleHeight
  Line (0, i)-(ScaleWidth, i), clr(i)
 Next
 Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor, B

 PSet (0, 0), vbMagenta
 PSet (ScaleWidth - 1, 0), vbMagenta
 PSet (0, ScaleHeight - 1), vbMagenta
 PSet (ScaleWidth - 1, ScaleHeight - 1), vbMagenta
 
 j = GetLightColor(TranslateColor(m_BorderColor), 60)
 PSet (1, 1), j
 PSet (ScaleWidth - 2, 1), j
 PSet (1, ScaleHeight - 2), j
 PSet (ScaleWidth - 2, ScaleHeight - 2), j
 
 SetRgn
 ScaleMode = 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
 tmrTip.Enabled = 0
End Sub

Private Sub lblTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Unload Me
End Sub

Private Sub tmrTip_Timer()
 If GetActiveWindow() <> CtlHWnd Then Unload Me
End Sub

Private Sub SetRgn()
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    X = ScaleWidth
    Y = ScaleHeight
    
    Sum = CreateRectRgn(1, 0, X - 1, 1)
    CombineRgn Sum, Sum, CreateRectRgn(0, 1, X, Y - 1), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, Y, X - 1, Y - 1), 2
    SetWindowRgn hWnd, Sum, True   'Sets corners transparent
End Sub

