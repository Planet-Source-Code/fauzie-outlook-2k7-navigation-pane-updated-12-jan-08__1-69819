VERSION 5.00
Object = "*\ANavPane.vbp"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Navigation Pane Testing Project"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6255
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   3  'Windows Default
   Begin NavigationPane.NavPane NavPane1 
      Height          =   7215
      Left            =   240
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   12726
      BorderColor     =   -2147483632
      ForeColor       =   9126421
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionForeColor=   9126421
      ButtonCount     =   0
   End
   Begin VB.PictureBox Ico1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3840
      Picture         =   "fTest.frx":1CFA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4200
      Picture         =   "fTest.frx":2284
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4560
      Picture         =   "fTest.frx":280E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4920
      Picture         =   "fTest.frx":2D98
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5280
      Picture         =   "fTest.frx":3322
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5640
      Picture         =   "fTest.frx":38AC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ico0 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "fTest.frx":3E36
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rename Current Button"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtButtonCount 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox txtExpandedButton 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtActiveButton 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "fTest.frx":43C0
      Left            =   4080
      List            =   "fTest.frx":43D0
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Ico0 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3480
      Picture         =   "fTest.frx":4407
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5640
      Picture         =   "fTest.frx":4AF1
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5280
      Picture         =   "fTest.frx":51DB
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4920
      Picture         =   "fTest.frx":58C5
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4560
      Picture         =   "fTest.frx":5FAF
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4200
      Picture         =   "fTest.frx":6699
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Ico1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3840
      Picture         =   "fTest.frx":6D83
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Current Button"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Button"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Button Count:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblExpandedButton 
      Caption         =   "Expanded Button:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblActiveButton 
      Caption         =   "Active Button:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Scheme:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub Combo1_Change()
  NavPane1.Theme = Combo1.ListIndex
End Sub

Private Sub Combo1_Click()
  NavPane1.Theme = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
  i = i + 1
  NavPane1.ActiveButton = NavPane1.AddButtonex("Button " & i)
End Sub

Private Sub Command2_Click()
  NavPane1.AboutBox
End Sub

Private Sub Command3_Click()
   If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then NavPane1.RemoveButton NavPane1.ActiveButton
End Sub

Private Sub Command4_Click()
  NavPane1.ButtonCaption(NavPane1.ActiveButton) = InputBox("Enter name :", , NavPane1.ButtonCaption(NavPane1.ActiveButton))
End Sub

Private Sub Form_Load()
  Combo1.ListIndex = NavPane1.Theme
  txtExpandedButton = NavPane1.ExpandedButtons
  txtActiveButton = 1
  NavPane1.AddButtonex "Mail", , , , , , Ico0(0).Picture, Ico0(1).Picture
  NavPane1.AddButtonex "Calendar", , , , , , Ico1(0).Picture, Ico1(1).Picture
  NavPane1.AddButtonex "Contacts", , , , , , Ico2(0).Picture, Ico2(1).Picture
  NavPane1.AddButtonex "Tasks", , , , , , Ico3(0).Picture, Ico3(1).Picture
  NavPane1.AddButtonex "Notes", , , , , , Ico4(0).Picture, Ico4(1).Picture
  NavPane1.AddButtonex "Folder List", , , , , , Ico5(0).Picture, Ico5(1).Picture
  NavPane1.AddButtonex "Shortcuts", , , , , , Ico6(0).Picture, Ico6(1).Picture
  NavPane1.ActiveButton = 1
End Sub

Private Sub NavPane1_ButtonChanged(Button As NavigationPane.ButtonItem)
  txtActiveButton = NavPane1.ActiveButton
End Sub

Private Sub NavPane1_ButtonCountChanged(NewCount As Integer)
  txtButtonCount = NewCount
End Sub

Private Sub NavPane1_Resize()
  txtExpandedButton = NavPane1.ExpandedButtons
End Sub
