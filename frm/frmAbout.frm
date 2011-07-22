VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 IHTool"
   ClientHeight    =   3435
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370.898
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   360
      Picture         =   "frmAbout.frx":3F3A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   720
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2865
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   112.686
      X2              =   5183.565
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label Lqg 
      AutoSize        =   -1  'True
      Caption         =   "官方用户QQ群：44034090"
      Height          =   180
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label N1 
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label Lbuild 
      AutoSize        =   -1  'True
      Caption         =   "build 103"
      Height          =   180
      Left            =   2760
      TabIndex        =   11
      Top             =   600
      Width           =   810
   End
   Begin VB.Label Lintro 
      AutoSize        =   -1  'True
      Caption         =   "IHTool - 映像劫持修改工具。"
      Height          =   180
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label Lurl 
      AutoSize        =   -1  'True
      Caption         =   "官方网站："
      Height          =   180
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   900
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "http://www.odets.net/ihtool.htm"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   2520
      MouseIcon       =   "frmAbout.frx":45C6
      TabIndex        =   7
      Top             =   960
      Width           =   2790
   End
   Begin VB.Label Lblog 
      AutoSize        =   -1  'True
      Caption         =   "官方博客："
      Height          =   180
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label blog 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "http://blog.odets.net"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   2520
      MouseIcon       =   "frmAbout.frx":48D0
      TabIndex        =   5
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label bbs 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "http://bbs.odets.net"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   2520
      MouseIcon       =   "frmAbout.frx":4BDA
      TabIndex        =   4
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Label Lbbs 
      AutoSize        =   -1  'True
      Caption         =   "官方论坛："
      Height          =   180
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      Caption         =   "本程序由多重工作室 oDet Studio 编写"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   2865
      Width           =   3150
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cache = Lbuild.Caption
    Call ChangeLanguage(Lan, frmAbout)
    
    lblVersion.Caption = lblVersion.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    Lbuild.Caption = cache
    
End Sub


Private Sub bbs_Click()
    ShellExecute Me.hwnd, "open", "http://bbs.odets.net", "", "", 5
End Sub

Private Sub url_Click()
    ShellExecute Me.hwnd, "open", "http://www.odets.net/ihtool.html", "", "", 5
End Sub

Private Sub blog_Click()
    ShellExecute Me.hwnd, "open", "http://blog.odets.net", "", "", 5
End Sub
