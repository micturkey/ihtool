VERSION 5.00
Begin VB.Form setd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置错误提示"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "setd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5160
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton OK 
      Caption         =   "确认"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame DIYTip 
      Caption         =   "自定义"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4455
         TabIndex        =   7
         Top             =   240
         Width           =   4455
         Begin VB.OptionButton Q 
            Caption         =   "问号"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   16
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox Tcaption 
            Enabled         =   0   'False
            Height          =   375
            Left            =   840
            TabIndex        =   15
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox Ttitle 
            Enabled         =   0   'False
            Height          =   390
            Left            =   840
            TabIndex        =   13
            Top             =   480
            Width           =   2895
         End
         Begin VB.OptionButton I 
            Caption         =   "通知图标"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2680
            TabIndex        =   11
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton E 
            Caption         =   "警告图标"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1520
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton C 
            Caption         =   "严重错误图标"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Lcaption 
            AutoSize        =   -1  'True
            Caption         =   "内容"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Ltitle 
            AutoSize        =   -1  'True
            Caption         =   "标题"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   360
         End
      End
   End
   Begin VB.Frame ChooseTip 
      Caption         =   "选择错误提示"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   4455
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         Begin VB.OptionButton Option4 
            Caption         =   "自定义"
            Height          =   300
            Left            =   0
            TabIndex        =   6
            Top             =   1200
            Width           =   4455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "找不到文件,请确定文件名是否正确后,再试一次"
            Height          =   300
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   4455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "指定的路径不存在 请检查路径,然后再试一次"
            Height          =   300
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   4455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "文件名\目录名或卷标语法不正确"
            Height          =   300
            Left            =   0
            TabIndex        =   3
            Top             =   840
            Width           =   4455
         End
      End
   End
   Begin VB.Label Ltip2 
      AutoSize        =   -1  'True
      Caption         =   "使用自定义需要将所有程序从压缩包中解压出来."
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   2400
      Width           =   3870
   End
   Begin VB.Label Ltip1 
      AutoSize        =   -1  'True
      Caption         =   "注:除自定义外,其他选项错误提示框标题均为原文件路径."
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   4590
   End
End
Attribute VB_Name = "setd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt


Private Sub OK_Click()
    If Option1.Value = True Then newkey.FilePath2.Text = "error.err"
    If Option2.Value = True Then newkey.FilePath2.Text = "c:\error^1\error.err"
    If Option3.Value = True Then newkey.FilePath2.Text = "c:\error*1\error.err"
    If Option4.Value = True Then
        If C.Value = True Then newkey.FilePath2.Text = App.Path + "\msgshow.exe" + " /16 " + Ttitle.Text + " " + Tcaption.Text
        If E.Value = True Then newkey.FilePath2.Text = App.Path + "\msgshow.exe" + " /48 " + Ttitle.Text + " " + Tcaption.Text
        If I.Value = True Then newkey.FilePath2.Text = App.Path + "\msgshow.exe" + " /64 " + Ttitle.Text + " " + Tcaption.Text
        If Q.Value = True Then newkey.FilePath2.Text = App.Path + "\msgshow.exe" + " /32 " + Ttitle.Text + " " + Tcaption.Text
    End If
    Unload Me
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    x = newkey.FilePath2.Text
    If x = "error.err" Then Option1.Value = True
    If x = "c:\error^1\error.err" Then Option2.Value = True
    If x = "c:\error*1\error.err" Then Option3.Value = True
    If Left(x, Len(App.Path) + 12) = App.Path + "\msgshow.exe" Then
        x = Right(x, Len(x) - Len(App.Path))
        Y = Split(x, " ", -1, vbTextCompare)
        
        opt = Mid(Y(1), 2)
        Ttitle.Text = Y(2)
        Tcaption.Text = Y(3)
        Option4.Value = True
        Debug.Print x, Y(1), Y(2), Y(3)
        If opt = "16" Then C.Value = True
        If opt = "48" Then E.Value = True
        If opt = "64" Then I.Value = True
        If opt = "32" Then Q.Value = True
        C.Enabled = True
        E.Enabled = True
        I.Enabled = True
        Q.Enabled = True
        Ttitle.Enabled = True
        Tcaption.Enabled = True
        DIYTip.Enabled = True
        Ltitle.Enabled = True
        Lcaption.Enabled = True
    End If
    Call ChangeLanguage(Lan, setd)
    
End Sub

Private Sub Option1_Click()
    C.Enabled = False
    E.Enabled = False
    I.Enabled = False
    Q.Enabled = False
    Ttitle.Enabled = False
    Tcaption.Enabled = False
    DIYTip.Enabled = False
    Ltitle.Enabled = False
    Lcaption.Enabled = False
End Sub

Private Sub Option2_Click()
    C.Enabled = False
    E.Enabled = False
    I.Enabled = False
    Q.Enabled = False
    Ttitle.Enabled = False
    Tcaption.Enabled = False
    DIYTip.Enabled = False
    Ltitle.Enabled = False
    Lcaption.Enabled = False
End Sub

Private Sub Option3_Click()
    C.Enabled = False
    E.Enabled = False
    I.Enabled = False
    Q.Enabled = False
    Ttitle.Enabled = False
    Tcaption.Enabled = False
    DIYTip.Enabled = False
    Ltitle.Enabled = False
    Lcaption.Enabled = False
End Sub

Private Sub Option4_Click()
    C.Enabled = True
    E.Enabled = True
    I.Enabled = True
    Q.Enabled = True
    Ttitle.Enabled = True
    Tcaption.Enabled = True
    DIYTip.Enabled = True
    Ltitle.Enabled = True
    Lcaption.Enabled = True
End Sub

