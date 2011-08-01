VERSION 5.00
Begin VB.Form msgshow 
   Caption         =   "Nothing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "msgshow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Caption         =   "Nothing"
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "msgshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Sub Form_Load()
    Lan = language()
    Call ChangeLanguage(Lan, msgshow)
    Me.Hide
    If Command <> "" Then
        X = Command
        Y = Split(X, " ", -1, vbTextCompare)
        opt = Mid(Y(0), 2)
        I = Split(X, "*", -1, vbTextCompare)
        a = MsgBox(I(2), CLng(opt), I(1))
        End
    Else
        Me.Show
    End If
    
End Sub

