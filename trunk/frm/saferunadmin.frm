VERSION 5.00
Begin VB.Form saferunadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ȫ�򿪣�����ԱȨ�ް棩"
   ClientHeight    =   2130
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   4890
   Icon            =   "saferunadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4890
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton open2 
      Caption         =   "��ȫ��ԭ�ļ�"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton open3 
      Caption         =   "�ر�"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton open1 
      Caption         =   "��ȫ��Ŀ���ļ�"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "����ⲻ�����Ĳ����������������Ƿ��в�����"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label label3 
      AutoSize        =   -1  'True
      Caption         =   "���й����п�����Ҫ�ð�ȫ�������"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   "�޲���"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "saferunadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txt As String
Dim read As String
Dim runtxt As String
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub open1_Click()
    RunExecute (read)
    Unload Me
End Sub

Private Sub open3_Click()
    Unload Me
End Sub

Private Sub open2_Click()
    RunExecute (runtxt)
    Unload Me
End Sub

'Copyright (C) 2011 oDet Studio
'���������GPL V3Э��,���������Ŀ¼��gplv3.txt
Private Sub form_Initialize()
    InitCommonControls
    
End Sub

Private Sub Form_Load()
    On Error GoTo error
    Lan = language()
    a = Split(GetString(Lan, "saferunadmin.no"), "/", -1, vbTextCompare)
    Call ChangeLanguage(Lan, saferunadmin)
    If Command <> "" Then
        x = Command
        Y = Split(x, " ", -1, vbTextCompare)
        txt = Mid(Y(0), 2)
        runtxt = Mid(x, Len(txt) + 3)
        read = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & txt & "\saferun")
        open1.Enabled = True
        open2.Enabled = True
        z = Split(Label1.Caption, "/", -1, vbTextCompare)
        Label1.Caption = z(0) & runtxt & z(1)
        If read = "" Then
            Me.Hide
            Call open2_Click
        End If
    Else
        Label1.Caption = GetString(Lan, "saferunadmin.nothing")
    End If
    Exit Sub
error:  Label1.Caption = a(0) & runtxt & a(1)
End Sub


