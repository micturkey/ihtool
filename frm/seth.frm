VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form seth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������з���"
   ClientHeight    =   4260
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   5670
   Icon            =   "seth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5670
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox CheckAdmin 
      Caption         =   "ԭ�����Ƿ���Ҫ����ԱȨ��(Only for Win 6.X+)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton CommandOpen 
      Caption         =   "��..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Choose 
      Caption         =   "ѡ������"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton Option2 
            Caption         =   "��ȫ��"
            Height          =   375
            Left            =   2520
            TabIndex        =   3
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "����"
            Height          =   375
            Left            =   600
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
      End
   End
   Begin VB.Label LTip2 
      Caption         =   "�벻Ҫ������Ҫ����ԱȨ�޵�ԭ�ļ�ָ����Ҫ����ԱȨ�޵��ļ���������޷��򿪡�"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   5415
   End
   Begin VB.Label Ltip3 
      AutoSize        =   -1  'True
      Caption         =   "�˴�������գ��Ͳ�����ѡ��ֱ�Ӱ�ȫ��ԭ�ļ������β���Ӱ�졣"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   5400
   End
   Begin VB.Label LTip1 
      Caption         =   "ע��ѡ���ļ��󣬴��κ��ļ���Ϊ���ļ�������������ѡ�ļ���"
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5400
   End
   Begin VB.Label LPath 
      AutoSize        =   -1  'True
      Caption         =   "��ȫ���ļ�·����"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1620
   End
End
Attribute VB_Name = "seth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'���������GPL V3Э��,���������Ŀ¼��gplv3.txt
Private Sub commandOpen_Click()
    On Error GoTo err
    CommonDialog1.Filter = GetString(Lan, "openexe1")
    
    CommonDialog1.DialogTitle = GetString(Lan, "openexe2")
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
    Exit Sub
err: Text1.Text = ""
End Sub


Private Sub OK_Click()
    newkey.Tag = Option1.Value
    newkey.FilePath1.Tag = Text1.Text
    newkey.FilePath2.Tag = CheckAdmin.Value
    If newkey.FilePath2.Tag = "0" Then
        newkey.FilePath2.Text = App.Path + "\saferun.exe" + " /" + newkey.FilePath1.Text
    Else
        newkey.FilePath2.Text = App.Path + "\saferunadmin.exe" + " /" + newkey.FilePath1.Text
    End If
    
    Unload Me
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Debug.Print newkey.Tag
    Call ChangeLanguage(Lan, seth)
    
    Dim x
    x = Split(LTip1.Caption, "/", -1, vbTextCompare)
    LTip1.Caption = x(0) + newkey.FilePath1.Text + x(1)
    If newkey.Tag = "" Then
        Option1.Value = True
    Else
        Option1.Value = newkey.Tag
    End If
    Text1.Text = newkey.FilePath1.Tag
    If Option1.Value = False Then
        Option2.Value = True
        Text1.Enabled = True
        CommandOpen.Enabled = True
        CheckAdmin.Value = newkey.FilePath2.Tag
        CheckAdmin.Enabled = True
    End If
    Debug.Print newkey.FilePath2.Tag
    
End Sub


Private Sub Option1_Click()
    CheckAdmin.Value = 0
    Text1.Enabled = False
    CommandOpen.Enabled = False
    CheckAdmin.Enabled = False
End Sub

Private Sub Option2_Click()
    Text1.Enabled = True
    CommandOpen.Enabled = True
    CheckAdmin.Enabled = True
End Sub
