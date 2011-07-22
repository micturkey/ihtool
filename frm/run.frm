VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form run 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "突破映像劫持运行程序"
   ClientHeight    =   735
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   5730
   Icon            =   "run.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5730
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "运行"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开.."
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "程序路径："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    CommonDialog1.Filter = "可执行文件(*.exe;)|*.exe"
    
    CommonDialog1.DialogTitle = "打开可执行文件(*.exe)"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Then
        On Error GoTo err
        A = MsgBox("请先选择需要打开的文件", vbInformation, "突破映像劫持运行程序")
    Else
        RunExecute (Text1.Text)
    End If
    Exit Sub
err: Text1.Text = ""
End Sub
