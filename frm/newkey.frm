VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form newkey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "增加映像劫持"
   ClientHeight    =   4080
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   4740
   Icon            =   "newkey.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4740
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton SetMsg 
      Caption         =   "设置错误提示"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CheckBox UseMsg 
      Caption         =   "弹出错误框(默认为找不到文件)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton SetCon 
      Caption         =   "设置运行方法"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox UseCon 
      Caption         =   "使用IHTool管理(默认是屏蔽)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Add 
      Caption         =   "添加"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Open2 
      Caption         =   "打开..."
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Open1 
      Caption         =   "打开..."
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox FilePath2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox FilePath1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label LB2 
      Caption         =   "Win6.X+下请让原文件和指向文件均为或均不为管理员权限"
      Height          =   420
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   4230
   End
   Begin VB.Label LB1 
      AutoSize        =   -1  'True
      Caption         =   "如果使用IHTool管理，请先把程序从压缩包中解压。"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   4140
   End
   Begin VB.Label Filename2 
      AutoSize        =   -1  'True
      Caption         =   "指向文件名："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Filename1 
      AutoSize        =   -1  'True
      Caption         =   "文件名："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "newkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Const HKEY_LOCAL_MACHINE = &H80000002
Dim F1, F2 As Boolean


Private Sub Usecon_Click()
    If F2 Then Exit Sub
    F1 = True
    UseMsg.Value = 0
    SetMsg.Enabled = False
    F1 = False
    If UseCon.Value = 1 Then
        If FilePath2.Tag = "0" Then
            FilePath2.Text = App.Path + "\saferun.exe" + " /" + FilePath1.Text
        Else
            FilePath2.Text = App.Path + "\saferunadmin.exe" + " /" + FilePath1.Text
        End If
        FilePath2.Enabled = False
        Open2.Enabled = False
        SetCon.Enabled = True
        If Open1.Tag <> "change" Then
            newkey.Tag = "True"
        End If
    Else
        FilePath2.Text = ""
        FilePath2.Enabled = True
        Open2.Enabled = True
        SetCon.Enabled = False
        newkey.Tag = ""
    End If
    
End Sub

Private Sub UseMsg_Click()
    If F1 Then Exit Sub
    F2 = True
    UseCon.Value = 0
    SetCon.Enabled = False
    F2 = False
    If UseMsg.Value = 1 Then
        FilePath2.Text = "error.err"
        FilePath2.Enabled = False
        SetMsg.Enabled = True
        Open2.Enabled = False
    Else
        FilePath2.Text = ""
        FilePath2.Enabled = True
        SetMsg.Enabled = False
        Open2.Enabled = True
    End If
End Sub

Private Sub Open1_Click()
    On Error GoTo err
    CommonDialog1.Filter = GetString(Lan, "openexe1")
    
    CommonDialog1.DialogTitle = GetString(Lan, "openexe2")
    CommonDialog1.ShowOpen
    
    FilePath1.Text = CommonDialog1.FileTitle
    CommonDialog1.Tag = CommonDialog1.FileName
    Exit Sub
err: FilePath1.Text = ""
End Sub

Private Sub Open2_Click()
    On Error GoTo err
    CommonDialog1.Filter = GetString(Lan, "openexe1")
    
    CommonDialog1.DialogTitle = GetString(Lan, "openexe2")
    CommonDialog1.ShowOpen
    FilePath2.Text = CommonDialog1.FileName
    If FilePath2.Text = CommonDialog1.Tag Then
        a = MsgBox(GetString(Lan, "selectsame1"), vbInformation, GetString(Lan, "selectsame2"))
        FilePath2.Text = ""
    End If
    Exit Sub
err: FilePath2.Text = ""
End Sub



Private Sub Add_Click()
    On Error GoTo error
    If FilePath1.Text = "" Or FilePath2.Text = "" Then
        a = MsgBox(GetString(Lan, "selectfirst"), vbInformation, GetString(Lan, "newkey"))
    Else
        If Right(FilePath1.Text, 3) = "exe" Or Right(FilePath1.Text, 3) = "EXE" Or Right(FilePath2.Text, 3) = "exe" Or Right(FilePath2.Text, 3) = "EXE" Then
            If FilePath1.Text = "ihtool.exe" Or FilePath1.Text = "saferun.exe" Or FilePath1.Text = "saferunadmin.exe" Or FilePath1.Text = "msgshow.exe" Then
                a = MsgBox(GetString(Lan, "notselectihtool"), vbInformation, GetString(Lan, "newkey"))
            Else
                If FilePath2.Text <> App.Path + "\saferun.exe" Then
                    ret = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & FilePath1.Text, lphKey)
                    Set reg = CreateObject("wscript.shell")
                    a = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & FilePath1.Text & "\debugger", FilePath2.Text, "REG_SZ")
                    a = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & FilePath1.Text & "\ihtool", "safe", "REG_SZ")
                    If newkey.Tag = "False" Then
                        a = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & FilePath1.Text & "\saferun", FilePath1.Tag, "REG_SZ")
                    Else
                        Call Delreg
                    End If
                    ihtool.scan_Click
                    Unload Me
                Else
                    a = MsgBox(GetString(Lan, "notselectsaferun1") + Chr(13) + GetString(Lan, "notselectsaferun2"), vbInformation, GetString(Lan, "newkey"))
                End If
            End If
        Else
            a = MsgBox(GetString(Lan, "selectagain"), vbInformation, GetString(Lan, "newkey"))
        End If
    End If
    
    Exit Sub
error:  a = MsgBox(GetString(Lan, "changeerr"), vbExclamation, "IHTool")
    
End Sub
Private Sub Delreg()
    On Error Resume Next
    Set reg = CreateObject("wscript.shell")
    a = reg.regdelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & FilePath1.Text & "\saferun")
End Sub
Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub SetCon_Click()
    seth.Show
End Sub

Private Sub SetMsg_Click()
    setd.Show
End Sub

Private Sub Form_Load()
    newkey.Tag = ""
    FilePath2.Tag = "0"
    Call ChangeLanguage(Lan, newkey)
    
End Sub

Private Sub FilePath1_Change()
    
    If FilePath1.Text <> "" Then
        UseCon.Enabled = True
    Else
        UseCon.Enabled = False
    End If
    If UseCon.Value = 1 Then
        
        If FilePath2.Tag = "0" Then
            FilePath2.Text = App.Path + "\saferun.exe" + " /" + FilePath1.Text
        Else
            FilePath2.Text = App.Path + "\saferunadmin.exe" + " /" + FilePath1.Text
        End If
    End If
    If FilePath1.Text <> "" Then
        UseMsg.Enabled = True
    Else
        UseMsg.Enabled = False
    End If
    If UseMsg.Value = 1 Then
        FilePath2.Text = "c:\error^1\error.exe"
    End If
End Sub

