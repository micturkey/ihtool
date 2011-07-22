VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ihtool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IHTool"
   ClientHeight    =   6030
   ClientLeft      =   75
   ClientTop       =   690
   ClientWidth     =   9960
   Icon            =   "ihtool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9960
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   8
      Top             =   5520
      Width           =   2415
      Begin VB.OptionButton backup 
         Caption         =   "保留"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton nbackup 
         Caption         =   "不保留"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Qbak 
      Caption         =   "删除是否保留备份？"
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CheckBox Checkall 
      Caption         =   "全选"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton change 
      Caption         =   "更改"
      Height          =   780
      Left            =   7560
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton del 
      Caption         =   "删除"
      Height          =   780
      Left            =   7560
      TabIndex        =   4
      Top             =   3300
      Width           =   2295
   End
   Begin VB.CommandButton new 
      Caption         =   "添加"
      Height          =   780
      Left            =   7560
      TabIndex        =   3
      Top             =   1260
      Width           =   2295
   End
   Begin MSComctlLib.ListView List 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "类型"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "文件名"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "指向文件"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton exit 
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Height          =   780
      Left            =   7560
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton scan 
      Caption         =   "扫描"
      Default         =   -1  'True
      Height          =   780
      Left            =   7560
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label N2 
      AutoSize        =   -1  'True
      Caption         =   "初次使用或无法扫描请点击菜单：映像劫持--增加注册表权限。"
      Height          =   180
      Left            =   4320
      TabIndex        =   12
      Top             =   5640
      Width           =   5040
   End
   Begin VB.Label N1 
      AutoSize        =   -1  'True
      Caption         =   "添加或删除时如果重名，将会替换原有项目。"
      Height          =   180
      Left            =   4320
      TabIndex        =   11
      Top             =   5400
      Width           =   3600
   End
   Begin VB.Menu IFEO 
      Caption         =   "映像劫持(&A)"
      Begin VB.Menu mscan 
         Caption         =   "扫描"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnew 
         Caption         =   "添加"
         Shortcut        =   ^N
      End
      Begin VB.Menu g1 
         Caption         =   "-"
      End
      Begin VB.Menu addregauth 
         Caption         =   "增加注册表权限"
         Shortcut        =   ^A
      End
      Begin VB.Menu mrun 
         Caption         =   "突破映像劫持"
         Shortcut        =   ^R
      End
      Begin VB.Menu g2 
         Caption         =   "-"
      End
      Begin VB.Menu mbin 
         Caption         =   "回收站"
         Shortcut        =   ^B
      End
      Begin VB.Menu g3 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu about 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "ihtool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_LOCAL_MACHINE = &H80000002


Const ERROR_SUCCESS = 0&
Private Declare Sub InitCommonControls Lib "comctl32" ()

Dim listindex As Long
Dim listtext As String
Dim backtimer As Long

'Dim back1 As String
'Dim back2 As String
'Dim back3 As String


Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub addregauth_Click()
    If MsgBox(GetString(Lan, "authority"), vbOKCancel + vbInformation, GetString(Lan, "authtitle")) = vbOK Then
        Open App.Path & "\authority.txt" For Output As #1
        Print #1, "\Registry\Machine\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options [1 17]"
        Close #1
        If Shell("cmd.exe /c regini """ + App.Path + """\authority.txt") <> 0 Then
            msg = MsgBox("成功/Successed", vbInformation + vbOKOnly, GetString(Lan, "authtitle"))
        End If
        Kill (App.Path + "\authority.txt")
    End If
End Sub

Private Sub change_Click()
    On Error GoTo error
    newkey.Show
    For I = List.ListItems.Count To 1 Step -1
        If List.ListItems(I).Checked = True Then
            newkey.FilePath1.Text = List.ListItems(I).SubItems(1)
            newkey.FilePath2.Text = List.ListItems(I).SubItems(2)
            newkey.FilePath1.Enabled = False
            newkey.Open1.Enabled = False
            newkey.Add.Caption = GetString(Lan, "change")
            newkey.Caption = GetString(Lan, "changetitle")
            newkey.Open1.Tag = "change"
            If newkey.FilePath2.Text = Left(GetString(Lan, "pointto1"), Len(newkey.FilePath2.Text)) Then
                'If MsgBox("ae", vbOKCancel, "a") = vbOK Then
                'MsgBox GetString(Lan, "pointto1")
                'newkey.FilePath2.Text = App.Path + "\saferun.exe" + " /" + newkey.FilePath1.Text
                Set reg = CreateObject("wscript.shell")
                newkey.FilePath2.Text = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & newkey.FilePath1.Text & "\debugger")
                newkey.UseCon.Value = 1
                newkey.FilePath1.Tag = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & newkey.FilePath1.Text & "\saferun")
                newkey.Tag = "False"
                newkey.FilePath2.Tag = "0"
            ElseIf newkey.FilePath2.Text = Left(GetString(Lan, "pointto2"), Len(newkey.FilePath2.Text)) Then
                'newkey.FilePath2.Text = App.Path + "\saferunadmin.exe" + " /" + newkey.FilePath1.Text
                Set reg = CreateObject("wscript.shell")
                newkey.FilePath2.Text = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & newkey.FilePath1.Text & "\debugger")
                newkey.UseCon.Value = 1
                newkey.FilePath1.Tag = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & newkey.FilePath1.Text & "\saferun")
                newkey.Tag = "False"
                newkey.FilePath2.Tag = "1"
            ElseIf newkey.FilePath2.Text = Left(GetString(Lan, "pointto3"), Len(newkey.FilePath2.Text)) Or newkey.FilePath2.Text = Left(GetString(Lan, "pointto4"), Len(newkey.FilePath2.Text)) Or newkey.FilePath2.Text = Left(GetString(Lan, "pointto5"), Len(newkey.FilePath2.Text)) Or newkey.FilePath2.Text = Left(GetString(Lan, "pointto6"), Len(newkey.FilePath2.Text)) Then
                Set reg = CreateObject("wscript.shell")
                newkey.UseMsg.Value = 1
                newkey.FilePath2.Text = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & newkey.FilePath1.Text & "\debugger")
            End If
        End If
    Next I
    
    
    Debug.Print newkey.Tag
    Exit Sub
error: newkey.Tag = "True"
End Sub

Private Sub checkall_Click()
    change.Enabled = False
    'On Error Resume Next
    If Checkall.Value = 1 Then
        For I = List.ListItems.Count To 1 Step -1
            List.ListItems(I).Checked = True
        Next
        del.Enabled = True
        For I = List.ListItems.Count To 1 Step -1
            If List.ListItems(I).Checked = True Then
                If I = 0 Then
                    change.Enabled = True
                    
                End If
            End If
        Next I
    Else
        For I = List.ListItems.Count To 1 Step -1
            List.ListItems(I).Checked = False
            change.Enabled = False
            del.Enabled = False
        Next
    End If
    
End Sub

Private Sub form_Initialize()
    InitCommonControls
    
End Sub

Private Sub Form_Load()
    'backtime = 0
    Lan = language()
    Checkall.Enabled = False
    change.Enabled = False
    del.Enabled = False
    If App.PrevInstance Then
        a = MsgBox(GetString(Lan, "running"), vbInformation, "IHTool")
        End                                                                     '   退出新运行的程序
    End If
    
    Call ChangeLanguage(Lan, ihtool)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    For Each frm In Forms
        Unload frm
    Next frm
    End
End Sub

Private Sub list_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Checkall.Value = 0
    Dim itm As ListItem
    Dim I As Integer
    
    For Each itm In List.ListItems
        If itm.Checked = True Then I = I + 1
    Next
    
    
    If I > 1 Then
        change.Enabled = False
        del.Enabled = True
    ElseIf I = 1 Then
        If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
            change.Enabled = True
            del.Enabled = True
        End If
    ElseIf I = 0 Then
        del.Enabled = False
        change.Enabled = False
    End If
End Sub

Private Function LengthAmendment(STR As String) As Integer                      '字符串长度修正
    Dim a As String, r$, I%
    For I = 1 To Len(STR)
        a = Mid(STR, I, 1)
        If Asc(a) < 0 Then
            r = r + a
        End If
    Next
    LengthAmendment = Len(r)
End Function

Private Sub list_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lst As ListItem
    On Error Resume Next
    Set lst = List.HitTest(0, Y)
    lst.Checked = Not lst.Checked
    Checkall.Value = 0
    Dim itm As ListItem
    Dim I As Integer
    
    For Each itm In List.ListItems
        If itm.Checked = True Then I = I + 1
    Next
    
    
    If I > 1 Then
        change.Enabled = False
        del.Enabled = True
    ElseIf I = 1 Then
        If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
            change.Enabled = True
            del.Enabled = True
        End If
    ElseIf I = 0 Then
        del.Enabled = False
        change.Enabled = False
    End If
End Sub

Private Sub mbin_Click()
    bin.Show
End Sub



Private Sub mexit_Click()
    exit_Click
End Sub

Private Sub mnew_Click()
    new_Click
End Sub

Private Sub mrun_Click()
    run.Show
End Sub

Private Sub mscan_Click()
    scan_Click
End Sub


Public Sub scan_Click()
    On Error GoTo error
    Dim I As Integer
    Dim hKey As Long, Cnt As Long, sName As String, sData As String, ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    
    ret = BUFFER_SIZE
    List.ListItems.Clear
    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options", hKey) = 0 Then
        sName = Space(BUFFER_SIZE)
        While RegEnumKeyEx(hKey, Cnt, sName, ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            C = (Left$(sName, ret - LengthAmendment(sName)))
            If Right(C, 3) = "exe" Or Right(C, 3) = "com" Then
                
                del.Enabled = True
                Dim cc As String
                On Error Resume Next
                cc = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\debugger")
                If err Then
                Else
                    I = I + 1
                    'Set itm = list.ListItems.Add(, , (Left$(sName, Ret - LengthAmendment(sName))))
                    safe = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\ihtool")
                    stype = GetString(Lan, "nomade")
                    On Error GoTo error
                    If safe = "safe" Then
                        stype = GetString(Lan, "madeby")
                    End If
                    
                    Set itm = List.ListItems.Add(, , stype)
                    itm.ListSubItems.Add , , (Left$(sName, ret - LengthAmendment(sName)))
                    If cc = App.Path + "\saferun.exe /" + Left$(sName, ret - LengthAmendment(sName)) Then
                        cc = GetString(Lan, "pointto1")
                    ElseIf cc = App.Path + "\saferunadmin.exe /" + Left$(sName, ret - LengthAmendment(sName)) Then
                        cc = GetString(Lan, "pointto2")
                    ElseIf Left(cc, Len(App.Path) + Len("\msgshow.exe")) = App.Path + "\msgshow.exe" Then
                        cc = GetString(Lan, "pointto3")
                    ElseIf cc = "error.err" Then
                        cc = GetString(Lan, "pointto4")
                    ElseIf cc = "c:\error^1\error.err" Then
                        cc = GetString(Lan, "pointto5")
                    ElseIf cc = "c:\error*1\error.err" Then
                        cc = GetString(Lan, "pointto6")
                    End If
                    itm.ListSubItems.Add , , cc
                End If
            End If
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
            
            
            
            
        Wend
        RegCloseKey hKey
    Else
        MsgBox "Read error!"
    End If
    Checkall.Enabled = True
    If STR(I) = 0 Then
        List.ListItems.Clear
        Set itm = List.ListItems.Add(, , "")
        itm.ListSubItems.Add , , (GetString(Lan, "nofound"))
        del.Enabled = False
        change.Enabled = False
        Checkall.Enabled = False
    End If
    change.Enabled = False
    del.Enabled = False
    Exit Sub
error: C = MsgBox(GetString(Lan, "geterr"), vbExclamation, "IHTool")
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub new_Click()
    newkey.Show
End Sub

Private Sub del_Click()
    On Error GoTo error
    If MsgBox(GetString(Lan, "delete?"), vbOKCancel + vbInformation, GetString(Lan, "delete")) = vbOK Then
        
        For I = List.ListItems.Count To 1 Step -1
            If List.ListItems(I).Checked = True Then
                If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
                    
                    Set reg = CreateObject("wscript.shell")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\debugger")
                    
                    If backup.Value = True Then
                        a = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtdel", List.ListItems(I).SubItems(2), "REG_SZ")
                        a = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtbin", Now, "REG_SZ")
                    Else
                        
                        reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\")
                    End If
                    
                    List.ListItems.Remove (I)
                    ' back3 = list.SelectedItem.SubItems(1)
                    
                End If
            End If
        Next I
        Checkall.Value = 0
        scan_Click
    End If
    Exit Sub
error: C = MsgBox(GetString(Lan, "delerr"), vbExclamation, "IHTool")
End Sub



Private Sub list_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If List.SortOrder = lvwAscending Then
        List.SortOrder = lvwDescending
    Else
        List.SortOrder = lvwAscending
    End If
    List.Sorted = True
    List.SortKey = ColumnHeader.Index - 1
    List.Sorted = False
End Sub


