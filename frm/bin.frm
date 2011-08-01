VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form bin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "回收站"
   ClientHeight    =   6030
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   9615
   Icon            =   "bin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9615
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox Checkall 
      Caption         =   "全选"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton fresh 
      Caption         =   "刷新"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton exit 
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton delete 
      Caption         =   "彻底删除"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton reback 
      Caption         =   "还原"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   5400
      Width           =   1815
   End
   Begin MSComctlLib.ListView List 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
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
         Text            =   "删除时间"
         Object.Width           =   3704
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
End
Attribute VB_Name = "bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_LOCAL_MACHINE = &H80000002
Dim hR As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Const ERROR_SUCCESS = 0&
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

Private Sub checkall_Click()
    'On Error Resume Next
    If Checkall.Value = 1 Then
        For I = List.ListItems.Count To 1 Step -1
            List.ListItems(I).Checked = True
            If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
                delete.Enabled = True
                reback.Enabled = True
            End If
        Next
    Else
        For I = List.ListItems.Count To 1 Step -1
            List.ListItems(I).Checked = False
            delete.Enabled = False
            reback.Enabled = False
        Next
    End If
End Sub

Private Sub delete_Click()
    Dim ret As Long, hKey As Long
    If MsgBox(GetString(Lan, "delete?"), vbOKCancel + vbInformation, GetString(Lan, "delete")) = vbOK Then
        For I = List.ListItems.Count To 1 Step -1
            On Error Resume Next
            If List.ListItems(I).Checked = True Then
                
                Set reg = CreateObject("wscript.shell")
                RegOpenKey HKEY_LOCAL_MACHINE, ("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1)), hR '这里是项名
                If RegQueryValueEx(hR, "debugger", 0, 0, 0, 0) = 2 Then
                    On Error GoTo error
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtdel")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtbin")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtpath")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\")
                    
                Else
                    On Error GoTo error
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtdel")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtbin")
                    reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtpath")
                    
                    
                End If
            End If
        Next I
        Checkall.Value = 0
        fresh_Click
    End If
    Exit Sub
error:  a = MsgBox(GetString(Lan, "delerr"), vbExclamation, "IHTool")
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call ChangeLanguage(Lan, bin)
    
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
                On Error Resume Next
                Dim cc As String
                cc = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\ihtdel")
                btime = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\ihtbin")
                If err Then
                Else
                    I = I + 1
                    
                    
                    Set itm = List.ListItems.Add(, , btime)
                    itm.ListSubItems.Add , , (Left$(sName, ret - LengthAmendment(sName)))
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
    delete.Enabled = False
    reback.Enabled = False
    If STR(I) = 0 Then
        List.ListItems.Clear
        Set itm = List.ListItems.Add(, , "")
        itm.ListSubItems.Add , , (GetString(Lan, "nofound"))
        Checkall.Enabled = False
    End If
End Sub

Private Sub fresh_Click()
    Checkall.Enabled = True
    Checkall.Value = 0
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
                On Error Resume Next
                Dim cc As String
                cc = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\ihtdel")
                btime = CreateObject("WScript.Shell").regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & C & "\ihtbin")
                If err Then
                Else
                    I = I + 1
                    
                    
                    Set itm = List.ListItems.Add(, , btime)
                    itm.ListSubItems.Add , , (Left$(sName, ret - LengthAmendment(sName)))
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
    delete.Enabled = False
    reback.Enabled = False
    If STR(I) = 0 Then
        List.ListItems.Clear
        Set itm = List.ListItems.Add(, , "")
        itm.ListSubItems.Add , , (GetString(Lan, "nofound"))
        Checkall.Enabled = False
    End If
End Sub
Private Sub list_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Checkall.Value = 0
    Dim itm As ListItem
    Dim I As Integer
    
    For Each itm In List.ListItems
        If itm.Checked = True Then I = I + 1
    Next
    
    If I > 1 Then
        reback.Enabled = True
        delete.Enabled = True
    ElseIf I = 1 Then
        If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
            reback.Enabled = True
            delete.Enabled = True
        End If
    ElseIf I = 0 Then
        delete.Enabled = False
        reback.Enabled = False
    End If
    
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
        reback.Enabled = True
        delete.Enabled = True
    ElseIf I = 1 Then
        If List.ListItems(I).SubItems(1) <> Left(GetString(Lan, "nofound"), Len(List.ListItems(I).SubItems(1))) Then
            reback.Enabled = True
            delete.Enabled = True
        End If
    ElseIf I = 0 Then
        delete.Enabled = False
        reback.Enabled = False
    End If
End Sub

Private Sub reback_Click()
    On Error GoTo error
    
    For I = List.ListItems.Count To 1 Step -1
        If List.ListItems(I).Checked = True Then
            Set reg = CreateObject("wscript.shell")
            If List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto1"), Len(List.ListItems(I).SubItems(2))) Then
                List.ListItems(I).SubItems(2) = App.Path + "\saferun.exe" + " /" + List.ListItems(I).SubItems(1)
            ElseIf List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto2"), Len(List.ListItems(I).SubItems(2))) Then
                List.ListItems(I).SubItems(2) = App.Path + "\saferunadmin.exe" + " /" + List.ListItems(I).SubItems(1)
            ElseIf List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto3"), Len(List.ListItems(I).SubItems(2))) Or List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto4"), Len(List.ListItems(I).SubItems(2))) Or List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto5"), Len(List.ListItems(I).SubItems(2))) Or List.ListItems(I).SubItems(2) = Left(GetString(Lan, "pointto6"), Len(List.ListItems(I).SubItems(2))) Then
                List.ListItems(I).SubItems(2) = reg.regread("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtpath")
                
            End If
            b = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\debugger", List.ListItems(I).SubItems(2), "REG_SZ")
            C = reg.regwrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtool", "safe", "REG_SZ")
            reg.regdelete ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & List.ListItems(I).SubItems(1) & "\ihtdel")
        End If
    Next I
    fresh_Click
    ihtool.scan_Click
    Me.SetFocus
    Exit Sub
error:  a = MsgBox("还原错误，请确定是否以管理员身份打开！" & Chr(13) & "并请确认是否被安全软件拦截！", vbExclamation, "IHTool")
    
End Sub
