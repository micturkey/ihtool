Attribute VB_Name = "hookrun"
'Copyright (C) 2011 oDet Studio
'本程序遵从GPL V3协议,详情请见根目录下gplv3.txt
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal ApplicationName As String, ByVal CommandLine As String, ByVal ProcessAttributes As Long, ByVal ThreadAttributes As Long, ByVal InheritHandles As Long, ByVal CreationFlags As Long, ByVal Environment As Long, ByVal CurrentDriectory As String, StartInfo As STARTUPINFO, ProcessInformation As PROCESS_INFORMATION) As Long
'CreateProcess用于创建一个新的进程，并返回其信息。为了方便大家看，偶改了声明和结构。
Private Type STARTUPINFO                                                        '关于一些无聊的启动设置
    CB As Long
    Reserved As String
    Desktop As String
    Title As String
    x As Long
    Y As Long
    XSize As Long
    YSize As Long
    XCountChars As Long
    YCountChars As Long
    FillAttribute As Long
    Flags As Long
    ShowWindow As Integer
    IntReserved As Integer
    LngReserved As Long
    StdInput As Long
    StdOutput As Long
    StdError As Long
End Type
Private Type PROCESS_INFORMATION                                                '返回进程信息
    ProcessHandle As Long                                                       '返回被创建的进程句柄
    ThreadHandle As Long                                                        '返回被创建的进程主线程句柄
    ProcessID As Long                                                           '返回被创建的进程ID
    ThreadID As Long                                                            '返回被创建的主线程ID
End Type
Private Declare Function DbgUiConnectToDbg Lib "ntdll" () As Long               '连接DBG
Private Declare Function DbgUiStopDebugging Lib "ntdll" (ByVal hProcess As Long) As Long '停止对一个进程的附加调试
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub RunExecute(ByVal ExeFile As String)                                  '等同:Shell "", vbNormalFocus，可直接加上参数。
    Dim PI As PROCESS_INFORMATION, SI As STARTUPINFO
    DbgUiConnectToDbg
    CreateProcess vbNullString, ExeFile, 0, 0, 0, &H1 Or &H2, 0, vbNullString, SI, PI
    DbgUiStopDebugging PI.ProcessHandle
    CloseHandle PI.ProcessHandle
    CloseHandle PI.ThreadHandle
End Sub
