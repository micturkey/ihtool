Attribute VB_Name = "hookrun"
'Copyright (C) 2011 oDet Studio
'���������GPL V3Э��,���������Ŀ¼��gplv3.txt
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal ApplicationName As String, ByVal CommandLine As String, ByVal ProcessAttributes As Long, ByVal ThreadAttributes As Long, ByVal InheritHandles As Long, ByVal CreationFlags As Long, ByVal Environment As Long, ByVal CurrentDriectory As String, StartInfo As STARTUPINFO, ProcessInformation As PROCESS_INFORMATION) As Long
'CreateProcess���ڴ���һ���µĽ��̣�����������Ϣ��Ϊ�˷����ҿ���ż���������ͽṹ��
Private Type STARTUPINFO                                                        '����һЩ���ĵ���������
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
Private Type PROCESS_INFORMATION                                                '���ؽ�����Ϣ
    ProcessHandle As Long                                                       '���ر������Ľ��̾��
    ThreadHandle As Long                                                        '���ر������Ľ������߳̾��
    ProcessID As Long                                                           '���ر������Ľ���ID
    ThreadID As Long                                                            '���ر����������߳�ID
End Type
Private Declare Function DbgUiConnectToDbg Lib "ntdll" () As Long               '����DBG
Private Declare Function DbgUiStopDebugging Lib "ntdll" (ByVal hProcess As Long) As Long 'ֹͣ��һ�����̵ĸ��ӵ���
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub RunExecute(ByVal ExeFile As String)                                  '��ͬ:Shell "", vbNormalFocus����ֱ�Ӽ��ϲ�����
    Dim PI As PROCESS_INFORMATION, SI As STARTUPINFO
    DbgUiConnectToDbg
    CreateProcess vbNullString, ExeFile, 0, 0, 0, &H1 Or &H2, 0, vbNullString, SI, PI
    DbgUiStopDebugging PI.ProcessHandle
    CloseHandle PI.ProcessHandle
    CloseHandle PI.ThreadHandle
End Sub
