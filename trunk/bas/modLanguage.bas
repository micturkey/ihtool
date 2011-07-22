Attribute VB_Name = "modLanguage"
'��л���ƾ��ṩģ��
'�޸Ĳ��� Copyright (C) 2011 oDet Studio
'���������GPL V3Э��,���������Ŀ¼��gplv3.txt


Option Explicit
'API����
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
'Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'#�����ļ�, �����޸�Ϊ����Ҫ�������ļ�����
Private Const LanguageFile = "Language.lng"
'#�޷���ȡ���ȡ����ʱ���ص��ַ���
Private Const UnknowString = "Lost language Pack."
Public Lan As String

'=====================================
'�޸����Ժ���
'=====================================
Public Function GetLanguage() As String()
    On Error Resume Next
    Dim strReturn As String, lenReturn As Long
    strReturn = vbNullString
    If Check Then
        strReturn = Space(&HFE)
        lenReturn = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, strReturn, &HFF, App.Path & Chr(&H5C) & LanguageFile)
        GetLanguage = Split(Trim(Replace(Left(strReturn, lenReturn), Chr(0) & Chr(0), Chr(0))), Chr(0))
    Else
        GetLanguage = Split("None")
    End If
End Function

Public Sub ChangeLanguage(language As String, frm As Form)
    
    Dim I As Long, Ctrl As Control
    
    
    For Each Ctrl In frm.Controls
        Call ChangeLanguageSub(Ctrl, language, frm)
    Next
    
End Sub

Public Sub ChangeLanguageSub(ctlTarget As Control, language As String, frm As Form)
    On Error Resume Next
    'Debug.Print TypeName(ctlTarget) & ":" & ctlTarget.Name & ":" & ctlTarget.hWnd
    If TypeOf ctlTarget Is Frame Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    ElseIf TypeOf ctlTarget Is Label Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    ElseIf TypeOf ctlTarget Is CheckBox Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    ElseIf TypeOf ctlTarget Is OptionButton Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    ElseIf TypeOf ctlTarget Is CommandButton Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    ElseIf TypeOf ctlTarget Is ListView Then
        Dim I
        For I = 1 To 3
            ctlTarget.ColumnHeaders(I).Text = GetString(language, frm.Name & "." & ctlTarget.Name & ".ColumnHeaders(" & I & ")")
        Next I
    ElseIf TypeOf ctlTarget Is Menu Then
        ctlTarget.Caption = GetString(language, frm.Name & "." & ctlTarget.Name)
    End If
    frm.Caption = GetString(language, frm.Name)
    
End Sub

'=====================================
'�����ļ���д����
'=====================================
Private Function Check()
    Check = IIf(Dir(App.Path & Chr(&H5C) & LanguageFile) = LanguageFile, True, False)
End Function

Public Function language() As String
    Dim LocaleID     As Long
    LocaleID = GetSystemDefaultLCID()
    Select Case LocaleID
    Case &H404:     language = "Chinese Tr"
    Case &H804:     language = "Chinese Si"
    Case 1033:     language = "English"
    Case Else:     language = "Else"
    End Select
End Function

Public Function GetString(language As String, Key As String) As String
    On Error Resume Next
    Dim strReturn As String
    Dim lenReturn As Long
    strReturn = vbNullString
    If Check Then
        strReturn = Space(&HFE)
        lenReturn = GetPrivateProfileString(language, Key, Key, strReturn, &HFF, App.Path & Chr(&H5C) & LanguageFile)
        GetString = Left(strReturn, lenReturn)
    Else
        GetString = UnknowString
    End If
End Function
