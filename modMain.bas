Attribute VB_Name = "modMain"
Option Explicit

Public Const APP_NAME As String = "ResourceBrowser"
Public Const vbOrange = &H80FF&

Public Sub Main()
    If App.PrevInstance Then Exit Sub

    frmMain.Show
End Sub


Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    InIDE = (Err > 0)
    On Error GoTo 0
End Function


'save on typing
Function GetStr(ByVal KeyName As String, Optional ByVal DefaultVal As String) As String

    GetStr = GetSetting(APP_NAME, "Settings", KeyName, DefaultVal)
    'in some cases vb doesnt apply the default value on a "" value, we will force it
    If GetStr = "" Then GetStr = DefaultVal
End Function

Function GetInt(ByVal KeyName As String, Optional ByVal DefaultVal As Long) As Long
    GetInt = CLng(GetSetting(APP_NAME, "Settings", KeyName, CStr(DefaultVal)))
End Function

Sub SaveStr(ByVal KeyName As String, ByVal Value As String)
    SaveSetting APP_NAME, "Settings", KeyName, Value
End Sub

Sub SaveInt(ByVal KeyName As String, ByVal Value As Long)
    SaveSetting APP_NAME, "Settings", KeyName, CStr(Value)
End Sub
