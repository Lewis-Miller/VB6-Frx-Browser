Attribute VB_Name = "CommDlg"
Option Explicit

Public Const MAX_PATH As Long = 260


'the following api calls and structures are used within this module
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (opfn As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (opfn As OPENFILENAME) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCusFilt(7) As Byte
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    lpFlags As Long
    nFiller(3) As Byte
    lpstrDefExt As String
    lCustData(11) As Byte
End Type

Private strCurrentDirectory As String


'the next 3 Show#### functions negate the need in most cases
'to include the bastard child "comdlg32.ocx" with your application.
'function Arguments are all optional and should be self explanatory...

'ShowOpen() 'shows the open file dialog window
'ShowSave() 'shows the save file dialog window
'ShowColor() 'shows the choose color window
'showfont() and showprint() dialog not yet done...

'function relies on GetOpenFileName() api call and OPENFILENAME structure
Function ShowOpen(Optional ByVal strFilter As String, Optional ByVal strDefaultPath As String, Optional ByVal strTitle As String, Optional ByVal strDefaultExt As String, Optional ByVal strInitDir As String) As String

    ShowOpen = ShowDialogHelper(False, strFilter, strDefaultPath, strTitle, strDefaultExt, strInitDir)

End Function

'function relies on GetSaveFileName() api call and OPENFILENAME structure
Function ShowSave(Optional ByVal strFilter As String, Optional ByVal strDefaultPath As String, Optional ByVal strTitle As String, Optional ByVal strDefaultExt As String, Optional ByVal strInitDir As String) As String

    ShowSave = ShowDialogHelper(True, strFilter, strDefaultPath, strTitle, strDefaultExt, strInitDir)

End Function

'this is a helper function for the ShowOpen and ShowSave functions. You should
'not call this function from code. It initializes the OPENFILENAME structure.
'this function relies on the GetActiveWindow() api call and OPENFILENAME structure
Private Function ShowDialogHelper(ByVal ShowSaveDialog As Boolean, Optional ByVal strFilter As String, Optional ByVal strDefaultPath As String, Optional ByVal strTitle As String, Optional ByVal strDefaultExt As String, Optional ByVal strInitDir As String) As String
    'coded by Lewis Miller

    Dim udtOpenFile As OPENFILENAME, lngResult As Long

    On Error Resume Next
    If Len(strInitDir) = 0 Then strInitDir = CurDir$
    If Len(strTitle) = 0 Then strTitle = "Select File..."
    If Len(strFilter) = 0 Then strFilter = "All Files (*.*)|*.*"
    With udtOpenFile
        .lStructSize = Len(udtOpenFile)
        .hWndOwner = GetActiveWindow
        .lpstrInitialDir = strInitDir
        .lpstrDefExt = strDefaultExt
        .lpstrTitle = strTitle
        '.lpstrFilter = Replace$(Replace$(strFilter, "|", vbNullChar), ";", vbNullChar) & vbNullChar & vbNullChar
         .lpstrFilter = Replace$(strFilter, "|", vbNullChar) & vbNullChar & vbNullChar
       .nFilterIndex = 1
        .nMaxFile = 260
        .nMaxFileTitle = 260
        .lpstrFile = String$(260, vbNullChar)
        .lpstrFileTitle = .lpstrFile
        If ShowSaveDialog Then
            lngResult = GetSaveFileName(udtOpenFile)
        Else
            lngResult = GetOpenFileName(udtOpenFile)
        End If
        If lngResult Then
            .nMaxFile = InStr(.lpstrFile, vbNullChar)
            If .nMaxFile > 0 Then
                .lpstrFile = Left$(.lpstrFile, .nMaxFile - 1)
                If Len(.lpstrFile) > 0 Then
                    ShowDialogHelper = .lpstrFile
                    Exit Function
                End If
            End If
        End If
        ShowDialogHelper = strDefaultPath
    End With

    On Error GoTo 0

End Function


Public Function ShowBrowseFolder(ByVal hWndOwner As Long, ByVal sPrompt As String, Optional ByVal sStartFolder As String = "") As String
    Dim oShell As Object
    Dim oFolder As Object
    
    Set oShell = CreateObject("Shell.Application")
    If Len(sStartFolder) > 0 Then
        Set oFolder = oShell.BrowseForFolder(hWndOwner, sPrompt, &H41, CVar(sStartFolder))
    Else
        Set oFolder = oShell.BrowseForFolder(hWndOwner, sPrompt, &H41)
    End If
    
    If Not oFolder Is Nothing Then
        ShowBrowseFolder = oFolder.Self.Path
    End If
    
    Set oFolder = Nothing
    Set oShell = Nothing
End Function
