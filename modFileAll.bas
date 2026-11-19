Attribute VB_Name = "FileAll"
Option Explicit
'some file manipulation functions, this is a pre-requisite file i include
'in most of my projects, usually i go thru an strip unused functions

'All Code Written (Or *Heavily* Modified) by Lewis Miller

'the following api calls and structures are used within this module
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long


Public Enum BYTE_TYPE
    TYPE_BYTE
    TYPE_KB
    TYPE_MB
    TYPE_GB
    TYPE_TB
End Enum


'function relies on shellexecute() and GetActiveWindow api calls
'use this only for paths and params that will be less than 260 in length
Function FileExecute(ByVal strFilePath As String, Optional ByVal strVerb As String = "open", Optional ByVal strWorkingDir As String, Optional ByVal strCommandLine As String, Optional ByVal lngWindowState As Long = 1) As Boolean
    'coded by Lewis Miller

    Dim lngReturnValue As Long


    If Len(strFilePath) > 1 Then
        strFilePath = Trim$(strFilePath)

        If Len(strWorkingDir) = 0 Then    'no folder path
            'make a folder path from file path
            lngReturnValue = InStrRev(strFilePath, "\")
            If lngReturnValue > 0 Then
                strWorkingDir = Left$(strFilePath, lngReturnValue - 1)
                If Len(strWorkingDir) = 2 Then    'C:
                    strWorkingDir = strWorkingDir & "\"
                End If
            Else
                strWorkingDir = CurDir$    'use default
            End If
        End If

        If Len(strVerb) = 0 Then strVerb = "open"

        'quote filepath to prevent long filename errors
        If Left$(strFilePath, 1) <> """" And Right$(strFilePath, 1) <> """" Then
            strFilePath = """" & strFilePath & """"
        End If

        lngReturnValue = ShellExecute(GetActiveWindow, strVerb, strFilePath, strCommandLine, strWorkingDir, lngWindowState)

    Else
        'fail, no filepath...
        lngReturnValue = 2
    End If

    If lngReturnValue <= 32 Then
        '        Select Case lngReturnValue
        '            Case 2, 3
        '                MsgBox "The file was not found. Please make sure that the file is in the folder.", vbCritical
        '            Case 8
        '                MsgBox "There was insufficient memory to execute the file.", vbCritical
        '            Case Else
        '                MsgBox "An unknown error occured when trying to start the program.", vbCritical
        '        End Select
    Else
        FileExecute = True
    End If


End Function

'supports long file names and arguments over 280 chars
Function FileExecuteWide(ByVal strFilePath As String, Optional ByVal strCommandLine As String, Optional ByVal strVerb As String = "open", Optional ByVal strWorkingDir As String, Optional ByVal lngWindowState As Long = 1) As Boolean
    'coded by Lewis Miller

    'here are the window show states if needed
    'Private Const SW_HIDE As Long = 0
    'SW_HIDE  Hides the window and activates another window.
    'Private Const SW_NORMAL As Long = 1
    'Private Const SW_SHOWNORMAL As Long = 1
    'SW_SHOWNORMAL  Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
    'Private Const SW_SHOWMINIMIZED As Long = 2
    'SW_SHOWMINIMIZED  Activates the window and displays it as a minimized window.
    'Private Const SW_MAXIMIZE As Long = 3
    'SW_MAXIMIZE  Maximizes the specified window.
    'Private Const SW_SHOWMAXIMIZED As Long = 3
    'SW_SHOWMAXIMIZED  Activates the window and displays it as a maximized window.
    'Private Const SW_SHOWNOACTIVATE As Long = 4
    'SW_SHOWNOACTIVATE  Displays a window in its most recent size and position. The active window remains active.
    'Private Const SW_SHOW As Long = 5
    'SW_SHOW  Activates the window and displays it in its current size and position.
    'Private Const SW_MINIMIZE As Long = 6
    'SW_MINIMIZE  Minimizes the specified window and activates the next top-level window in the z-order.
    'Private Const SW_SHOWMINNOACTIVE As Long = 7
    'SW_SHOWMINNOACTIVE  Displays the window as a minimized window. The active window remains active.
    'Private Const SW_SHOWNA As Long = 8
    'SW_SHOWNA  Displays the window in its current state. The active window remains active.
    'Private Const SW_RESTORE As Long = 9
    'SW_RESTORE  Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    'Private Const SW_SHOWDEFAULT As Long = 10
    'SW_SHOWDEFAULT  Sets the show state based on the SW_ flag specified in theSTARTUPINFO structure passed to theCreateProcess function by the program that started the application. An application should callShowWindow with this flag to set the initial show state of its main window.
    'Private Const SW_FORCEMINIMIZE As Long = 11

    Dim lngReturnValue As Long

    If Len(strFilePath) > 1 Then
        strFilePath = Trim$(strFilePath)

        If Len(strWorkingDir) = 0 Then    'no folder path
            'make a folder path from file path
            lngReturnValue = InStrRev(strFilePath, "\")
            If lngReturnValue > 0 Then
                strWorkingDir = Left$(strFilePath, lngReturnValue - 1)
                If Len(strWorkingDir) = 2 Then    'C:
                    strWorkingDir = strWorkingDir & "\"
                End If
            Else
                strWorkingDir = CurDir$    'use default
            End If
        End If

        'VERBS that can be used
        '-------------------------------------------------------
        'edit: Launches an editor and opens the document for editing. If lpFile is not a document file, the function will fail.
        'explore: Explores a folder specified by lpFile.
        'Find: Initiates a search beginning in the directory specified by lpDirectory.
        'open: Opens the item specified by the lpFile parameter. The item can be a file or folder.
        'Print: Prints the file specified by lpFile. If lpFile is not a document file, the function fails.
        'runas: Launches an application as Administrator. User Account Control (UAC) will prompt the user for consent to run the
        '       application elevated or enter the credentials of an administrator account used to run the application.
        'NULL: The default verb is used, if available. If not, the "open" verb is used. If neither verb is available, the system
        '       uses the first verb listed in the registry.

        'quote file path to prevent long filename errors
        If InStr(strFilePath, " ") Then
            If Left$(strFilePath, 1) <> """" And Right$(strFilePath, 1) <> """" Then
                strFilePath = """" & strFilePath & """"
            End If
        End If

        lngReturnValue = ShellExecuteW(GetActiveWindow(), ByVal StrPtr(strVerb), ByVal StrPtr(strFilePath), ByVal StrPtr(strCommandLine), ByVal StrPtr(strWorkingDir), lngWindowState)
    Else
        'fail, no filepath...
        lngReturnValue = 2
    End If

    If lngReturnValue <= 32 Then
        '        Select Case lngReturnValue
        '            Case 2, 3
        '                MsgBox "The file was not found. Please make sure that the file is in the folder.", vbCritical
        '            Case 8
        '                MsgBox "There was insufficient memory to execute the file.", vbCritical
        '            Case Else
        '                MsgBox "An unknown error occured when trying to start the program.", vbCritical
        '        End Select
    Else
        FileExecuteWide = True
    End If


End Function

'delete a file
Sub FileKill(ByVal strFilePath As String)

    On Error Resume Next

    'SetAttr strFilePath, vbNormal
    Kill strFilePath
    On Error GoTo 0

End Sub

'checks to see if a file exists
Function FileExist(ByVal strFilePath As String) As Boolean
    'coded by Lewis Miller

    On Error GoTo ErrHandle

    '    If Not ((GetAttr(strFilePath) And vbDirectory) = vbDirectory) Then
    '        FileExist = (Dir$(strFilePath) <> vbNullString)
    '    End If
    FileExist = (GetAttr(strFilePath) And vbDirectory) = 0
ErrHandle:
    On Error GoTo 0

End Function

'checks to see if a file is readable (can also be used as a fileexist)
Function FileCanRead(ByVal strFilePath As String) As Boolean
    'coded by Lewis Miller

    Dim intFileNum As Integer
    intFileNum = FreeFile
    If Len(strFilePath) > 2 Then
        On Error Resume Next
        Open strFilePath For Input As #intFileNum
        FileCanRead = (Err = 0)
        Close #intFileNum
        On Error GoTo 0
    End If

End Function

'dumps a file's contents into a string
Function FileToString(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim intFileNum As Integer

    If FileCanRead(strFilePath) Then
        On Error Resume Next
        intFileNum = FreeFile
        Open strFilePath For Binary As #intFileNum
        FileToString = Space$(LOF(intFileNum))
        Get #intFileNum, , FileToString
        Close #intFileNum
        On Error GoTo 0
    End If

End Function



'saves a string to a file as is (overwrites any previous file data)
Public Function FileSave(ByVal strFileBuffer As String, ByVal strFilePath As String) As Boolean
    'coded by Lewis Miller

    Dim intFileNum As Integer

    If FileCanRead(strFilePath) Then FileKill strFilePath
     
    On Error Resume Next
    intFileNum = FreeFile
    Open strFilePath For Binary As #intFileNum
    Put #intFileNum, , strFileBuffer
    Close #intFileNum
    FileSave = (Err.Number = 0)
    On Error GoTo 0
    
End Function




'adds a new file extension to a file name or path
'can also remove extension (and dot) by making strNewExtension = ""
Function FileAddExtension(ByVal strFilePath As String, ByVal strNewExtension As String) As String
    'coded by Lewis Miller

    Dim lngPathLength As Long, lngFindPlace As Long

    lngPathLength = Len(strFilePath)
    If lngPathLength > 0 Then
        lngFindPlace = InStrRev(strFilePath, "\")
        If (lngFindPlace > 0) And (lngFindPlace < lngPathLength) Then
            FileAddExtension = Left$(strFilePath, lngFindPlace)
            strFilePath = Mid$(strFilePath, lngFindPlace + 1)
        End If
        lngFindPlace = InStrRev(strFilePath, ".")
        If lngFindPlace > 0 Then
            strFilePath = Left$(strFilePath, lngFindPlace - 1)
        End If
    End If

    If Len(strNewExtension) > 0 Then
        If Left$(strNewExtension, 1) <> "." Then
            strNewExtension = "." & strNewExtension
        End If
    End If
    FileAddExtension = FileAddExtension & strFilePath & strNewExtension

End Function

Function FileGetExtension(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim lngPathLength As Long, lngFindPlace As Long

    lngPathLength = Len(strFilePath)
    If lngPathLength > 0 Then
        lngFindPlace = InStrRev(strFilePath, "\")
        If (lngFindPlace > 0) And (lngFindPlace < lngPathLength) Then
            strFilePath = Mid$(strFilePath, lngFindPlace + 1)
        End If
        lngFindPlace = InStrRev(strFilePath, ".")
        If lngFindPlace > 0 And lngFindPlace < Len(strFilePath) Then
            strFilePath = Mid$(strFilePath, lngFindPlace + 1)
        End If
    End If

    FileGetExtension = strFilePath

End Function

Function FileGetFolder(ByVal strFilePath As String) As String
    FileGetFolder = FolderGetPath(strFilePath)
End Function


'grabs folder path from "strFilePath"
'Note: this must be passed a file path, a folder path with no trailing slash
'      will be treated as a file path with no extension. Also drives with
'      trailing slashes will be returned without them, ex: getfolderpath("C:\") will = "C:"
Function FolderGetPath(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim lngPathLength As Long, lngSlashMark As Long

    lngPathLength = Len(strFilePath)

    If lngPathLength > 2 Then
        If Right$(strFilePath, 1) = "\" Then    'trim any trailing slashes and return
            Do While (lngPathLength > 0) And (Right$(strFilePath, 1) = "\")
                strFilePath = Left$(strFilePath, lngPathLength - 1)
                lngPathLength = Len(strFilePath)
            Loop
            FolderGetPath = strFilePath
            Exit Function
        End If
    End If
    If lngPathLength > 0 Then
        lngSlashMark = InStrRev(strFilePath, "\")
        If lngSlashMark > 0 Then
            strFilePath = Left$(strFilePath, lngSlashMark - 1)
        End If
    End If

    FolderGetPath = strFilePath

End Function


'grabs file name from file path
Function FileGetName(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim lngPathLen As Long, lngSlashMark As Long

    lngPathLen = Len(strFilePath)
    If lngPathLen > 0 Then
        If Right$(strFilePath, 1) = "\" Then
            Exit Function
        End If
        lngSlashMark = InStr(strFilePath, "\")
        If lngSlashMark > 0 And lngSlashMark < lngPathLen Then
            strFilePath = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
        End If
    End If
    FileGetName = strFilePath

End Function


'checks to see if a folder exists
Function FolderExist(ByVal strFolderPath As String) As Boolean
    'coded by Lewis Miller

    On Error Resume Next
    If Len(strFolderPath) > 1 Then
        If Right$(strFolderPath, 1) = "\" Then
            strFolderPath = strFolderPath & "nul"    'microsoft will probably remove this little MSDOS trick eventually
        Else
            strFolderPath = strFolderPath & "\nul"
        End If
        FolderExist = (Dir$(strFolderPath) <> vbNullString)
    End If
    On Error GoTo 0

End Function

'error trapped wrapper to create a folder
Sub FolderCreate(ByVal strFolderPath As String)
    'coded by Lewis Miller

    On Error Resume Next
    MkDir strFolderPath
    On Error GoTo 0

End Sub

Sub MakeDir(ByVal strPath As String)
    If Not FolderExist(strPath) Then
        FolderCreate strPath
    End If
End Sub






'simple function that does some checking of a strFolderPath
'and removes trailing slashes
Function FormatPath(ByVal strFolderPath As String) As String
    'coded by Lewis Miller

    On Error Resume Next
    If Len(strFolderPath) > 2 Then
        Do While Right$(strFolderPath, 1) = "\"
            strFolderPath = Left$(strFolderPath, Len(strFolderPath) - 1)
        Loop
        strFolderPath = Replace$(strFolderPath, "/", "\")
    End If

    If Len(strFolderPath) > 2 Then
        FormatPath = Left$(strFolderPath, 2) & Replace$(Mid$(strFolderPath, 3), "\\", "\")

    Else
        FormatPath = strFolderPath
    End If

End Function

'makes a composite file path from a base folder path and a relative file path
'assumes correct input...
Function AbsoluteFromRelative(ByVal strBaseDir As String, ByVal strFilePath As String) As String
    'if strFilePath is already a full drive or UNC path, return as-is
    If Mid$(strFilePath, 2, 1) = ":" Or Left$(strFilePath, 2) = "\\" Then
        AbsoluteFromRelative = strFilePath
        Exit Function
    End If

    'strip any trailing backslash from strBaseDir so InStrRev always steps up a folder
    If Right$(strBaseDir, 1) = "\" Then
        strBaseDir = Left$(strBaseDir, Len(strBaseDir) - 1)
    End If

    'strip leading ".\" if present
    If Left$(strFilePath, 2) = ".\" Then
        strFilePath = Mid$(strFilePath, 3)
    End If

    If InStr(strFilePath, "..\") Then
        On Error Resume Next
        Do While Left$(strFilePath, 3) = "..\" And Len(strBaseDir) > 2
            strBaseDir = Left$(strBaseDir, InStrRev(strBaseDir, "\") - 1)
            strFilePath = Mid$(strFilePath, 4)
        Loop
        On Error GoTo 0
    End If

    AbsoluteFromRelative = FormatPath(strBaseDir & "\" & strFilePath)
End Function
