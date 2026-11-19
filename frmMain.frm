VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FRX Browser"
   ClientHeight    =   12495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17880
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   833
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1192
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportAll 
      Caption         =   "Export All"
      Height          =   495
      Left            =   11940
      TabIndex        =   20
      ToolTipText     =   "Export All resources in this Entry to a Folder"
      Top             =   540
      Width           =   1500
   End
   Begin VB.CommandButton cmdExportOne 
      Caption         =   "Export File"
      Height          =   495
      Left            =   10380
      TabIndex        =   19
      ToolTipText     =   "Export Single resource Entry to file"
      Top             =   540
      Width           =   1500
   End
   Begin VB.CommandButton cmdHexView 
      Caption         =   "View Hex data"
      Height          =   495
      Left            =   8820
      TabIndex        =   18
      Top             =   540
      Width           =   1500
   End
   Begin VB.TextBox txtContent 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "frmMain.frx":0442
      Top             =   2760
      Width           =   9555
   End
   Begin VB.PictureBox picStripViewPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   6780
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   15
      Top             =   1140
      Width           =   10995
      Begin VB.PictureBox picStripCanvas 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   0
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   429
         TabIndex        =   16
         Top             =   0
         Width           =   6495
         Begin VB.Image imgStrip 
            BorderStyle     =   1  'Fixed Single
            Height          =   720
            Index           =   0
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   915
         End
      End
   End
   Begin VB.HScrollBar hsbStrip 
      Height          =   375
      LargeChange     =   100
      Left            =   6780
      SmallChange     =   20
      TabIndex        =   14
      Top             =   2220
      Width           =   10995
   End
   Begin VB.PictureBox picFullViewPort 
      Height          =   9195
      Left            =   6720
      ScaleHeight     =   609
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   2700
      Width           =   10635
      Begin VB.PictureBox picFullCanvas 
         AutoSize        =   -1  'True
         Height          =   7935
         Left            =   0
         ScaleHeight     =   525
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   681
         TabIndex        =   13
         Top             =   0
         Width           =   10275
      End
   End
   Begin VB.VScrollBar vsbFull 
      Height          =   9135
      LargeChange     =   50
      Left            =   17460
      SmallChange     =   10
      TabIndex        =   11
      Top             =   2820
      Width           =   375
   End
   Begin VB.HScrollBar hsbFull 
      Height          =   435
      LargeChange     =   50
      Left            =   6720
      SmallChange     =   10
      TabIndex        =   10
      Top             =   12000
      Width           =   10575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   540
      Width           =   1575
   End
   Begin VB.ListBox lstEntries 
      Height          =   10260
      Left            =   2820
      TabIndex        =   7
      Top             =   2220
      Width           =   3735
   End
   Begin VB.ListBox lstFiles 
      Height          =   10260
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   2655
   End
   Begin VB.TextBox txtProject 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1260
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   495
      Left            =   6180
      TabIndex        =   1
      ToolTipText     =   "Browse For File"
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox txtPath 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   5955
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource File Entries"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      TabIndex        =   8
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource Files"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   6
      Top             =   1860
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "VB6 Project"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add File: Form, Designers, Resource, And Project Files Are Accepted."
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Module-level or local to ProcessFormFile
Private mCtrlStack() As String
Private mCtrlStackCount As Long


Private mEntryIdx As Long
Private mResIdx As Long

Private Sub cmdHexView_Click()
    Dim sTitle As String
    If mEntryIdx < 0 Or mEntryIdx >= mEntryCount Then Exit Sub
    If mAllEntries(mEntryIdx).ResCount = 0 Then Exit Sub
    
    With mAllEntries(mEntryIdx)
        sTitle = .ControlName & "." & .PropName & " [Item " & (mResIdx + 1) & " of " & .ResCount & " | " & .Res(mResIdx).Size & " bytes]"
        frmViewer.ShowHex sTitle, .Res(mResIdx).DataBytes
    End With
End Sub

Private Sub cmdBrowse_Click()
    Dim S As String

    S = ShowOpen("All Supported|*.frm;*.frx;*.vbp;*.pag;*.dsr;*.ctl;*.dsx;*.pgx;*.ctx|Project Files|*.vbp|Form Files|*.frm|Frx Files|*.frx|UserControl Files|*.ctl|Property Pages|*.pag|ActiveX Designer|*.dsr", "", "Select File", "", CurDir$)
    If Len(S) = 0 Then Exit Sub
    LoadFile S

End Sub

Private Sub cmdRefresh_Click()

    If FileExist(txtPath) Then
        LoadFile txtPath
    Else
        If Len(txtPath) = 0 Then
            cmdBrowse_Click
        Else
            MsgBox "The file cannot be accessed or no longer exists! Please select another file.", vbCritical
        End If
    End If

End Sub

Sub LoadFile(ByVal strPath As String)
    Dim S As String
    Dim E As String
    Dim sCompanion As String
    
    S = Trim$(strPath)
    
    If Not FileExist(S) Then
       MsgBox "The selected file could not be accessed or no longer exists!", vbCritical
       Exit Sub
    End If
    
    txtPath = S
    SaveStr "ProjectPath", S
        
    E = FileGetExtension(S)
    BaseFolder = FileGetFolder(S)

    lstFiles.Clear
    lstEntries.Clear
    mFrxFileCount = 0
    Erase mFrxFiles
    mEntryCount = 0
    Erase mAllEntries
    txtProject = ""

    Select Case LCase$(E)
        Case "vbp"
            txtProject = FileGetName(S)
            ProcessProjectFile S

        Case "frm", "ctl", "pag", "dsr"
            txtProject = ""
            ProcessFormFile S, True

        Case "frx", "pgx", "dsx", "ctx"
            If Dir$(FileAddExtension(S, "frm")) <> "" Then
                sCompanion = FileAddExtension(S, "frm")
            ElseIf Dir$(FileAddExtension(S, "ctl")) <> "" Then
                sCompanion = FileAddExtension(S, "ctl")
            ElseIf Dir$(FileAddExtension(S, "pag")) <> "" Then
                sCompanion = FileAddExtension(S, "pag")
            ElseIf Dir$(FileAddExtension(S, "dsr")) <> "" Then
                sCompanion = FileAddExtension(S, "dsr")
            End If
            
            If Len(sCompanion) > 0 Then
                LoadFile sCompanion
                Exit Sub ' Prevent outer call from re-running ListIndex = 0
            Else
                ScanStandaloneFrx S
            End If
            
        Case Else
            MsgBox "Sorry That File Type Is Not Supported.", vbCritical
            Exit Sub
    End Select
    
    If lstFiles.ListCount > 0 Then
        lstFiles.ListIndex = 0
    Else
        MsgBox "No FRX binary resources were found in the selected file.", vbInformation
    End If
End Sub

Private Sub ProcessProjectFile(ByVal sVbpPath As String)
    Dim iFile As Integer
    Dim sLine As String
    Dim sRelPath As String
    Dim sFullFrm As String
    
    iFile = FreeFile
    Open sVbpPath For Input As #iFile
    Do Until EOF(iFile)
        Line Input #iFile, sLine
        sLine = Trim$(sLine)
        
        If InStr(1, sLine, "Form=", vbTextCompare) = 1 Or _
           InStr(1, sLine, "UserControl=", vbTextCompare) = 1 Or _
           InStr(1, sLine, "Designer=", vbTextCompare) = 1 Or _
           InStr(1, sLine, "PropertyPage=", vbTextCompare) = 1 Then
           
            sRelPath = Mid$(sLine, InStr(1, sLine, "=") + 1)
            If InStr(1, sRelPath, ";") > 0 Then
                sRelPath = Trim$(Mid$(sRelPath, InStr(1, sRelPath, ";") + 1))
            End If
            
            sFullFrm = AbsoluteFromRelative(BaseFolder, sRelPath)
            
            ' ProcessFormFile will register the .frx via RegisterFrxFile only if entries exist
            ProcessFormFile sFullFrm, False
        End If
    Loop
    Close #iFile
End Sub

Private Sub ProcessFormFile(ByVal sFrmPath As String, Optional ByVal bResetArray As Boolean = True)
    Dim iFile As Integer
    Dim sLine As String
    Dim lPosFrx As Long
    Dim sFrxName As String
    Dim sHexOffset As String
    Dim sProp As String
    Dim sFullFrx As String
    Dim sParts() As String
    Dim sKey As String
    Dim lQuoteStart As Long

    If bResetArray Then
        mEntryCount = 0
        Erase mAllEntries
    End If

    ' Reset control stack for this form
    mCtrlStackCount = 0
    Erase mCtrlStack

    If Not FileExist(sFrmPath) Then Exit Sub

    ' Seed with the form name itself
    PushControl FileGetName(sFrmPath)

    iFile = FreeFile
    Open sFrmPath For Input As #iFile
    Do Until EOF(iFile)
        Line Input #iFile, sLine
        sLine = Trim$(sLine)

        If Left$(sLine, 6) = "Begin " Then
            sParts = Split(sLine, " ")
            If UBound(sParts) >= 2 Then
                PushControl sParts(2)
            End If
        ElseIf sLine = "End" Then          ' <-- exact match only
            PopControl
        ElseIf InStr(1, sLine, "=") > 0 Then
            sKey = Trim$(Left$(sLine, InStr(1, sLine, "=") - 1))
            If StrComp(sKey, "Index", vbTextCompare) = 0 Then
                If mCtrlStackCount > 0 Then
                    mCtrlStack(mCtrlStackCount - 1) = mCtrlStack(mCtrlStackCount - 1) & _
                        "(" & Trim$(Mid$(sLine, InStr(1, sLine, "=") + 1)) & ")"
                End If
            End If
        End If

        lPosFrx = FindBinaryRefPos(sLine)
        If lPosFrx > 0 Then
            sProp = Trim$(Left$(sLine, InStr(1, sLine, "=") - 1))
            lQuoteStart = InStrRev(sLine, """", lPosFrx)
            sFrxName = Mid$(sLine, lQuoteStart + 1, (lPosFrx + 4) - lQuoteStart - 1)
            sHexOffset = Trim$(Mid$(sLine, lPosFrx + 6))
            sFullFrx = AbsoluteFromRelative(FileGetFolder(sFrmPath), sFrxName)

            ReDim Preserve mAllEntries(mEntryCount)
            With mAllEntries(mEntryCount)
                .SourceForm = FileGetName(sFrmPath)
                .ControlName = CurrentControl()
                .PropName = sProp
                .FrxIndex = RegisterFrxFile(sFullFrx)
                .Offset = Val("&H" & sHexOffset & "&")
                ExtractFrxEntry mAllEntries(mEntryCount)
            End With
            mEntryCount = mEntryCount + 1
        End If
    Loop
    Close #iFile
End Sub

Private Function RegisterFrxFile(ByVal sFullFrxPath As String) As Long
    Dim i As Long
    Dim sDisplay As String
    Dim sBase As String
    
    'check if we already registered this exact full path
    For i = 0 To mFrxFileCount - 1
        If StrComp(mFrxFiles(i), sFullFrxPath, vbTextCompare) = 0 Then
            RegisterFrxFile = i
            Exit Function
        End If
    Next i
    
    ' add to frx files array
    ReDim Preserve mFrxFiles(mFrxFileCount)
    mFrxFiles(mFrxFileCount) = sFullFrxPath
    
    ' format display name for lstFiles
    sBase = BaseFolder
    If Right$(sBase, 1) <> "\" Then sBase = sBase & "\"
    
    If StrComp(Left$(sFullFrxPath, Len(sBase)), sBase, vbTextCompare) = 0 Then
        sDisplay = Mid$(sFullFrxPath, Len(sBase) + 1)
    Else
        sDisplay = sFullFrxPath
    End If
    
    lstFiles.AddItem sDisplay
    lstFiles.ItemData(lstFiles.NewIndex) = mFrxFileCount
    
    RegisterFrxFile = mFrxFileCount
    mFrxFileCount = mFrxFileCount + 1
End Function

Public Sub ScanStandaloneFrx(ByVal sFrxPath As String)
    Dim lFrxIdx As Long
    Dim iFile As Integer
    Dim bFile() As Byte
    Dim lFileSize As Long
    Dim i As Long
    Dim lLastEnd As Long
    Dim lImgSize As Long
    Dim bRawImg() As Byte
    Dim sExt As String
    Dim lGapStart As Long
    Dim lGapLen As Long
    
    lFrxIdx = RegisterFrxFile(sFrxPath)
    
    On Error GoTo Error_Handle
    iFile = FreeFile
    Open sFrxPath For Binary Access Read As #iFile
    lFileSize = LOF(iFile)
    If lFileSize = 0 Then
        Close #iFile
        Exit Sub
    End If
    ReDim bFile(0 To lFileSize - 1)
    Get #iFile, , bFile
    Close #iFile
    On Error GoTo 0
    
    lLastEnd = 0
    i = 4
    Do While i <= lFileSize - 8
        If bFile(i) = &H6C And bFile(i + 1) = &H74 And bFile(i + 2) = &H0 And bFile(i + 3) = &H0 Then
            CopyMemory lImgSize, bFile(i + 4), 4
            If lImgSize > 0 And (i + 8 + lImgSize) <= lFileSize Then
                ' capture any meaningful unclaimed gap before this picture
                '    (skip 20-byte CLSID_StdPicture wrapper headers so they don't clutter the list)
                lGapLen = (i - 4) - lLastEnd
                If lGapLen > 20 Then
                    AddRawGapEntry lFrxIdx, bFile, lLastEnd, lGapLen, FileGetName(sFrxPath)
                End If
                
                ' copy the raw image bytes out of the file buffer
                ReDim bRawImg(0 To lImgSize - 1)
                CopyMemory bRawImg(0), bFile(i + 8), lImgSize
                sExt = DetectImageExt(bRawImg)
                
                ' create the entry and unpack multi-icon .ico (with single-image fallback)
                ReDim Preserve mAllEntries(mEntryCount)
                With mAllEntries(mEntryCount)
                    .FrxIndex = lFrxIdx
                    .SourceForm = FileGetName(sFrxPath)
                    .ControlName = "CarvedImage_" & mEntryCount
                    .PropName = "Picture"
                    .Offset = i - 4
                    .ResCount = 0
                    
                    If sExt = "ico" Then
                        UnpackIconContainer bRawImg, mAllEntries(mEntryCount)
                    End If
                    
                    If .ResCount = 0 Then
                        .ResCount = 1
                        ReDim .Res(0)
                        .Res(0).DataType = resPicture
                        .Res(0).Size = lImgSize
                        .Res(0).DataBytes = bRawImg
                        .Res(0).FileExt = sExt
                    End If
                End With
                mEntryCount = mEntryCount + 1
                
                ' jump past this entire image
                lLastEnd = i + 8 + lImgSize
                i = lLastEnd + 4
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    
    ' capture any trailing bytes after the last picture
    If lLastEnd < lFileSize Then
        AddRawGapEntry lFrxIdx, bFile, lLastEnd, lFileSize - lLastEnd, FileGetName(sFrxPath)
    End If
    
    Exit Sub
Error_Handle:
    MsgBox "Error Opening resource file!" & vbCrLf & Err.Description, vbCritical
    On Error GoTo 0
End Sub


Private Sub Form_Load()
    txtPath = GetStr("ProjectPath", "")
    picStripViewPort.Visible = False
    hsbStrip.Visible = False
    picFullViewPort.Top = picStripViewPort.Top
    txtContent.Top = picStripViewPort.Top
    Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveStr "ProjectPath", txtPath
   ShutdownGdiplus
   FileKill App.Path & "\frx_view.*"
   FileKill App.Path & "\frx_view(*).*"  'in case frx_viewer
End Sub


Private Sub Form_Resize()
    Dim lRightW As Long
    Dim lBottomH As Long
    Dim lBottomEdge As Long
    Dim lViewH As Long
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Enforce a sensible minimum window size
    If Me.Width < 14000 Then Me.Width = 14000
    If Me.Height < 9500 Then Me.Height = 9500
    
    lBottomH = Me.ScaleHeight - lstFiles.Top - 16
    If lBottomH > 100 Then
        lstFiles.Height = lBottomH
        lstEntries.Height = lBottomH
    End If
    
    lBottomEdge = lstFiles.Top + lstFiles.Height
    
    lRightW = Me.ScaleWidth - picFullViewPort.Left - vsbFull.Width - 12
    If lRightW > 150 Then
        picStripViewPort.Width = lRightW
        hsbStrip.Width = lRightW
        picFullViewPort.Width = lRightW
        hsbFull.Width = lRightW
        vsbFull.Left = picFullViewPort.Left + lRightW + 2
        txtContent.Width = lRightW + vsbFull.Width + 2
    End If
    
    hsbFull.Top = lBottomEdge - hsbFull.Height
    
    lViewH = (hsbFull.Top - 4) - picFullViewPort.Top
    If lViewH > 80 Then
        picFullViewPort.Height = lViewH
        vsbFull.Height = lViewH
        vsbFull.Top = picFullViewPort.Top
    End If
    txtContent.Height = lBottomEdge - txtContent.Top
    
    If mEntryCount > 0 And mEntryIdx >= 0 Then
        ShowResource mEntryIdx, mResIdx
    End If
End Sub


Private Sub lstFiles_Click()
    Dim lFileIdx As Long
    Dim i As Long
    Dim sExt As String
    
    lstEntries.Clear
    If lstFiles.ListIndex < 0 Then Exit Sub
    
    lFileIdx = lstFiles.ItemData(lstFiles.ListIndex)
    
    For i = 0 To mEntryCount - 1
        If mAllEntries(i).FrxIndex = lFileIdx Then
            If mAllEntries(i).ResCount > 0 Then
                sExt = mAllEntries(i).Res(0).FileExt
                If mAllEntries(i).ResCount > 1 Then
                    sExt = sExt & " x" & mAllEntries(i).ResCount
                End If
            Else
                sExt = "empty"
            End If
            
            lstEntries.AddItem mAllEntries(i).ControlName & "." & _
                               mAllEntries(i).PropName & "  [0x" & Hex$(mAllEntries(i).Offset) & " | " & sExt & "]"
            lstEntries.ItemData(lstEntries.NewIndex) = i
        End If
    Next i
    
    If lstEntries.ListCount > 0 Then
        lstEntries.ListIndex = 0
    End If
        
    
End Sub

Private Sub lstEntries_Click()
    If lstEntries.ListIndex < 0 Or mEntryCount = 0 Then Exit Sub
    ShowResource lstEntries.ItemData(lstEntries.ListIndex), 0

    'DebugDumpOffset BaseFolder & "\" & lstFiles.List(lstFiles.ListIndex), mAllEntries(Index).Offset
End Sub



Public Sub ShowResource(ByVal lEntryIdx As Long, ByVal lResIdx As Long)
    mEntryIdx = lEntryIdx
    mResIdx = lResIdx

    If mEntryIdx < 0 Or mEntryIdx >= mEntryCount Then Exit Sub

    With mAllEntries(mEntryIdx)
        Me.Caption = "FRX Browser - " & .ControlName & "." & .PropName & " (" & .ResCount & " item(s))"

        If .ResCount = 0 Or lResIdx >= .ResCount Then
            picStripViewPort.Visible = False
            hsbStrip.Visible = False
            picFullViewPort.Visible = False
            vsbFull.Visible = False
            hsbFull.Visible = False
            txtContent.Visible = True
            txtContent.Text = "(No extractable data at offset 0x" & Hex$(.Offset) & ")"
            Exit Sub
        End If

        If .Res(mResIdx).DataType = resPicture Then
            txtContent.Visible = False
            picFullViewPort.Visible = True
            vsbFull.Visible = True
            hsbFull.Visible = True

            ' show horizontal filmstrip when entry has multiple images (e.g., ImageList / Toolbar)
            If .ResCount > 1 Then
                picStripViewPort.Visible = True
                hsbStrip.Visible = True
                BuildHorizontalStrip mEntryIdx
                picFullViewPort.Top = hsbStrip.Top + hsbStrip.Height + 8
                picFullViewPort.Height = (hsbFull.Top - 6) - picFullViewPort.Top
                vsbFull.Top = picFullViewPort.Top
                vsbFull.Height = picFullViewPort.Height
            Else
                picStripViewPort.Visible = False
                hsbStrip.Visible = False
                picFullViewPort.Top = picStripViewPort.Top
                picFullViewPort.Height = (hsbFull.Top - 6) - picFullViewPort.Top
                vsbFull.Top = picFullViewPort.Top
                vsbFull.Height = picFullViewPort.Height
            End If

            DisplayFullPicture mResIdx
        Else
            picStripViewPort.Visible = False
            hsbStrip.Visible = False
            picFullViewPort.Visible = False
            vsbFull.Visible = False
            hsbFull.Visible = False

            txtContent.Top = picStripViewPort.Top
            txtContent.Left = picFullViewPort.Left
            txtContent.Width = picFullViewPort.Width
            txtContent.Height = (hsbFull.Top + hsbFull.Height) - txtContent.Top
            txtContent.Visible = True
            txtContent.Text = FormatResourcePreview(.Res(mResIdx))
        End If
    End With
End Sub

Private Sub BuildHorizontalStrip(ByVal mEntryIdx As Long)
    Dim i As Long
    Dim lTotalWidth As Long

    For i = imgStrip.UBound To 1 Step -1
        Unload imgStrip(i)
    Next i
    picStripCanvas.Left = 0
    hsbStrip.Value = 0

    For i = 0 To mAllEntries(mEntryIdx).ResCount - 1
        If i > 0 Then Load imgStrip(i)
        With imgStrip(i)
            .Left = 6 + i * 54
            .Top = 6
            .Width = 48
            .Height = 48
            Set .Picture = LoadBytesAsPicture(mAllEntries(mEntryIdx).Res(i))
            .Visible = True
        End With
    Next i

    lTotalWidth = 6 + mAllEntries(mEntryIdx).ResCount * 54
    picStripCanvas.Width = lTotalWidth

    If lTotalWidth > picStripViewPort.ScaleWidth Then
        hsbStrip.Max = lTotalWidth - picStripViewPort.ScaleWidth
        hsbStrip.Enabled = True
    Else
        hsbStrip.Max = 0
        hsbStrip.Enabled = False
    End If
    
    HighlightStripItem mResIdx
End Sub

Private Sub imgStrip_Click(Index As Integer)
    mResIdx = Index
    HighlightStripItem Index
    DisplayFullPicture Index
End Sub

Private Sub hsbStrip_Change()
    picStripCanvas.Left = -hsbStrip.Value
End Sub

Private Sub hsbStrip_Scroll()
    picStripCanvas.Left = -hsbStrip.Value
End Sub

Private Sub DisplayFullPicture(ByVal lResIdx As Long)
    picFullCanvas.Left = 0
    picFullCanvas.Top = 0
    Set picFullCanvas.Picture = LoadBytesAsPicture(mAllEntries(mEntryIdx).Res(lResIdx))

    ' Configure full-image scrollbars if picture exceeds viewport
    If picFullCanvas.Width > picFullViewPort.ScaleWidth Then
        hsbFull.Max = picFullCanvas.Width - picFullViewPort.ScaleWidth
        hsbFull.Enabled = True
    Else
        hsbFull.Max = 0
        hsbFull.Enabled = False
    End If

    If picFullCanvas.Height > picFullViewPort.ScaleHeight Then
        vsbFull.Max = picFullCanvas.Height - picFullViewPort.ScaleHeight
        vsbFull.Enabled = True
    Else
        vsbFull.Max = 0
        vsbFull.Enabled = False
    End If
End Sub

Private Function LoadBytesAsPicture(ByRef tRes As FrxResource) As StdPicture
    Dim oPic As StdPicture
    Dim sTempFile As String
    Dim iFile As Integer
    
    Set oPic = LoadPictureFromBytes(tRes.DataBytes, vbPicTypeBitmap, picFullViewPort.BackColor)
    If oPic Is Nothing Then
        sTempFile = App.Path & "\frx_view." & tRes.FileExt
        Do While FileExist(sTempFile)
              iFile = iFile + 1
              sTempFile = App.Path & "\frx_view(" & iFile & ")." & tRes.FileExt
        Loop
        If FileExist(sTempFile) Then FileKill sTempFile
        'this is ugly
        iFile = FreeFile
        Open sTempFile For Binary Access Write As #iFile
        Put #iFile, , tRes.DataBytes
        Close #iFile
        On Error Resume Next
        Set oPic = LoadPicture(sTempFile)
        On Error GoTo 0
    End If
    Set LoadBytesAsPicture = oPic
End Function


Private Sub hsbFull_Change()
    picFullCanvas.Left = -hsbFull.Value
End Sub

Private Sub hsbFull_Scroll()
    picFullCanvas.Left = -hsbFull.Value
End Sub

Private Sub vsbFull_Change()
    picFullCanvas.Top = -vsbFull.Value
End Sub

Private Sub vsbFull_Scroll()
    picFullCanvas.Top = -vsbFull.Value
End Sub


Private Sub cmdExportOne_Click()
    Dim sDefaultName As String
    Dim sExt As String
    Dim sFilter As String
    Dim sDestPath As String
    Dim bIsMultiIco As Boolean
    Dim bIsMultiPic As Boolean
    
    If mEntryIdx < 0 Or mEntryIdx >= mEntryCount Then Exit Sub
    If mAllEntries(mEntryIdx).ResCount = 0 Then Exit Sub
    
    With mAllEntries(mEntryIdx)
        bIsMultiIco = (.ResCount > 1 And LCase$(.Res(0).FileExt) = "ico" And InStr(1, .PropName, "listimages", vbTextCompare) = 0)
        bIsMultiPic = (.ResCount > 1 And .Res(0).DataType = resPicture And Not bIsMultiIco)
        
        ' only pop up options if this entry actually has multiple frames/images
        If bIsMultiIco Or bIsMultiPic Then
            If Not frmExportOpts.PromptExportOptions(Me, False) Then Exit Sub
        End If
        
        If bIsMultiIco And gExportOpts.PackMultiIcons Then
            sDefaultName = SanitizeFileName(Replace$(Replace$(.ControlName, "(", "_"), ")", "") & "_" & .PropName) & ".ico"
            sDestPath = ShowSave("Icon Files|*.ico|All Files|*.*", sDefaultName, "Export Multi-Resolution Icon", "ico")
            If Len(sDestPath) = 0 Then Exit Sub
            If PackEntryToMultiIco(mAllEntries(mEntryIdx), sDestPath) Then
                MsgBox "Exported " & .ResCount & " icon resolutions to:" & vbCrLf & sDestPath, vbInformation
            End If
            Exit Sub
        End If
        
        If bIsMultiPic And gExportOpts.ExportImageListsAsStrip Then
            sDefaultName = SanitizeFileName(Replace$(Replace$(.ControlName, "(", "_"), ")", "") & "_" & .PropName) & "_strip.png"
            sDestPath = ShowSave("PNG Images|*.png|All Files|*.*", sDefaultName, "Export Image Strip", "png")
            If Len(sDestPath) = 0 Then Exit Sub
            If ExportEntryAsHorizontalStrip(mAllEntries(mEntryIdx), sDestPath) Then
                MsgBox "Exported " & .ResCount & "-frame PNG strip to:" & vbCrLf & sDestPath, vbInformation
            End If
            Exit Sub
        End If
        
        ' Standard single-item export
        sDefaultName = BuildResourceFileName(mAllEntries(mEntryIdx), mResIdx)
        sExt = FileGetExtension(sDefaultName)
        sFilter = UCase$(sExt) & " Files|*." & sExt & "|All Files|*.*"
        
        sDestPath = ShowSave(sFilter, sDefaultName, "Export Resource", sExt)
        If Len(sDestPath) = 0 Then Exit Sub
        
        If SaveResourceToFile(.Res(mResIdx), sDestPath) Then
            MsgBox "Exported successfully to:" & vbCrLf & sDestPath, vbInformation
        End If
    End With
End Sub

Private Sub cmdExportAll_Click()
    Dim sChosenDir As String
    Dim sRootOut As String
    Dim sTargetDir As String
    Dim sPrefix As String
    Dim sBaseName As String
    Dim i As Long
    Dim J As Long
    Dim lSavedCount As Long
    Dim bIsMultiIco As Boolean
    Dim bIsMultiPic As Boolean
                    Dim sDummy As String

    If mEntryCount = 0 Then
        MsgBox "No resources are currently loaded to export.", vbInformation
        Exit Sub
    End If

    If Not frmExportOpts.PromptExportOptions(Me, True) Then Exit Sub

    sChosenDir = ShowBrowseFolder(Me.hwnd, "Select destination folder for exported resources:", BaseFolder)
    If Len(sChosenDir) = 0 Then Exit Sub

    If Right$(sChosenDir, 1) <> "\" Then sChosenDir = sChosenDir & "\"
    sRootOut = sChosenDir & FileAddExtension(FileGetName(txtPath), "") & "_Exported"
    MakeDir sRootOut

    'export all entries according to gExportOpts
    For i = 0 To mEntryCount - 1
        With mAllEntries(i)
            If .ResCount > 0 Then
                If gExportOpts.SkipUnindexedBlobs And .Res(0).DataType = resCustomBlob And LCase$(.Res(0).FileExt) = "bin" Then
                    If Not TryDecodePropertyBag(.Res(0).DataBytes, sDummy) Then GoTo NextEntry
                End If

                If gExportOpts.GroupByFormFolder Then
                    sTargetDir = sRootOut & "\" & SanitizeFileName(FileAddExtension(.SourceForm, ""))
                    MakeDir sTargetDir
                    sPrefix = ""
                Else
                    sTargetDir = sRootOut
                    sPrefix = SanitizeFileName(FileAddExtension(.SourceForm, "")) & "_"
                End If

                sBaseName = sPrefix & SanitizeFileName(Replace$(Replace$(.ControlName, "(", "_"), ")", "") & "_" & .PropName)
                bIsMultiIco = (.ResCount > 1 And LCase$(.Res(0).FileExt) = "ico" And InStr(1, .PropName, "listimages", vbTextCompare) = 0)
                bIsMultiPic = (.ResCount > 1 And .Res(0).DataType = resPicture And Not bIsMultiIco)

                If bIsMultiIco And gExportOpts.PackMultiIcons Then
                    If PackEntryToMultiIco(mAllEntries(i), sTargetDir & "\" & sBaseName & ".ico") Then
                        lSavedCount = lSavedCount + 1
                    End If
                ElseIf bIsMultiPic And gExportOpts.ExportImageListsAsStrip Then
                    If ExportEntryAsHorizontalStrip(mAllEntries(i), sTargetDir & "\" & sBaseName & "_strip.png") Then
                        lSavedCount = lSavedCount + 1
                    End If
                Else
                    For J = 0 To .ResCount - 1
                        If SaveResourceToFile(.Res(J), sTargetDir & "\" & sPrefix & BuildResourceFileName(mAllEntries(i), J)) Then
                            lSavedCount = lSavedCount + 1
                        End If
                    Next J
                End If
            End If
        End With
NextEntry:
    Next i

    MsgBox "Batch export complete! Saved " & lSavedCount & " resource file(s) to:" & vbCrLf & sRootOut, vbInformation
    Shell "explorer.exe """ & sRootOut & """", vbNormalFocus
End Sub

Private Sub HighlightStripItem(ByVal lSelectedIdx As Long)
    Dim lLeft As Long
    Dim lTop As Long
    Dim lRight As Long
    Dim lBottom As Long
    
    picStripCanvas.Cls
    If lSelectedIdx < 0 Or lSelectedIdx > imgStrip.UBound Then Exit Sub
    If Not imgStrip(lSelectedIdx).Visible Then Exit Sub
    
    With imgStrip(lSelectedIdx)
        lLeft = .Left - 3
        lTop = .Top - 3
        lRight = .Left + .Width + 2
        lBottom = .Top + .Height + 2
    End With
    
    ' draw a system highlight border around the selected thumbnail
    picStripCanvas.Line (lLeft, lTop)-(lRight, lBottom), vbHighlight, B
    picStripCanvas.Line (lLeft + 1, lTop + 1)-(lRight - 1, lBottom - 1), vbHighlight, B
End Sub


Private Sub PushControl(ByVal sName As String)
    ReDim Preserve mCtrlStack(mCtrlStackCount)
    mCtrlStack(mCtrlStackCount) = sName
    mCtrlStackCount = mCtrlStackCount + 1
End Sub

Private Sub PopControl()
    If mCtrlStackCount > 0 Then
        mCtrlStackCount = mCtrlStackCount - 1
    End If
End Sub

Private Function CurrentControl() As String
    If mCtrlStackCount > 0 Then
        CurrentControl = mCtrlStack(mCtrlStackCount - 1)
    End If
End Function
