Attribute VB_Name = "modVars"
Option Explicit

Public Type ExportOptions
    PackMultiIcons As Boolean       ' True = combine icon frames into 1 multi-res .ico; False = split into _01, _02...
    ExportImageListsAsStrip As Boolean    ' True = stitch ImageList items into 1 horizontal PNG strip; False = individual files
    SkipUnindexedBlobs As Boolean   ' True = skip raw .bin / UnindexedData chunks during batch export
    GroupByFormFolder As Boolean    ' True = FormName\Control_Prop.ext; False = flat folder with FormName_ prefix
    OptionsCancelled As Boolean
End Type

Public gExportOpts As ExportOptions


Public Enum FrxResourceType
    resPicture = 1      ' BMP, ICO, CUR, GIF, JPG, WMF, EMF
    resString = 2       ' Multi-line Text / Caption
    resStringList = 3   ' ComboBox / ListBox .List
    resItemData = 4     ' ComboBox / ListBox .ItemData
    resCustomBlob = 5   ' ActiveX PropertyBag / Unclaimed Hex/Text
End Enum

Public Type FrxResource
    DataType As FrxResourceType
    FileExt As String       ' "bmp", "ico", "jpg", "gif", "cur", "wmf", "txt", "bin"
    Size As Long            ' Byte length of DataBytes
    DataBytes() As Byte     ' Raw payload (clean image file bytes, or raw string/list bytes)
End Type

Public Type FrxEntry
    FrxIndex As Long        ' Direct index into mFrxFiles()
    SourceForm As String
    ControlName As String
    PropName As String
    Offset As Long
    ResCount As Long
    Res() As FrxResource    ' Usually 1 item; >1 for ImageLists / Toolbars
End Type

'Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public mAllEntries() As FrxEntry
Public mEntryCount As Long
Public mFrxFiles() As String
Public mFrxFileCount As Long

Public BaseFolder As String


Public Sub ExtractFrxEntry(ByRef tEntry As FrxEntry)
    Dim sFrxPath As String
    Dim iFile As Integer
    Dim lFileLen As Long
    Dim lPos As Long
    Dim sProp As String

    tEntry.ResCount = 0
    sFrxPath = mFrxFiles(tEntry.FrxIndex)
    If Dir$(sFrxPath) = "" Then Exit Sub

    iFile = FreeFile
    Open sFrxPath For Binary Access Read As #iFile
    lFileLen = LOF(iFile)
    lPos = tEntry.Offset + 1    ' VB6 Binary Seek is 1-based

    If lPos < 1 Or lPos > lFileLen Then
        Close #iFile
        Exit Sub
    End If

    sProp = LCase$(tEntry.PropName)

    Select Case sProp
        Case "picture", "image", "icon", "mouseicon", "dragicon", "downpicture", "disabledpicture", "maskpicture", "palette", "bkgimage"
            ExtractPictureAt iFile, lPos, tEntry

        Case "text", "textrtf", "cuebanner", "caption"
            ExtractStringAt iFile, lPos, tEntry
            If sProp = "textrtf" And tEntry.ResCount > 0 Then
                tEntry.Res(0).FileExt = "rtf"
            End If

        Case "list"
            ExtractStringListAt iFile, lPos, tEntry

        Case "itemdata"
            ExtractItemDataAt iFile, lPos, tEntry

        Case Else
            If (InStr(1, sProp, "picture") > 0 Or InStr(1, sProp, "image") > 0) And _
               InStr(1, sProp, "listimages") = 0 Then
                ExtractPictureAt iFile, lPos, tEntry
            Else
                ' ActiveX / Custom PropertyBag (e.g., ImageList InitListImages, Toolbar, etc.)
                ExtractCustomBlobAt iFile, lPos, tEntry
            End If


    End Select

    Close #iFile
End Sub

Private Sub ExtractPictureAt(ByVal iFile As Integer, ByVal lPos As Long, ByRef tEntry As FrxEntry)
    Dim bHdr(0 To 31) As Byte
    Dim lImgSize As Long
    Dim lDataStart As Long
    Dim bRawImg() As Byte
    Dim sExt As String

    If (LOF(iFile) - lPos + 1) < 12 Then Exit Sub
    Seek #iFile, lPos

    If (LOF(iFile) - lPos + 1) >= 32 Then
        Get #iFile, , bHdr
    Else
        Dim bShort(0 To 11) As Byte
        Get #iFile, , bShort
        CopyMemory bHdr(0), bShort(0), 12
    End If

    If bHdr(4) = &H6C And bHdr(5) = &H74 Then
        ' Case A: Standard 12-byte VB6 picture header ("lt\0\0" at offset 4)
        CopyMemory lImgSize, bHdr(8), 4
        lDataStart = lPos + 12
    ElseIf bHdr(20) = &H6C And bHdr(21) = &H74 Then
        ' Case B: 28-byte CLSID_StdPicture header (ImageList ListImage, "lt\0\0" at offset 20)
        CopyMemory lImgSize, bHdr(24), 4
        lDataStart = lPos + 28
    ElseIf bHdr(4) = &H0 And bHdr(5) = &H74 And bHdr(6) = &HE8 And bHdr(7) = &HB1 Then
        ' Case C: 32-byte OLE stream header (AlphaImageCtl Image/BkgImage, "00 74 E8 B1" at offset 4)
        CopyMemory lImgSize, bHdr(24), 4
        lDataStart = lPos + 32
    ElseIf bHdr(0) = &H0 And bHdr(1) = &H74 And bHdr(2) = &HE8 And bHdr(3) = &HB1 Then
        ' Case D: 28-byte unprefixed OLE stream header ("00 74 E8 B1" at offset 0)
        CopyMemory lImgSize, bHdr(20), 4
        lDataStart = lPos + 28
    Else
        Exit Sub
    End If

    If lImgSize > 0 And lImgSize <= (LOF(iFile) - lDataStart + 1) Then
        Seek #iFile, lDataStart
        ReDim bRawImg(0 To lImgSize - 1)
        Get #iFile, , bRawImg
        sExt = DetectImageExt(bRawImg)

        If sExt = "ico" Then
            UnpackIconContainer bRawImg, tEntry
        Else
            ReDim tEntry.Res(0)
            With tEntry.Res(0)
                .DataType = resPicture
                .Size = lImgSize
                .DataBytes = bRawImg
                .FileExt = sExt
            End With
            tEntry.ResCount = 1
        End If
    End If
End Sub


Private Sub ExtractStringAt(ByVal iFile As Integer, ByVal lPos As Long, ByRef tEntry As FrxEntry)
    Dim iLen16 As Integer
    Dim lLen32 As Long
    Dim bProbe(0 To 31) As Byte
    Dim bUtf16() As Byte
    Dim sDecoded As String
    Dim lUtfOffset As Long
    Dim lLenOffset As Long

    If (LOF(iFile) - lPos + 1) < 2 Then Exit Sub

    ' look for first 32 bytes for OLE COM wide-string stream ("00 74 E8 B1")
    If (LOF(iFile) - lPos + 1) >= 28 Then
        Seek #iFile, lPos
        Get #iFile, , bProbe

        If bProbe(0) = &H0 And bProbe(1) = &H74 And bProbe(2) = &HE8 And bProbe(3) = &HB1 Then
            ' 28-byte header (RichTextBox TextRTF, etc.): length at 20 (0x14), UTF-16LE data at 28 (0x1C)
            lLenOffset = 20
            lUtfOffset = 28
        ElseIf bProbe(4) = &H0 And bProbe(5) = &H74 And bProbe(6) = &HE8 And bProbe(7) = &HB1 Then
            ' 32-byte header (OleStream) : length at 24 (0x18), UTF-16LE data at 32 (0x20)
            lLenOffset = 24
            lUtfOffset = 32
        End If

        If lUtfOffset > 0 Then
            CopyMemory lLen32, bProbe(lLenOffset), 4
            If lLen32 >= 0 And lLen32 <= (LOF(iFile) - (lPos + lUtfOffset) + 1) Then
                ReDim tEntry.Res(0)
                With tEntry.Res(0)
                    .DataType = resString
                    .FileExt = "txt"
                    If lLen32 > 0 Then
                        ReDim bUtf16(0 To lLen32 - 1)
                        Seek #iFile, lPos + lUtfOffset
                        Get #iFile, , bUtf16
                        '                        sDecoded = bUtf16
                        '                        .DataBytes = StrConv(sDecoded, vbFromUnicode)
                        sDecoded = Space$(lLen32 \ 2)
                        CopyMemory ByVal StrPtr(sDecoded), bUtf16(0), lLen32
                        .DataBytes = StrConv(sDecoded, vbFromUnicode)
                        .Size = UBound(.DataBytes) + 1
                    Else
                        .Size = 0
                    End If
                End With
                tEntry.ResCount = 1
                Exit Sub
            End If
        End If
    End If

    ' check for 4-byte Long ANSI length prefix (LabelW.Caption, standard RichTextBox.TextRTF)
    If (LOF(iFile) - lPos + 1) >= 5 Then
        Seek #iFile, lPos
        Get #iFile, , lLen32
        If lLen32 > 0 And lLen32 <= (LOF(iFile) - (lPos + 4) + 1) Then
            Dim bFirstChar As Byte
            Get #iFile, , bFirstChar
            If bFirstChar >= 9 And bFirstChar <= 126 Then
                ReDim tEntry.Res(0)
                With tEntry.Res(0)
                    .DataType = resString
                    .FileExt = "txt"
                    .Size = lLen32
                    ReDim .DataBytes(0 To lLen32 - 1)
                    Seek #iFile, lPos + 4
                    Get #iFile, , .DataBytes
                End With
                tEntry.ResCount = 1
                Exit Sub
            End If
        End If
    End If

    '  VB6 2-byte (or FF FF + 4-byte) ANSI String format
    Seek #iFile, lPos
    Get #iFile, , iLen16

    If iLen16 = -1 Then
        If (LOF(iFile) - lPos + 1) < 6 Then Exit Sub
        Get #iFile, , lLen32
    Else
        lLen32 = (CLng(iLen16) And &HFFFF&)
    End If

    If lLen32 >= 0 And lLen32 <= (LOF(iFile) - Seek(iFile) + 1) Then
        ReDim tEntry.Res(0)
        With tEntry.Res(0)
            .DataType = resString
            .FileExt = "txt"
            .Size = lLen32
            If lLen32 > 0 Then
                ReDim .DataBytes(0 To lLen32 - 1)
                Get #iFile, , .DataBytes
            End If
        End With
        tEntry.ResCount = 1
    End If
End Sub


Private Sub ExtractStringListAt(ByVal iFile As Integer, ByVal lPos As Long, ByRef tEntry As FrxEntry)
    Dim iCount As Integer
    Dim iMaxLen As Integer
    Dim i As Long
    Dim iItemLen As Integer
    Dim lStartPos As Long
    Dim lEndPos As Long
    Dim lTotalBytes As Long

    If (LOF(iFile) - lPos + 1) < 4 Then Exit Sub
    lStartPos = lPos
    Seek #iFile, lPos
    Get #iFile, , iCount
    Get #iFile, , iMaxLen

    If iCount < 0 Then Exit Sub

    ' loop thru the items to find exact byte span
    For i = 1 To iCount
        If Seek(iFile) + 1 > LOF(iFile) Then Exit Sub
        Get #iFile, , iItemLen
        If iItemLen < 0 Then Exit Sub
        Seek #iFile, Seek(iFile) + iItemLen
    Next i

    lEndPos = Seek(iFile)
    lTotalBytes = lEndPos - lStartPos

    If lTotalBytes > 0 Then
        Seek #iFile, lStartPos
        ReDim tEntry.Res(0)
        With tEntry.Res(0)
            .DataType = resStringList
            .FileExt = "txt"
            .Size = lTotalBytes
            ReDim .DataBytes(0 To lTotalBytes - 1)
            Get #iFile, , .DataBytes
        End With
        tEntry.ResCount = 1
    End If
End Sub

Private Function DecodeStringList(ByRef bData() As Byte) As String
    Dim iCount As Integer
    Dim lPos As Long
    Dim i As Long
    Dim iLen As Integer
    Dim sTemp As String
    Dim sOut() As String

    If UBound(bData) < 3 Then Exit Function
    CopyMemory iCount, bData(0), 2
    If iCount <= 0 Then DecodeStringList = "(Empty List)": Exit Function

    ReDim sOut(0 To iCount - 1)
    lPos = 4    ' start at byte 4 (skip 2-byte Count + 2-byte MaxLen)
    For i = 0 To iCount - 1
        If lPos + 1 > UBound(bData) Then Exit For
        CopyMemory iLen, bData(lPos), 2
        lPos = lPos + 2
        If iLen > 0 And (lPos + iLen - 1) <= UBound(bData) Then
            sTemp = Space$(iLen)
            CopyMemory ByVal sTemp, bData(lPos), iLen
            sOut(i) = "[" & i & "] " & sTemp
            lPos = lPos + iLen
        Else
            sOut(i) = "[" & i & "] "
        End If
    Next i

    DecodeStringList = Join(sOut, vbCrLf)
End Function


Private Sub ExtractItemDataAt(ByVal iFile As Integer, ByVal lPos As Long, ByRef tEntry As FrxEntry)
    Dim iCount As Integer
    Dim lTotalBytes As Long

    If (LOF(iFile) - lPos + 1) < 2 Then Exit Sub
    Seek #iFile, lPos
    Get #iFile, , iCount
    If iCount < 0 Then Exit Sub

    ' 2-byte count + (iCount * 4-byte Long values)
    lTotalBytes = 2 + (CLng(iCount) * 4)
    If lTotalBytes <= (LOF(iFile) - lPos + 1) Then
        Seek #iFile, lPos
        ReDim tEntry.Res(0)
        With tEntry.Res(0)
            .DataType = resItemData
            .FileExt = "txt"
            .Size = lTotalBytes
            ReDim .DataBytes(0 To lTotalBytes - 1)
            Get #iFile, , .DataBytes
        End With
        tEntry.ResCount = 1
    End If
End Sub

Private Sub ExtractCustomBlobAt(ByVal iFile As Integer, ByVal lPos As Long, ByRef tEntry As FrxEntry)
    Dim lBlobSize As Long
    Dim bBlob() As Byte
    Dim i As Long
    Dim lImgSize As Long
    Dim bRawImg() As Byte
    Dim sExt As String

    If (LOF(iFile) - lPos + 1) < 4 Then Exit Sub
    Seek #iFile, lPos
    Get #iFile, , lBlobSize

    If lBlobSize <= 0 Or lBlobSize > (LOF(iFile) - lPos - 3) Then Exit Sub

    ReDim bBlob(0 To lBlobSize - 1)
    Get #iFile, , bBlob

    ' scan inside the ActiveX blob for embedded "lt\0\0" images (ImageLists, Toolbars, etc.)
    For i = 0 To lBlobSize - 8
        If bBlob(i) = &H6C And bBlob(i + 1) = &H74 And bBlob(i + 2) = &H0 And bBlob(i + 3) = &H0 Then
            CopyMemory lImgSize, bBlob(i + 4), 4
            If lImgSize > 0 And (i + 8 + lImgSize) <= lBlobSize Then
                ReDim bRawImg(0 To lImgSize - 1)
                CopyMemory bRawImg(0), bBlob(i + 8), lImgSize
                sExt = DetectImageExt(bRawImg)

                If sExt = "ico" Then
                    UnpackIconContainer bRawImg, tEntry
                Else
                    ReDim Preserve tEntry.Res(tEntry.ResCount)
                    With tEntry.Res(tEntry.ResCount)
                        .DataType = resPicture
                        .Size = lImgSize
                        .DataBytes = bRawImg
                        .FileExt = sExt
                    End With
                    tEntry.ResCount = tEntry.ResCount + 1
                End If

                i = i + 8 + lImgSize - 1
            End If
        End If
    Next i

    ' if no embedded pictures found, keep raw PropertyBag blob as resCustomBlob
    If tEntry.ResCount = 0 Then
        Dim sBagCheck As String
        ReDim tEntry.Res(0)
        With tEntry.Res(0)
            .DataType = resCustomBlob
            .Size = lBlobSize
            .DataBytes = bBlob
            If TryDecodePropertyBag(.DataBytes, sBagCheck) Then
                .FileExt = "txt"
            Else
                .FileExt = DetectBlobExt(.DataBytes)
            End If
        End With
        tEntry.ResCount = 1
    End If

End Sub


Public Sub AddRawGapEntry(ByVal lFrxIdx As Long, ByRef bFile() As Byte, ByVal lStart As Long, ByVal lLen As Long, Optional ByVal sSourceName As String = "(Standalone FRX)")
    ReDim Preserve mAllEntries(mEntryCount)
    With mAllEntries(mEntryCount)
        .FrxIndex = lFrxIdx
        .SourceForm = sSourceName
        .ControlName = "DataChunk_" & mEntryCount
        .PropName = "UnindexedData"
        .Offset = lStart
        .ResCount = 1
        ReDim .Res(0)
        .Res(0).DataType = resCustomBlob
        .Res(0).FileExt = "bin"
        .Res(0).Size = lLen
        ReDim .Res(0).DataBytes(0 To lLen - 1)
        CopyMemory .Res(0).DataBytes(0), bFile(lStart), lLen
    End With
    mEntryCount = mEntryCount + 1
End Sub

Public Function FormatResourcePreview(ByRef tRes As FrxResource) As String

    Dim sBagText As String
    Dim sBanner As String

    Select Case tRes.DataType
        Case resString
            If tRes.Size = 0 Then
                FormatResourcePreview = "(Empty String)"
            Else
                FormatResourcePreview = Replace$(StrConv(tRes.DataBytes, vbUnicode), vbNullChar, "")
            End If

        Case resStringList
            FormatResourcePreview = DecodeStringList(tRes.DataBytes)

        Case resItemData
            FormatResourcePreview = DecodeItemData(tRes.DataBytes)

        Case resCustomBlob
            If TryDecodePropertyBag(tRes.DataBytes, sBagText) Then
                FormatResourcePreview = sBagText
            Else
                Select Case LCase$(tRes.FileExt)
                    Case "effects"
                        sBanner = "=== LaVolpe AlphaImageControl Effects object (" & tRes.Size & " bytes) ===" & vbCrLf & vbCrLf
                    Case "avi"
                        sBanner = "=== Embedded RIFF AVI Video (" & tRes.Size & " bytes) — Click 'Export File' to save as .avi ===" & vbCrLf & vbCrLf
                    Case "wav"
                        sBanner = "=== Embedded RIFF WAVE Audio (" & tRes.Size & " bytes) — Click 'Export File' to save as .wav ===" & vbCrLf & vbCrLf
                    Case "mid"
                        sBanner = "=== Embedded MIDI Audio (" & tRes.Size & " bytes) — Click 'Export File' to save as .mid ===" & vbCrLf & vbCrLf
                    Case "ole"
                        sBanner = "=== OLE Compound / Structured Storage (" & tRes.Size & " bytes) ===" & vbCrLf & vbCrLf
                    Case "ils"
                        sBanner = "=== comctl32 ImageList Stream (" & tRes.Size & " bytes) ===" & vbCrLf & vbCrLf
                    Case Else
                        sBanner = "=== Raw Binary Blob (" & tRes.Size & " bytes) ===" & vbCrLf & vbCrLf
                End Select
                FormatResourcePreview = sBanner & FormatHexDump(tRes.DataBytes, 4096)
            End If
    End Select
End Function

Public Function DetectImageExt(ByRef bData() As Byte) As String
    Dim lLen As Long
    On Error Resume Next
    lLen = UBound(bData) + 1
    If Err.Number <> 0 Or lLen < 4 Then
        DetectImageExt = "bin"
        Exit Function
    End If
    On Error GoTo 0

    If bData(0) = &H42 And bData(1) = &H4D Then DetectImageExt = "bmp": Exit Function
    If bData(0) = &H47 And bData(1) = &H49 And bData(2) = &H46 Then DetectImageExt = "gif": Exit Function
    If bData(0) = &HFF And bData(1) = &HD8 Then DetectImageExt = "jpg": Exit Function
    If bData(0) = &H89 And bData(1) = &H50 And bData(2) = &H4E And bData(3) = &H47 Then DetectImageExt = "png": Exit Function
    If bData(0) = &H0 And bData(1) = &H0 And bData(2) = &H1 And bData(3) = &H0 Then DetectImageExt = "ico": Exit Function
    If bData(0) = &H0 And bData(1) = &H0 And bData(2) = &H2 And bData(3) = &H0 Then DetectImageExt = "cur": Exit Function
    If bData(0) = &HD7 And bData(1) = &HCD And bData(2) = &HC6 And bData(3) = &H9A Then DetectImageExt = "wmf": Exit Function
    If bData(0) = &H1 And bData(1) = &H0 And bData(2) = &H0 And bData(3) = &H0 Then DetectImageExt = "emf": Exit Function

    DetectImageExt = "bin"
End Function

Public Function DetectBlobExt(ByRef bData() As Byte) As String
    Dim lLen As Long

    On Error Resume Next
    lLen = UBound(bData) + 1
    If Err.Number <> 0 Or lLen < 4 Then
        DetectBlobExt = "bin"
        Exit Function
    End If
    On Error GoTo 0

    If lLen >= 16 And bData(0) = &H0 And bData(1) = &H74 And bData(2) = &HE8 And bData(3) = &HB1 Then
        ' Small OLE persistence blob (common for LaVolpe AlphaImageControl .Effects)
        If lLen <= 64 Then
            DetectBlobExt = "effects"
            Exit Function
        End If
    End If

    'RIFF containers ("RIFF" = 52 49 46 46): AVI or WAV
    If lLen >= 12 And bData(0) = &H52 And bData(1) = &H49 And bData(2) = &H46 And bData(3) = &H46 Then
        If bData(8) = &H41 And bData(9) = &H56 And bData(10) = &H49 And bData(11) = &H20 Then
            DetectBlobExt = "avi"
            Exit Function
        ElseIf bData(8) = &H57 And bData(9) = &H41 And bData(10) = &H56 And bData(11) = &H45 Then
            DetectBlobExt = "wav"
            Exit Function
        End If
    End If

    ' OLE Compound File / Structured Storage (D0 CF 11 E0 A1 B1 1A E1)
    If lLen >= 8 And bData(0) = &HD0 And bData(1) = &HCF And bData(2) = &H11 And bData(3) = &HE0 Then
        DetectBlobExt = "ole"
        Exit Function
    End If

    'MIDI audio ("MThd" = 4D 54 68 64)
    If bData(0) = &H4D And bData(1) = &H54 And bData(2) = &H68 And bData(3) = &H64 Then
        DetectBlobExt = "mid"
        Exit Function
    End If

    'comctl32 5.0 ImageList stream ("IL\x01\x00" = 49 4C 01 00)
    If bData(0) = &H49 And bData(1) = &H4C And bData(2) = &H1 And bData(3) = &H0 Then
        DetectBlobExt = "ils"
        Exit Function
    End If

    DetectBlobExt = "bin"
End Function

Public Sub UnpackIconContainer(ByRef bIcoData() As Byte, ByRef tEntry As FrxEntry)
    Dim iCount As Integer
    Dim i As Long
    Dim lDirOffset As Long
    Dim lImgSize As Long
    Dim lImgOffset As Long
    Dim bSingleIco() As Byte

    If UBound(bIcoData) < 6 Then Exit Sub
    CopyMemory iCount, bIcoData(4), 2
    If iCount <= 0 Or iCount > 256 Then Exit Sub

    For i = 0 To iCount - 1
        lDirOffset = 6 + (i * 16)
        If lDirOffset + 15 > UBound(bIcoData) Then Exit For

        CopyMemory lImgSize, bIcoData(lDirOffset + 8), 4
        CopyMemory lImgOffset, bIcoData(lDirOffset + 12), 4

        If lImgSize > 0 And (lImgOffset + lImgSize) <= (UBound(bIcoData) + 1) Then
            ' build a standalone .ico file with 1 icon: [6-byte hdr][16-byte dir][image bytes]
            ReDim bSingleIco(0 To 21 + lImgSize)
            bSingleIco(0) = 0: bSingleIco(1) = 0
            bSingleIco(2) = 1: bSingleIco(3) = 0    ' Type 1 = Icon
            bSingleIco(4) = 1: bSingleIco(5) = 0    ' Count = 1

            ' copy the 16-byte ICONDIRENTRY and patch its data offset to 22 (0x16)
            CopyMemory bSingleIco(6), bIcoData(lDirOffset), 16
            Dim lNewOffset As Long
            lNewOffset = 22
            CopyMemory bSingleIco(18), lNewOffset, 4

            ' copy the raw icon frame bytes
            CopyMemory bSingleIco(22), bIcoData(lImgOffset), lImgSize

            ReDim Preserve tEntry.Res(tEntry.ResCount)
            With tEntry.Res(tEntry.ResCount)
                .DataType = resPicture
                .FileExt = "ico"
                .Size = UBound(bSingleIco) + 1
                .DataBytes = bSingleIco
            End With
            tEntry.ResCount = tEntry.ResCount + 1
        End If
    Next i
End Sub

'searches for a reference position in a binary file
Public Function FindBinaryRefPos(ByVal sLine As String) As Long
    Dim lPos As Long
    lPos = InStr(1, sLine, ".frx\"":", vbTextCompare)
    If lPos = 0 Then lPos = InStr(1, sLine, ".ctx\"":", vbTextCompare)
    If lPos = 0 Then lPos = InStr(1, sLine, ".pgx\"":", vbTextCompare)
    If lPos = 0 Then lPos = InStr(1, sLine, ".dsx\"":", vbTextCompare)
    FindBinaryRefPos = lPos
End Function




Public Sub DebugDumpOffset(ByVal sFrxPath As String, ByVal lOffset As Long, Optional ByVal lBytes As Long = 96)
    Dim iFile As Integer
    Dim bBuf() As Byte
    Dim lRead As Long

    iFile = FreeFile
    Open sFrxPath For Binary Access Read As #iFile
    lRead = LOF(iFile) - lOffset
    If lRead > lBytes Then lRead = lBytes
    If lRead > 0 Then
        ReDim bBuf(0 To lRead - 1)
        Seek #iFile, lOffset + 1
        Get #iFile, , bBuf
        Debug.Print "=== " & FileGetName(sFrxPath) & " @ 0x" & Hex$(lOffset) & " ==="
        Debug.Print FormatHexDump(bBuf, lRead)
    End If
    Close #iFile
End Sub


Public Function TryDecodePropertyBag(ByRef bData() As Byte, ByRef sOutText As String) As Boolean
    Dim lSize As Long
    Dim lPos As Long
    Dim iVarType As Integer
    Dim iNameChars As Integer
    Dim lEntryBytes As Long
    Dim lNameBytes As Long
    Dim lValOffset As Long
    Dim sPropName As String
    Dim sValStr As String
    Dim sLines() As String
    Dim lCount As Long
    Dim i As Long

    On Error Resume Next
    lSize = UBound(bData) + 1
    If Err.Number <> 0 Or lSize < 36 Then Exit Function
    On Error GoTo 0

    ' locate the VB6 PropertyBag signature "93 B2 00 00" (usually at offset 0 or offset 28)
    lPos = -1
    For i = 0 To 32 Step 4
        If i + 8 <= lSize Then
            If bData(i) = &H93 And bData(i + 1) = &HB2 And bData(i + 2) = &H0 And bData(i + 3) = &H0 Then
                lPos = i + 8    ' skip 4-byte signature + 4-byte bag size
                Exit For
            End If
        End If
    Next i

    If lPos < 0 Then Exit Function

    ReDim sLines(0 To 63)
    sLines(0) = "=== Serialized VB6 PropertyBag ==="
    lCount = 1

    Do While lPos + 12 <= lSize
        CopyMemory iVarType, bData(lPos), 2
        CopyMemory iNameChars, bData(lPos + 2), 2
        CopyMemory lEntryBytes, bData(lPos + 8), 4

        If iNameChars <= 0 Or lEntryBytes < 12 Or (lPos + lEntryBytes) > lSize Then Exit Do

        lNameBytes = CLng(iNameChars) * 2
        lValOffset = lPos + 12 + lNameBytes
        If lValOffset > lPos + lEntryBytes Then Exit Do

        ' Extract utf16 property name
        sPropName = Space$(iNameChars)
        CopyMemory ByVal StrPtr(sPropName), bData(lPos + 12), lNameBytes

        ' decode the typed value based on VARTYPE
        sValStr = DecodeVariantBytes(iVarType, bData, lValOffset, (lPos + lEntryBytes) - lValOffset)

        If lCount > UBound(sLines) Then ReDim Preserve sLines(0 To UBound(sLines) + 64)
        sLines(lCount) = sPropName & " = " & sValStr
        lCount = lCount + 1

        lPos = lPos + lEntryBytes
    Loop

    If lCount > 1 Then
        ReDim Preserve sLines(0 To lCount - 1)
        sOutText = Join(sLines, vbCrLf)
        TryDecodePropertyBag = True
    End If
End Function

Private Function DecodeVariantBytes(ByVal iVarType As Integer, ByRef bData() As Byte, ByVal lOffset As Long, ByVal lAvail As Long) As String
    Dim iVal16 As Integer
    Dim lVal32 As Long
    Dim sngVal As Single
    Dim dblVal As Double
    Dim lStrBytes As Long
    Dim sStr As String

    Select Case iVarType
        Case 2    ' vbInteger
            If lAvail >= 2 Then
                CopyMemory iVal16, bData(lOffset), 2
                DecodeVariantBytes = CStr(iVal16)
            End If

        Case 3, 19    ' vbLong / vbULong / OLE_COLOR
            If lAvail >= 4 Then
                CopyMemory lVal32, bData(lOffset), 4
                DecodeVariantBytes = CStr(lVal32) & "  (&H" & Right$("00000000" & Hex$(lVal32), 8) & ")"
            End If

        Case 4    ' vbSingle
            If lAvail >= 4 Then
                CopyMemory sngVal, bData(lOffset), 4
                DecodeVariantBytes = CStr(sngVal)
            End If

        Case 5    ' vbDouble
            If lAvail >= 8 Then
                CopyMemory dblVal, bData(lOffset), 8
                DecodeVariantBytes = CStr(dblVal)
            End If

        Case 8    ' vbString (4-byte byte-length prefix + UTF16 characters)
            If lAvail >= 4 Then
                CopyMemory lStrBytes, bData(lOffset), 4
                If lStrBytes > 0 And (lOffset + 4 + lStrBytes) <= UBound(bData) + 1 Then
                    sStr = Space$(lStrBytes \ 2)
                    CopyMemory ByVal StrPtr(sStr), bData(lOffset + 4), lStrBytes
                    DecodeVariantBytes = """" & sStr & """"
                Else
                    DecodeVariantBytes = """"""
                End If
            End If

        Case 9    ' vbObject (IDispatch/StdPicture in StatusBar/Toolbar PropertyBag)
            DecodeVariantBytes = "[COM Object / Picture | " & lAvail & " bytes]"

        Case 11    ' vbBoolean
            If lAvail >= 2 Then
                CopyMemory iVal16, bData(lOffset), 2
                If iVal16 <> 0 Then
                    DecodeVariantBytes = "True"
                Else
                    DecodeVariantBytes = "False"
                End If
            End If

        Case 17    ' vbByte
            If lAvail >= 1 Then DecodeVariantBytes = CStr(bData(lOffset))

        Case 8209    ' vbArray + vbByte (0x2011: [2-byte Dims][4-byte Count][4-byte LBound][Bytes])
            Dim iDims As Integer
            If lAvail >= 2 Then
                CopyMemory iDims, bData(lOffset), 2
                If iDims = 0 Then
                    DecodeVariantBytes = """"""
                ElseIf iDims = 1 And lAvail >= 10 Then
                    CopyMemory lStrBytes, bData(lOffset + 2), 4
                    If lStrBytes = 0 Then
                        DecodeVariantBytes = """"""
                    ElseIf lStrBytes > 0 And (lOffset + 10 + lStrBytes) <= UBound(bData) + 1 Then
                        ' Check if payload is utf16 text (even length, printable first char, null high byte)
                        If (lStrBytes Mod 2 = 0) And bData(lOffset + 10) >= 32 And bData(lOffset + 10) <= 126 And bData(lOffset + 11) = 0 Then
                            sStr = Space$(lStrBytes \ 2)
                            CopyMemory ByVal StrPtr(sStr), bData(lOffset + 10), lStrBytes
                            DecodeVariantBytes = """" & sStr & """"
                        Else
                            DecodeVariantBytes = "[ByteArray | " & lStrBytes & " bytes]"
                        End If
                    Else
                        DecodeVariantBytes = "[ByteArray | " & lAvail & " bytes]"
                    End If
                Else
                    DecodeVariantBytes = "[ByteArray | " & lAvail & " bytes]"
                End If
            End If


        Case Else
            DecodeVariantBytes = "[Type " & iVarType & " | " & lAvail & " bytes]"
    End Select
End Function



Public Function SaveResourceToFile(ByRef tRes As FrxResource, ByVal sDestPath As String) As Boolean
    Dim iFile As Integer
    Dim sTextOut As String
    Dim bTextBytes() As Byte

    On Error GoTo ErrHandler
    If Dir$(sDestPath) <> "" Then Kill sDestPath

    Select Case tRes.DataType
        Case resPicture
            ' raw image file bytes (.png, .ico, .bmp, .gif, .jpg, .cur, .wmf, .emf)
            iFile = FreeFile
            Open sDestPath For Binary Access Write As #iFile
            Put #iFile, , tRes.DataBytes
            Close #iFile
            SaveResourceToFile = True

        Case resString, resStringList, resItemData
            '  decoded text / list / RTF markup
            sTextOut = FormatResourcePreview(tRes)
            If Len(sTextOut) > 0 Then
                bTextBytes = StrConv(sTextOut, vbFromUnicode)
                iFile = FreeFile
                Open sDestPath For Binary Access Write As #iFile
                Put #iFile, , bTextBytes
                Close #iFile
            Else
                ' create empty 0 byte file
                iFile = FreeFile
                Open sDestPath For Binary Access Write As #iFile
                Close #iFile
            End If
            SaveResourceToFile = True

        Case resCustomBlob
            ' If it's a decoded PropertyBag, export readable Key=Value text, otherwise raw .bin
            If TryDecodePropertyBag(tRes.DataBytes, sTextOut) Then
                bTextBytes = StrConv(sTextOut, vbFromUnicode)
                iFile = FreeFile
                Open sDestPath For Binary Access Write As #iFile
                Put #iFile, , bTextBytes
                Close #iFile
            Else
                iFile = FreeFile
                Open sDestPath For Binary Access Write As #iFile
                Put #iFile, , tRes.DataBytes
                Close #iFile
            End If
            SaveResourceToFile = True
    End Select
    Exit Function

ErrHandler:
    If iFile <> 0 Then Close #iFile
    SaveResourceToFile = False
End Function

Public Function BuildResourceFileName(ByRef tEntry As FrxEntry, ByVal lResIdx As Long) As String
    Dim sBase As String
    Dim sExt As String

    sBase = Replace$(tEntry.ControlName, "(", "_")
    sBase = Replace$(sBase, ")", "")
    sBase = sBase & "_" & tEntry.PropName

    If tEntry.ResCount > 1 Then
        sBase = sBase & "_" & Format$(lResIdx + 1, "00")
    End If

    sExt = tEntry.Res(lResIdx).FileExt
    If Len(sExt) = 0 Then sExt = "bin"

    ' if it's a PropertyBag, default to .txt instead of .bin
    Dim sDummy As String
    If tEntry.Res(lResIdx).DataType = resCustomBlob Then
        If TryDecodePropertyBag(tEntry.Res(lResIdx).DataBytes, sDummy) Then sExt = "txt"
    End If

    BuildResourceFileName = SanitizeFileName(sBase) & "." & sExt
End Function

Public Function SanitizeFileName(ByVal sName As String) As String
    Dim vBad As Variant
    Dim vChar As Variant
    vBad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each vChar In vBad
        sName = Replace$(sName, CStr(vChar), "_")
    Next vChar
    SanitizeFileName = sName
End Function

Public Function PackEntryToMultiIco(ByRef tEntry As FrxEntry, ByVal sOutIcoPath As String) As Boolean
    Dim i As Long
    Dim lCount As Long
    Dim bFramePayloads() As Variant
    Dim bDirEntries() As Byte
    Dim lTotalPayloadBytes As Long
    Dim lHeaderBytes As Long
    Dim lCurOffset As Long
    Dim bIcoOut() As Byte
    Dim iFile As Integer

    lCount = tEntry.ResCount
    If lCount <= 0 Then Exit Function

    ReDim bFramePayloads(0 To lCount - 1)
    ReDim bDirEntries(0 To (lCount * 16) - 1)

    ' extract the raw frame and 16 byte ICONDIRENTRY from each unpacked .ico item
    For i = 0 To lCount - 1
        If tEntry.Res(i).DataType <> resPicture Or LCase$(tEntry.Res(i).FileExt) <> "ico" Then
            Exit Function
        End If

        Dim bSrc() As Byte
        Dim lFrameLen As Long
        Dim bPayload() As Byte

        bSrc = tEntry.Res(i).DataBytes
        If UBound(bSrc) < 22 Then Exit Function

        'copy the 16-byte ICONDIRENTRY (bytes 6..21 of the raw icon)
        CopyMemory bDirEntries(i * 16), bSrc(6), 16

        'copy the raw frame data (bytes 22..end - either BITMAPINFOHEADER or 256x256 PNG)
        lFrameLen = (UBound(bSrc) + 1) - 22
        ReDim bPayload(0 To lFrameLen - 1)
        CopyMemory bPayload(0), bSrc(22), lFrameLen

        bFramePayloads(i) = bPayload
        lTotalPayloadBytes = lTotalPayloadBytes + lFrameLen
    Next i

    ' combine multi-resolution .ico buffer
    lHeaderBytes = 6 + (16 * lCount)
    ReDim bIcoOut(0 To lHeaderBytes + lTotalPayloadBytes - 1)

    bIcoOut(0) = 0: bIcoOut(1) = 0
    bIcoOut(2) = 1: bIcoOut(3) = 0    ' type 1 = Icon
    Dim iFrameCount As Integer
    iFrameCount = CInt(lCount)
    CopyMemory bIcoOut(4), iFrameCount, 2

    lCurOffset = lHeaderBytes
    For i = 0 To lCount - 1
        Dim bChunk() As Byte
        Dim lChunkLen As Long
        Dim lDirPos As Long

        bChunk = bFramePayloads(i)
        lChunkLen = UBound(bChunk) + 1
        lDirPos = 6 + (i * 16)

        ' write 16-byte directory entry and patch new data offset
        CopyMemory bIcoOut(lDirPos), bDirEntries(i * 16), 16
        CopyMemory bIcoOut(lDirPos + 12), lCurOffset, 4

        ' write the data
        CopyMemory bIcoOut(lCurOffset), bChunk(0), lChunkLen
        lCurOffset = lCurOffset + lChunkLen
    Next i

    ' save to disk
    If FileExist(sOutIcoPath) Then FileKill sOutIcoPath
    iFile = FreeFile
    Open sOutIcoPath For Binary Access Write As #iFile
    Put #iFile, , bIcoOut
    Close #iFile

    PackEntryToMultiIco = True
End Function
