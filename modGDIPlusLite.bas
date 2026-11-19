Attribute VB_Name = "modGDIPlusLite"
Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC_PIC
    CbSizeOfStruct As Long
    PicType As Long
    HgdiObj As Long
    HpalOrXyExt As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal Stream As Long, BITMAP As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Format As Long, ByVal Scan0 As Long, BITMAP As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFileName As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long

Private Declare Function OleCreatePictureIndirect_PIC Lib "OleAut32.dll" Alias "OleCreatePictureIndirect" (ByRef lpPictDesc As PICTDESC_PIC, riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long

Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Boolean
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private m_GdipToken As Long
Private m_IPictureInit As Boolean
Private m_IPictureIID As GUID

Public Function InitGdiplus() As Boolean
    Dim tSI As GdiplusStartupInput
    If m_GdipToken <> 0 Then
        InitGdiplus = True
        Exit Function
    End If
    tSI.GdiplusVersion = 1
    If GdiplusStartup(m_GdipToken, tSI) = 0 Then
        InitGdiplus = True
    End If
End Function

Public Sub ShutdownGdiplus()
    If m_GdipToken <> 0 Then
        GdiplusShutdown m_GdipToken
        m_GdipToken = 0
    End If
End Sub

Public Function LoadPictureFromBytes(ByRef ImageData() As Byte, _
                                     Optional ByVal imgType As PictureTypeConstants = vbPicTypeBitmap, _
                                     Optional ByVal BackgroundColor As Long = &H0) As StdPicture
    Dim lBitmap As Long
    Dim hBitmap As Long
    Dim argbColor As Long

    If Not InitGdiplus() Then Exit Function

    If (BackgroundColor And &H80000000) Then
        BackgroundColor = GetSysColor(BackgroundColor And &HFFFFFF)
    End If
    argbColor = AlphaRGBColor(BackgroundColor)

    If GDIPlusBitmapFromBytes(ImageData, lBitmap) = 0 Then
        If GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, argbColor) = 0 Then
            Set LoadPictureFromBytes = PictureFromBitmap(hBitmap, imgType, 0)
        End If
        If lBitmap <> 0 Then GdipDisposeImage lBitmap
    End If
End Function

Public Function ExportEntryAsHorizontalStrip(ByRef tEntry As FrxEntry, ByVal sDestPngPath As String) As Boolean
    Dim i As Long
    Dim lCount As Long
    Dim lBitmaps() As Long
    Dim lWidths() As Long
    Dim lHeights() As Long
    Dim lTotalW As Long
    Dim lMaxH As Long
    Dim lStripBmp As Long
    Dim hGfx As Long
    Dim lCurX As Long
    Dim tPngClsid As GUID
    
    lCount = tEntry.ResCount
    If lCount <= 0 Then Exit Function
    If Not InitGdiplus() Then Exit Function
    
    ReDim lBitmaps(0 To lCount - 1)
    ReDim lWidths(0 To lCount - 1)
    ReDim lHeights(0 To lCount - 1)
    
    ' 1. Decode each frame and measure total strip width & max height
    For i = 0 To lCount - 1
        If tEntry.Res(i).DataType = resPicture Then
            If GDIPlusBitmapFromBytes(tEntry.Res(i).DataBytes, lBitmaps(i)) = 0 Then
                GdipGetImageWidth lBitmaps(i), lWidths(i)
                GdipGetImageHeight lBitmaps(i), lHeights(i)
                lTotalW = lTotalW + lWidths(i)
                If lHeights(i) > lMaxH Then lMaxH = lHeights(i)
            End If
        End If
    Next i
    
    If lTotalW <= 0 Or lMaxH <= 0 Then GoTo Cleanup
    
    ' 2. Create 32-bit ARGB strip canvas and draw frames side-by-side at native pixel sizes
    If GdipCreateBitmapFromScan0(lTotalW, lMaxH, 0, &H26200A, 0, lStripBmp) = 0 Then
        If GdipGetImageGraphicsContext(lStripBmp, hGfx) = 0 Then
            lCurX = 0
            For i = 0 To lCount - 1
                If lBitmaps(i) <> 0 Then
                    GdipDrawImageRectI hGfx, lBitmaps(i), lCurX, 0, lWidths(i), lHeights(i)
                    lCurX = lCurX + lWidths(i)
                End If
            Next i
            GdipDeleteGraphics hGfx
        End If
        
        If GetPngEncoderClsid(tPngClsid) Then
            If Dir$(sDestPngPath) <> "" Then Kill sDestPngPath
            ExportEntryAsHorizontalStrip = (GdipSaveImageToFile(lStripBmp, StrPtr(sDestPngPath), tPngClsid, ByVal 0&) = 0)
        End If
        GdipDisposeImage lStripBmp
    End If
    
Cleanup:
    For i = 0 To lCount - 1
        If lBitmaps(i) <> 0 Then GdipDisposeImage lBitmaps(i)
    Next i
End Function

Private Function GDIPlusBitmapFromBytes(ByRef ImageData() As Byte, ByRef outBitmap As Long) As Long
    Dim hMem As Long
    Dim pMem As Long
    Dim oStream As Object
    Dim lSize As Long

    On Error Resume Next
    lSize = UBound(ImageData) - LBound(ImageData) + 1
    If Err.Number <> 0 Or lSize <= 0 Then
        GDIPlusBitmapFromBytes = 2
        Exit Function
    End If
    On Error GoTo 0

    hMem = GlobalAlloc(&H42, lSize)
    If hMem = 0 Then
        GDIPlusBitmapFromBytes = 3
        Exit Function
    End If

    pMem = GlobalLock(hMem)
    If pMem <> 0 Then
        CopyMemory ByVal pMem, ImageData(LBound(ImageData)), lSize
        GlobalUnlock hMem
    Else
        GDIPlusBitmapFromBytes = 1
        Exit Function
    End If

    If CreateStreamOnHGlobal(hMem, 1, oStream) = 0 Then
        GDIPlusBitmapFromBytes = GdipCreateBitmapFromStream(ObjPtr(oStream), outBitmap)
    Else
        GDIPlusBitmapFromBytes = 1
    End If
    Set oStream = Nothing
End Function

Private Function GetPngEncoderClsid(ByRef outClsid As GUID) As Boolean
    Dim numEncoders As Long
    Dim lSize As Long
    Dim i As Long
    Dim buf() As Byte
    Dim pEncoders As Long
    Dim pMime As Long
    Dim lLen As Long
    Dim sMime As String

    If GdipGetImageEncodersSize(numEncoders, lSize) = 0 And lSize > 0 Then
        ReDim buf(0 To lSize - 1)
        pEncoders = VarPtr(buf(0))
        If GdipGetImageEncoders(numEncoders, lSize, ByVal pEncoders) = 0 Then
            For i = 0 To numEncoders - 1
                CopyMemory pMime, ByVal (pEncoders + (i * 76) + 48), 4
                lLen = lstrlenW(pMime)
                If lLen > 0 Then
                    sMime = Space$(lLen)
                    CopyMemory ByVal StrPtr(sMime), ByVal pMime, lLen * 2
                    If LCase$(sMime) = "image/png" Then
                        CopyMemory outClsid, ByVal (pEncoders + (i * 76)), 16
                        GetPngEncoderClsid = True
                        Exit Function
                    End If
                End If
            Next i
        End If
    End If
End Function

Private Function PictureFromBitmap(ByVal hGDIHandle As Long, ByVal ePicType As PictureTypeConstants, Optional ByVal hPallette As Long = 0) As StdPicture
    Dim tPictDesc As PICTDESC_PIC

    With tPictDesc
        .CbSizeOfStruct = Len(tPictDesc)
        .PicType = ePicType
        .HgdiObj = hGDIHandle
        .HpalOrXyExt = hPallette
    End With

    If Not m_IPictureInit Then InitIPicture

    If OleCreatePictureIndirect_PIC(tPictDesc, m_IPictureIID, 1&, PictureFromBitmap) < 0 Then
        DeleteObject hGDIHandle
        Set PictureFromBitmap = Nothing
    End If
End Function

Private Function AlphaRGBColor(ByVal RGBColor As Long) As Long
    Dim R As Long, G As Long, B As Long
    R = RGBColor And &HFF&
    G = (RGBColor And &HFF00&) \ &H100&
    B = (RGBColor And &HFF0000) \ &H10000
    AlphaRGBColor = &HFF000000 Or (R * &H10000) Or (G * &H100&) Or B
End Function

Private Sub InitIPicture()
    With m_IPictureIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
        .Data4(4) = &H0: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With
    m_IPictureInit = True
End Sub
