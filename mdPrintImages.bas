Attribute VB_Name = "mdPrintImages"
'=========================================================================
'
' vbimg2pdf (c) 2018 by wqweto@gmail.com
'
' Convert jpeg/png images to multi-page pdf file
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for GetDeviceCaps
Private Const HORZRES                       As Long = 8
Private Const VERTRES                       As Long = 10
Private Const LOGPIXELSX                    As Long = 88
Private Const LOGPIXELSY                    As Long = 90
'--- for BITMAPINFOHEADER
Private Const BI_JPEG                       As Long = 4
Private Const BI_PNG                        As Long = 5
'--- for DocumentProperties
Private Const DM_OUT_BUFFER                 As Long = 2
Private Const DM_IN_BUFFER                  As Long = 8
Private Const IDOK                          As Long = 1
Private Const DM_ORIENTATION                As Long = &H1
Private Const DM_PAPERSIZE                  As Long = &H2
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, ByRef DOCINFO As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputBuf As Any, Optional ByVal outputBuf As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef nWidth As Single, ByRef nHeight As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Type DOCINFO
    cbSize              As Long
    lpszDocName         As String
    lpszOutput          As String
End Type

Private Type DEVMODE
    dmDeviceName        As String * 32
    dmSpecVersion       As Integer
    dmDriverVersion     As Integer
    dmSize              As Integer
    dmDriverExtra       As Integer
    dmFields            As Long
    dmOrientation       As Integer
    dmPaperSize         As Integer
    dmPaperLength       As Integer
    dmPaperWidth        As Integer
    dmScale             As Integer
    dmCopies            As Integer
    dmDefaultSource     As Integer
    dmPrintQuality      As Integer
    dmColor             As Integer
    dmDuplex            As Integer
    dmYResolution       As Integer
    dmTTOption          As Integer
    dmCollate           As Integer
    dmFormName          As String * 32
    dmLogPixels         As Integer
    dmBitsPerPel        As Long
    dmPelsWidth         As Long
    dmPelsHeight        As Long
    dmDisplayFlags      As Long
    dmDisplayFrequency  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function PrintImages( _
            sPrinterName As String, _
            vInputFiles As Variant, _
            Optional sOutputFile As String, _
            Optional ByVal lPaperSize As Long, _
            Optional ByVal lOrientation As Long, _
            Optional vMargins As Variant, _
            Optional sError As String) As Boolean
    Dim baDevMode()     As Byte
    Dim hDC             As Long
    Dim uInfo           As DOCINFO
    Dim lDpiX           As Long
    Dim lDpiY           As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lIdx            As Long
    Dim uHeader         As BITMAPINFOHEADER
    Dim baImage()       As Byte
    Dim lTargetX        As Long
    Dim lTargetY        As Long
    Dim lTargetW         As Long
    Dim lTargetH         As Long
    
    On Error GoTo EH
    '--- will use GDI+ to retrieve input images dimensions
    If Not StartGdip() Then
        GoTo QH
    End If
    '--- setup printer paper size/orientation
    If Not SetupDevMode(sPrinterName, lPaperSize, lOrientation, baDevMode, sError) Then
        GoTo QH
    End If
    '--- setup output file
    hDC = CreateDC("", sPrinterName, 0, baDevMode(0))
    If hDC = 0 Then
        sError = GetSystemMessage(Err.LastDllError)
        GoTo QH
    End If
    uInfo.cbSize = LenB(uInfo)
    uInfo.lpszDocName = App.ProductName & " - PrintImages"
    If LenB(sOutputFile) <> 0 Then
        uInfo.lpszOutput = CanonicalPath(sOutputFile)
        Call DeleteFile(uInfo.lpszOutput)
    End If
    '--- setup printable area
    lDpiX = GetDeviceCaps(hDC, LOGPIXELSX)
    lDpiY = GetDeviceCaps(hDC, LOGPIXELSY)
    lLeft = C_Dbl(At(vMargins, 0)) * lDpiX
    lTop = C_Dbl(At(vMargins, 1)) * lDpiY
    lWidth = GetDeviceCaps(hDC, HORZRES) - lLeft - C_Dbl(At(vMargins, 2)) * lDpiX
    lHeight = GetDeviceCaps(hDC, VERTRES) - lTop - C_Dbl(At(vMargins, 3)) * lDpiY
    '--- output images
    If StartDoc(hDC, uInfo) <= 0 Then
        sError = GetSystemMessage(Err.LastDllError)
        GoTo QH
    End If
    uHeader.biSize = LenB(uHeader)
    For lIdx = 0 To UBound(vInputFiles)
        Call StartPage(hDC)
        If Not GetImageDimensions(CStr(vInputFiles(lIdx)), uHeader.biWidth, uHeader.biHeight, sError) Then
            GoTo QH
        End If
        baImage = ReadBinaryFile(CStr(vInputFiles(lIdx)))
        uHeader.biSizeImage = UBound(baImage) + 1
        uHeader.biCompression = IIf(baImage(0) = &H89, BI_PNG, BI_JPEG)
        If CDbl(lHeight) * uHeader.biWidth > CDbl(lWidth) * uHeader.biHeight Then
            lTargetW = lWidth
            lTargetH = Int(CDbl(lWidth) * uHeader.biHeight / uHeader.biWidth + 0.5)
            lTargetX = 0
            lTargetY = Int(CDbl(lHeight - lTargetH) / 2 + 0.5)
        Else
            lTargetW = Int(CDbl(lHeight) * uHeader.biWidth / uHeader.biHeight + 0.5)
            lTargetH = lHeight
            lTargetX = Int(CDbl(lWidth - lTargetW) / 2 + 0.5)
            lTargetY = 0
        End If
        Call StretchDIBits(hDC, _
            lLeft + lTargetX, lTop + lTargetY, lTargetW, lTargetH, _
            0, 0, uHeader.biWidth, uHeader.biHeight, _
            baImage(0), uHeader, 0, vbSrcCopy)
        Call EndPage(hDC)
    Next
    Call EndDoc(hDC)
    '--- success
    PrintImages = True
QH:
    On Error Resume Next
    If hDC <> 0 Then
        Call DeleteDC(hDC)
        hDC = 0
    End If
    Exit Function
EH:
    sError = "[&H" & Hex(Err.Number) & "] Critical: " & Err.Description & " [PrintImages]"
    Resume QH
End Function

Private Function SetupDevMode( _
            sPrinterName As String, _
            ByVal lPaperSize As Long, _
            ByVal lOrientation As Long, _
            baDevMode() As Byte, _
            sError As String) As Boolean
    Dim hPrinter        As Long
    Dim lNeeded         As Long
    Dim uDevMode        As DEVMODE
    
    On Error GoTo EH
    If OpenPrinter(sPrinterName, hPrinter, 0) = 0 Then
        sError = GetSystemMessage(Err.LastDllError) & " [OpenPrinter]"
        GoTo QH
    End If
    lNeeded = DocumentProperties(0, hPrinter, sPrinterName, ByVal 0&, ByVal 0&, 0)
    If lNeeded <= 0 Then
        sError = GetSystemMessage(Err.LastDllError) & " [DocumentProperties]"
        GoTo QH
    End If
    '--- round up to next 2KB page
    ReDim baDevMode(0 To (lNeeded And -2048) + 2047) As Byte
    If DocumentProperties(0, hPrinter, sPrinterName, baDevMode(0), ByVal 0&, DM_OUT_BUFFER) <> IDOK Then
        sError = GetSystemMessage(Err.LastDllError) & " [DocumentProperties#2]"
        GoTo QH
    End If
    Call CopyMemory(uDevMode, baDevMode(0), Len(uDevMode))
    If lPaperSize <> 0 Then
        uDevMode.dmPaperSize = lPaperSize
        uDevMode.dmFields = uDevMode.dmFields Or DM_PAPERSIZE
    End If
    If lOrientation <> 0 Then
        uDevMode.dmOrientation = lOrientation
        uDevMode.dmFields = uDevMode.dmFields Or DM_ORIENTATION
    End If
    Call CopyMemory(baDevMode(0), uDevMode, Len(uDevMode))
    Call DocumentProperties(0, hPrinter, sPrinterName, baDevMode(0), baDevMode(0), DM_IN_BUFFER Or DM_OUT_BUFFER)
    '--- success
    SetupDevMode = True
QH:
    On Error Resume Next
    If hPrinter <> 0 Then
        Call ClosePrinter(hPrinter)
        hPrinter = 0
    End If
    Exit Function
EH:
    sError = "[&H" & Hex(Err.Number) & "] Critical: " & Err.Description & " [SetupDevMode]"
    Resume QH
End Function

Private Function GetImageDimensions(sFile As String, lWidth As Long, lHeight As Long, sError As String) As Boolean
    Dim hBitmap         As Long
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    
    On Error GoTo EH
    If GdipLoadImageFromFile(StrPtr(sFile), hBitmap) <> 0 Then
        If Err.LastDllError = 0 Then
            sError = "Invalid image: " & Mid$(sFile, InStrRev(sFile, "\") + 1) & " [GdipLoadImageFromFile]"
        Else
            sError = GetSystemMessage(Err.LastDllError) & " [GdipLoadImageFromFile]"
        End If
        GoTo QH
    End If
    If GdipGetImageDimension(hBitmap, sngWidth, sngHeight) <> 0 Then
        sError = GetSystemMessage(Err.LastDllError) & " [GdipGetImageDimension]"
        GoTo QH
    End If
    lWidth = sngWidth
    lHeight = sngHeight
    '--- success
    GetImageDimensions = True
QH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
    Exit Function
EH:
    sError = "[&H" & Hex(Err.Number) & "] Critical: " & Err.Description & " [GetImageDimensions]"
    Resume QH
End Function

Private Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    On Error GoTo EH
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    Close nFile
End Function

Private Function CanonicalPath(sPath As String) As String
    Dim oFSO            As FileSystemObject
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    With oFSO
        CanonicalPath = .GetAbsolutePathName(sPath)
    End With
End Function

Private Function StartGdip() As Boolean
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    '--- success
    StartGdip = True
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = "[" & lLastDllError & "] " & Left$(GetSystemMessage, lSize)
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    If IsArray(vArray) Then
        If lIdx >= LBound(vArray) And lIdx <= UBound(vArray) Then
            At = vArray(lIdx)
        End If
    End If
QH:
End Function

Private Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function
