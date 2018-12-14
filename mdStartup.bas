Attribute VB_Name = "mdStartup"
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

Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&
'--- for DeviceCapabilities
Private Const DC_PAPERS                     As Long = 2
Private Const DC_PAPERSIZE                  As Long = 3
Private Const DC_PAPERNAMES                 As Long = 16
Private Const PAPERNAME_SIZE                As Long = 64

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal lExitCode As Long)
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, ByVal dev As Long) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_PDF_PRINTER   As String = "Microsoft Print to PDF"

Private m_oOpt                  As Object

Private Type UcsPaperInfoType
    PaperSize       As Long
    Name            As String
    Width           As Single
    Height          As Single
End Type

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Dim lExitCode       As Long
    
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde Then
        Call ExitProcess(lExitCode)
    End If
End Sub

Private Function Process(vArgs As Variant) As Long
    Dim sPrinterName    As String
    Dim cFiles          As Collection
    Dim vInputFiles     As Variant
    Dim lIdx            As Long
    Dim sError          As String
    Dim lPos            As Long
    Dim sFolder         As String
    Dim lPaperSize      As Long
    Dim lOrientation    As Long
    Dim vMargins        As Variant
    Dim uPapers()       As UcsPaperInfoType
    Dim sText           As String
    
    Set m_oOpt = GetOpt(vArgs, "printer:orientation:paper:margins:o")
    If Not m_oOpt.Item("-nologo") And Not m_oOpt.Item("-q") Then
        ConsoleError App.ProductName & " " & App.Major & "." & App.Minor & " (c) 2018 by wqweto@gmail.com" & vbCrLf
        ConsoleError "Convert jpeg/png images to multi-page pdf file" & vbCrLf & vbCrLf
    End If
    If LenB(m_oOpt.Item("error")) <> 0 Then
        ConsoleError "Error in command line: " & m_oOpt.Item("error") & vbCrLf & vbCrLf
        If Not (m_oOpt.Item("-h") Or m_oOpt.Item("-?") Or m_oOpt.Item("arg0") = "?") Then
            Exit Function
        End If
    End If
    If m_oOpt.Item("#arg") < 0 Or m_oOpt.Item("-h") Or m_oOpt.Item("-?") Or m_oOpt.Item("arg0") = "?" Then
        ConsoleError "Usage: %1.exe [options] <in_file.jpg> ..." & vbCrLf & vbCrLf, App.EXEName
        ConsoleError "Options:" & vbCrLf & _
            "  -o OUTFILE         write result to OUTFILE" & vbCrLf & _
            "  -paper SIZE        output paper size (e.g. A4)" & vbCrLf & _
            "  -orientation ORNT  page orientation (e.g. portrait)" & vbCrLf & _
            "  -margins L[/T/R/B] page margins in inches (e.g. 0.25)" & vbCrLf & _
            "  -q                 in quiet operation outputs only errors" & vbCrLf & _
            "  -nologo            suppress startup banner" & vbCrLf
        If m_oOpt.Item("#arg") < 0 Then
            Process = 100
        End If
        Exit Function
    End If
    Set cFiles = New Collection
    For lIdx = 0 To m_oOpt.Item("#arg")
        If FileExists(m_oOpt.Item("arg" & lIdx)) Then
            cFiles.Add m_oOpt.Item("arg" & lIdx)
        Else
            lPos = InStrRev(m_oOpt.Item("arg" & lIdx), "\")
            If lPos > 0 Then
                sFolder = Left$(m_oOpt.Item("arg" & lIdx), lPos - 1)
            End If
            If DirectoryExists(sFolder) And lPos > 0 Then
                EnumFiles sFolder, Mid$(m_oOpt.Item("arg" & lIdx), lPos), RetVal:=cFiles
            Else
                If Not m_oOpt.Item("-q") Then
                    ConsoleError "Warning: '%1' not found" & vbCrLf, m_oOpt.Item("arg" & lIdx)
                End If
            End If
        End If
    Next
    ReDim vInputFiles(0 To cFiles.Count - 1) As String
    For lIdx = 1 To cFiles.Count
        vInputFiles(lIdx - 1) = cFiles.Item(lIdx)
    Next
    sPrinterName = m_oOpt.Item("-printer")
    If LenB(sPrinterName) = 0 Then
       sPrinterName = STR_PDF_PRINTER
    End If
    Select Case LCase$(m_oOpt.Item("-orientation"))
    Case "p", "portrait"
        lOrientation = 1
    Case "l", "landscape"
        lOrientation = 2
    End Select
    If Not IsEmpty(m_oOpt.Item("-paper")) Then
        lPaperSize = C_Dbl(m_oOpt.Item("-paper"))
        If lPaperSize = 0 Then
            uPapers = EnumPrinterPapers(sPrinterName)
            For lIdx = 0 To UBound(uPapers)
                sText = sText & ", '" & uPapers(lIdx).Name & "'"
                If LCase$(uPapers(lIdx).Name) = LCase$(m_oOpt.Item("-paper")) Then
                    lPaperSize = uPapers(lIdx).PaperSize
                    Exit For
                End If
            Next
        End If
        If lPaperSize = 0 Then
            If Not m_oOpt.Item("-q") Then
                If LenB(sText) <> 0 Then
                    sText = ". Not from " & Mid$(sText, 3)
                End If
                ConsoleError "Warning: '%1' paper ignored" & sText & vbCrLf, m_oOpt.Item("-paper")
            End If
        End If
    End If
    If Not IsEmpty(m_oOpt.Item("-margins")) Then
        vMargins = Split(m_oOpt.Item("-margins"), "/")
        If UBound(vMargins) = 0 And C_Dbl(At(vMargins, 0)) > 0 Then
            vMargins = C_Dbl(At(vMargins, 0))
            vMargins = Array(vMargins, vMargins, vMargins, vMargins)
        End If
    End If
    If Not PrintImages( _
            sPrinterName, _
            vInputFiles, _
            m_oOpt.Item("-o"), _
            lPaperSize:=lPaperSize, _
            lOrientation:=lOrientation, _
            vMargins:=vMargins, _
            sError:=sError) Then
        ConsoleError sError & vbCrLf & vbCrLf
        Process = 2
        GoTo QH
    End If
    For lIdx = 1 To 30
        If FileExists(m_oOpt.Item("-o")) Then
            Exit For
        End If
        Call Sleep(100)
    Next
    If FileExists(m_oOpt.Item("-o")) Then
        If Not m_oOpt.Item("-q") Then
            ConsoleError m_oOpt.Item("-o") & " output sucesfully!" & vbCrLf & vbCrLf
        End If
    End If
QH:
End Function

Private Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Private Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    ConsolePrint = pvConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), sText, CVar(A))
End Function

Private Function ConsoleError(ByVal sText As String, ParamArray A() As Variant) As String
    ConsoleError = pvConsoleOutput(GetStdHandle(STD_ERROR_HANDLE), sText, CVar(A))
End Function

Private Function pvConsoleOutput(ByVal hOut As Long, ByVal sText As String, A As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long

    If LenB(sText) = 0 Then
        Exit Function
    End If
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    pvConsoleOutput = Replace(sText, ChrW$(LNG_PRIVATE), "%")
    '--- output
    If hOut = 0 Then
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

Private Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Dictionary
    Dim oRetVal         As Dictionary
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant
    Dim sValue          As String

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        .Item("#arg") = -1&
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            sValue = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            sValue = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            sValue = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option `" & vElem & "` requires an argument"
                        End If
                        vElem = "-" & vElem
                        If Not .Exists(vElem) Then
                            .Item(vElem) = sValue
                        Else
                            .Item("#" & vElem) = .Item("#" & vElem) + 1
                            .Item(vElem & .Item("#" & vElem)) = sValue
                        End If
                        GoTo Continue
                    End If
                Next
                vElem = "-" & Mid$(At(vArgs, lIdx), 2)
                .Item(vElem) = True
            Case Else
                vElem = "arg"
                sValue = At(vArgs, lIdx)
                .Item("#" & vElem) = .Item("#" & vElem) + 1
                .Item(vElem & .Item("#" & vElem)) = sValue
            End Select
Continue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

Private Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
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

Private Function FileExists(sFile As String) As Boolean
    FileExists = GetFileAttributes(sFile) <> -1
End Function

Private Function DirectoryExists(sFile As String) As Boolean
    DirectoryExists = (GetFileAttributes(sFile) And vbDirectory + vbVolume) = vbDirectory
End Function

Private Function EnumFiles( _
            sFolder As String, _
            Optional sMask As String, _
            Optional ByVal eAttrib As VbFileAttribute, _
            Optional RetVal As Collection) As Collection
    Dim sFile           As String
    
    If RetVal Is Nothing Then
        Set RetVal = New Collection
    End If
    sFile = Dir(PathCombine(sFolder, sMask))
    Do While LenB(sFile) <> 0
        If sFile <> "." And sFile <> ".." Then
            sFile = PathCombine(sFolder, sFile)
            If (GetAttr(sFile) And eAttrib) = eAttrib Then
                RetVal.Add sFile
            End If
        End If
        sFile = Dir
    Loop
    Set EnumFiles = RetVal
End Function

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Function EnumPrinterPapers(sPrinterName As String) As UcsPaperInfoType()
    Dim lNum            As Long
    Dim lIdx            As Long
    Dim naPapers()      As Integer
    Dim sPaperNames     As String
    Dim laPaperSizes()  As Long
    Dim uRetVal()       As UcsPaperInfoType
    
    lNum = DeviceCapabilities(sPrinterName, vbNullString, DC_PAPERS, ByVal vbNullString, 0)
    If lNum <= 0 Then
        ReDim uRetVal(-1 To -1) As UcsPaperInfoType
        GoTo QH
    End If
    ReDim naPapers(0 To lNum - 1) As Integer
    Call DeviceCapabilities(sPrinterName, vbNullString, DC_PAPERS, naPapers(0), 0)
    sPaperNames = String$(PAPERNAME_SIZE * lNum, 0)
    Call DeviceCapabilities(sPrinterName, vbNullString, DC_PAPERNAMES, ByVal sPaperNames, 0)
    ReDim laPaperSizes(0 To 2 * lNum - 1) As Long
    Call DeviceCapabilities(sPrinterName, vbNullString, DC_PAPERSIZE, laPaperSizes(0), 0)
    ReDim uRetVal(0 To lNum - 1) As UcsPaperInfoType
    For lIdx = 0 To lNum - 1
        With uRetVal(lIdx)
            .PaperSize = naPapers(lIdx)
            .Name = Mid$(sPaperNames, PAPERNAME_SIZE * lIdx + 1, PAPERNAME_SIZE)
            .Name = Left$(.Name, InStr(1, .Name, Chr$(0)) - 1)
            .Width = laPaperSizes(2 * lIdx) / 10#
            .Height = laPaperSizes(2 * lIdx + 1) / 10#
        End With
    Next
QH:
    EnumPrinterPapers = uRetVal
End Function

Private Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

