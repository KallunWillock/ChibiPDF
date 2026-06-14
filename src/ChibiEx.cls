VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChibiEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||   ChibiEx (v1.0)   ||||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'  AUTHOR:   Kallun Willock
'            (Based on WinRT OCR + PDF work by Frank Schüler)
'
'  URL:      https://github.com/KallunWillock/ChibiPDF
'            https://www.activevb.de
'
'  NOTES:    * Unified, dependency-free PDF extraction engine for VBA.
'            * Combines WinRT PDF rendering and OCR into a single 'zero-copy' pipeline.
'            * Extracts text from any PDF (including scanned/image-based documents).
'            * Renders PDF pages to PNG images (files, memory, or StdPicture).
'            * Supports direct OCR from image files, HBITMAPs, StdPictures, and streams.
'            * Lazy-initialises the OCR engine.
'            * Relies on the WinRT API set - requires minimum of Win10 (10.0.10240.0 or later).
'
'  USAGE:
'
'      Dim doc As New ChibiEx
'
'      ' 1. Load a PDF document
'      If doc.LoadFile("C:\docs\report.pdf") Then
'
'          Debug.Print "Pages: " & doc.PageCount
'
'          ' 2. Extract text from a single page
'          Dim pageText As String
'          pageText = doc.ExtractText(1)
'          Debug.Print "Page 1 Text: " & pageText
'
'          ' 3. Extract text from all pages
'          Dim fullText As String
'          fullText = doc.ExtractAllText()
'          Debug.Print "Full Document Text: " & fullText
'
'          ' 4. Render a specific page to a StdPicture
'          Dim pic As StdPicture
'          Set pic = doc.ToPicture(2)
'
'          ' 5. Render pages to disk or memory
'          ' doc.ToFile "C:\out\", 1, 3           ' Range of pages
'          ' doc.ToFile "C:\out\", Array(1, 5, 9)  ' Specific pages
'
'      Else
'          Debug.Print "Load failed: " & doc.LastError
'      End If
'
'      ' --- Standalone OCR Usage ---
'      ' You can also use ChibiEx purely for OCR on standard images:
'
'      ' doc.SetLanguage 0  ' Select first available OCR language
'      ' If doc.RecogniseFile("C:\docs\scan.png") Then
'      '     Debug.Print doc.ResultText
'      ' End If
'
'      ' If doc.RecogniseStdPicture(LoadPicture("C:\image.jpg")) Then
'      '     Debug.Print doc.ResultText
'      ' End If
'
'  LICENSE:  MIT
'
'  VERSION:  1.0    09/06/2026

Option Explicit

#If Win64 Then
  Private Const PTR_SIZE As Long = 8
  Private Const vbLongPtr As Long = vbLongLong
#Else
  Private Const PTR_SIZE As Long = 4
  Private Const vbLongPtr As Long = vbLong
#End If

Private Const S_OK                      As Long = 0
Private Const CC_STDCALL                As Long = 4
Private Const BSOS_DEFAULT              As Long = 0
Private Const FILE_ACCESS_READ          As Long = 0
Private Const ASYNC_STATUS_STARTED      As Long = 0
Private Const ASYNC_STATUS_COMPLETED    As Long = 1
Private Const ASYNC_MAX_SPIN            As Long = 1500000
Private Const ISTREAM_SLOT_READ         As Long = 3
Private Const ISTREAM_SLOT_SEEK         As Long = 5
Private Const STREAM_SEEK_SET           As Long = 0
Private Const STREAM_SEEK_END           As Long = 2
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const INVALID_FILE_ATTRIBUTES   As Long = -1
Private Const GENERIC_WRITE             As Long = &H40000000
Private Const CREATE_ALWAYS             As Long = 2
Private Const CLS_PdfDocument           As String = "Windows.Data.Pdf.PdfDocument"
Private Const CLS_OcrEngine             As String = "Windows.Media.Ocr.OcrEngine"
Private Const CLS_BitmapDecoder         As String = "Windows.Graphics.Imaging.BitmapDecoder"
Private Const CLSID_GDIPPngEncoder      As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Private Const IID_IPdfDocumentStatics   As String = "{433A0B5F-C007-4788-90F2-08143D922599}"
Private Const IID_IRandomAccessStream   As String = "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"
Private Const IID_IAsyncInfo            As String = "{00000036-0000-0000-C000-000000000046}"
Private Const IID_IClosable             As String = "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"
Private Const IID_IBitmapFrame          As String = "{72A49A1C-8081-438D-91BC-94ECFC8185C6}"
Private Const IID_IOcrEngineStatics     As String = "{5BFFA85A-3384-3540-9940-699120D428A8}"
Private Const IID_IBitmapDecoderStatics As String = "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}"
Private Const IID_IBitmapFrameWithSoftwareBitmap As String = "{FE287C9A-420C-4963-87AD-691436E08383}"
Private Const CHIBEX_VERSION            As Single = 1#

Private Enum VTbl
  ' IUnknown
  IUnknown_QueryInterface = 0
  IUnknown_AddRef = 1
  IUnknown_Release = 2
  ' IAsyncInfo / IAsyncOperation
  IAsyncInfo_GetStatus = 7
  IAsyncInfo_Close = 10
  IAsyncOperation_GetResults = 8
  ' IClosable
  IClosable_Close = 6
  ' PDF
  IPdfDocStatics_LoadFromStreamAsync = 8
  IPdfDocument_GetPage = 6
  IPdfDocument_GetPageCount = 7
  IPdfPage_RenderToStreamAsync = 6
  ' IReadOnlyList
  IReadOnlyList_GetAt = 6
  IReadOnlyList_GetSize = 7
  ' ILanguage
  ILanguage_GetLanguageTag = 6
  ILanguage_GetDisplayName = 7
  ILanguage_GetNativeName = 8
  ILanguage_GetScript = 9
  ' IOcrEngineStatics
  IOcrEngineStatics_GetMaxImageDimension = 6
  IOcrEngineStatics_GetAvailableRecognizerLanguages = 7
  IOcrEngineStatics_IsLanguageSupported = 8
  IOcrEngineStatics_TryCreateFromLanguage = 9
  ' IOcrEngine
  IOcrEngine_RecognizeAsync = 6
  IOcrEngine_GetRecognizerLanguage = 7
  ' IOcrResult
  IOcrResult_GetLines = 6
  IOcrResult_GetTextAngle = 7
  IOcrResult_GetText = 8
  ' IOcrLine
  IOcrLine_GetWords = 6
  IOcrLine_GetText = 7
  ' IOcrWord
  IOcrWord_GetBoundingRect = 6
  IOcrWord_GetText = 7
  ' IBitmapDecoderStatics
  IBitmapDecoderStatics_CreateAsync = 14
  ' IBitmapFrame
  IBitmapFrame_GetPixelWidth = 12
  IBitmapFrame_GetPixelHeight = 13
  ' IBitmapFrameWithSoftwareBitmap
  IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapAsync = 6
  ' IReferenceDouble
  IReferenceDouble_GetValue = 6
End Enum

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Type RECTF
  X      As Single
  Y      As Single
  Width  As Single
  Height As Single
End Type

Private Type OcrLanguage
  LanguageTag  As String
  DisplayName  As String
  NativeName   As String
  Script       As String
End Type

Private Type OcrWord
  Text  As String
  Rect  As RECTF
End Type

Private Type ocrLine
  Text       As String
  WordCount  As Long
  Words()    As OcrWord
End Type

Private Type OcrResults
  Language   As OcrLanguage
  Angle      As Double
  Text       As String
  LineCount  As Long
  Lines()    As ocrLine
End Type

#If VBA7 Then
  Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal pString As LongPtr, ByRef pClsid As GUID) As Long
  Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, ByVal prgvt As Any, ByVal prgpvarg As Any, ByRef pvargResult As Variant) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pTo As Any, ByRef uFrom As Any, ByVal lSize As LongPtr)
  Private Declare PtrSafe Function WindowsCreateString Lib "combase.dll" (ByVal sourceString As LongPtr, ByVal Length As Long, ByRef hString As LongPtr) As Long
  Private Declare PtrSafe Function WindowsDeleteString Lib "combase.dll" (ByVal hString As LongPtr) As Long
  Private Declare PtrSafe Function WindowsGetStringRawBuffer Lib "combase.dll" (ByVal hString As LongPtr, ByRef Length As Long) As LongPtr
  Private Declare PtrSafe Function RoGetActivationFactory Lib "combase.dll" (ByVal activatableClassId As LongPtr, ByRef rIID As GUID, ByRef factory As LongPtr) As Long
  Private Declare PtrSafe Function SHCreateMemStream Lib "shlwapi.dll" (ByVal pInit As LongPtr, ByVal cbInit As Long) As LongPtr
  Private Declare PtrSafe Function CreateRandomAccessStreamOnFile Lib "shcore.dll" (ByVal FilePath As LongPtr, ByVal accessMode As Long, ByRef rIID As GUID, ByRef ppV As LongPtr) As Long
  Private Declare PtrSafe Function CreateRandomAccessStreamOverStream Lib "shcore.dll" (ByVal pStream As LongPtr, ByVal options As Long, ByRef rIID As GUID, ByRef ppV As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, ByRef bitmap As LongPtr) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As LongPtr) As Long
  Private Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus.dll" (ByVal mImage As LongPtr, ByVal mStream As LongPtr, ByRef mClsidEncoder As GUID, ByVal mEncoderParams As LongPtr) As Long
  Private Declare PtrSafe Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare PtrSafe Function WriteFile Lib "kernel32.dll" (ByVal hFile As LongPtr, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As LongPtr) As Long
#Else
  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, ByRef pClsid As GUID) As Long
  Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, ByVal prgvt As Any, ByVal prgpvarg As Any, ByRef pvargResult As Variant) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pTo As Any, ByRef uFrom As Any, ByVal lSize As Long)
  Private Declare Function WindowsCreateString Lib "combase.dll" (ByVal sourceString As Long, ByVal Length As Long, ByRef hString As Long) As Long
  Private Declare Function WindowsDeleteString Lib "combase.dll" (ByVal hString As Long) As Long
  Private Declare Function WindowsGetStringRawBuffer Lib "combase.dll" (ByVal hString As Long, ByRef Length As Long) As Long
  Private Declare Function RoGetActivationFactory Lib "combase.dll" (ByVal activatableClassId As Long, ByRef riid As GUID, ByRef factory As Long) As Long
  Private Declare Function SHCreateMemStream Lib "shlwapi.dll" (ByVal pInit As Long, ByVal cbInit As Long) As Long
  Private Declare Function CreateRandomAccessStreamOnFile Lib "shcore.dll" (ByVal FilePath As Long, ByVal accessMode As Long, ByRef riid As GUID, ByRef ppv As Long) As Long
  Private Declare Function CreateRandomAccessStreamOverStream Lib "shcore.dll" (ByVal pStream As Long, ByVal options As Long, ByRef riid As GUID, ByRef ppv As Long) As Long
  Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As Long, ByVal hPal As Long, ByRef bitmap As Long) As Long
  Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
  Private Declare Function GdipSaveImageToStream Lib "gdiplus.dll" (ByVal mImage As Long, ByVal mStream As Long, ByRef mClsidEncoder As GUID, ByVal mEncoderParams As Long) As Long
  Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
#End If

Private m_Factory                 As LongPtr
Private m_IsInitialized           As Boolean
Private m_PageCount               As Long
Private m_pDoc                    As LongPtr
Private m_FilePath                As String

Private m_OcrReady                As Boolean
Private m_pIOcrEngineStatics      As LongPtr
Private m_pIBitmapDecoderStatics  As LongPtr
Private m_pIOcrEngine             As LongPtr
Private m_MaxImageDimension       As Long
Private m_LanguageCount           As Long
Private m_Languages()             As OcrLanguage
Private m_Results                 As OcrResults
Private m_LastError               As String
Private m_DebugMode               As Boolean

Private Sub Class_Initialize()
  m_PageCount = 0
  m_LastError = vbNullString
  m_DebugMode = False
  m_IsInitialized = False
  m_Factory = 0
  m_pDoc = 0
  m_FilePath = vbNullString

  m_Factory = ActivateFactory(CLS_PdfDocument, IID_IPdfDocumentStatics)
  m_IsInitialized = (m_Factory <> 0)
  DebugPrint "PDF Factory initialisation: " & IIf(m_IsInitialized, "OK", "FAILED")
End Sub

Private Sub Class_Terminate()
  CloseFile
  ReleaseInterface m_Factory
  
  If m_OcrReady Then
    ReleaseInterface m_pIOcrEngine
    ReleaseInterface m_pIBitmapDecoderStatics
    ReleaseInterface m_pIOcrEngineStatics
  End If
End Sub

Public Property Get PageCount() As Long: PageCount = m_PageCount: End Property
Public Property Get LastError() As String: LastError = m_LastError: End Property
Public Property Get FilePath() As String: FilePath = m_FilePath: End Property

Public Property Get DebugMode() As Boolean: DebugMode = m_DebugMode: End Property
Public Property Let DebugMode(ByVal value As Boolean): m_DebugMode = value: End Property

Public Property Get MaxImageDimension() As Long
  EnsureOcrReady
  MaxImageDimension = m_MaxImageDimension
End Property

Public Property Get OCRLanguageCount() As Long
  EnsureOcrReady
  OCRLanguageCount = m_LanguageCount
End Property

Public Property Get OCRLanguageDisplayName(ByVal index As Long) As String
  EnsureOcrReady
  If index >= 0 And index < m_LanguageCount Then OCRLanguageDisplayName = m_Languages(index).DisplayName
End Property

Public Property Get ResultText() As String
  ResultText = m_Results.Text
End Property

Public Property Get ResultTextWithLineBreaks() As String
  Dim result  As String
  Dim i       As Long
  For i = 0 To m_Results.LineCount - 1
    result = result & m_Results.Lines(i).Text
    If i < m_Results.LineCount - 1 Then result = result & vbNewLine
  Next i
  ResultTextWithLineBreaks = result
End Property

Public Property Get ResultAngle() As Double
  ResultAngle = m_Results.Angle
End Property

Public Property Get ResultLine(ByVal lineIndex As Long) As String
  If m_Results.LineCount <= 0 Then Exit Property
  If (Not Not m_Results.Lines) = 0 Then Exit Property

  If lineIndex >= 0 And lineIndex < m_Results.LineCount Then
    ResultLine = m_Results.Lines(lineIndex).Text
  End If
End Property

Public Property Get ResultLineCount() As Long
  ResultLineCount = m_Results.LineCount
End Property

Public Property Get ResultLineWordCount(ByVal lineIndex As Long) As Long
  If lineIndex >= 0 And lineIndex < m_Results.LineCount Then
    ResultLineWordCount = m_Results.Lines(lineIndex).WordCount
  End If
End Property

Public Property Get ResultLineWord(ByVal lineIndex As Long, ByVal wordIndex As Long) As String
  If m_Results.LineCount <= 0 Then Exit Property
  If (Not Not m_Results.Lines) = 0 Then Exit Property
  If lineIndex < 0 Or lineIndex >= m_Results.LineCount Then Exit Property

  With m_Results.Lines(lineIndex)
    If .WordCount <= 0 Then Exit Property
    If (Not Not .Words) = 0 Then Exit Property
    If wordIndex >= 0 And wordIndex < .WordCount Then
      ResultLineWord = .Words(wordIndex).Text
    End If
  End With
End Property

Public Property Get Version() As Single
  Version = CHIBEX_VERSION
End Property

Public Function LoadFile(ByVal pdfPath As String) As Boolean
  CloseFile
  If Not FileExists(pdfPath) Then
    m_LastError = "PDF file not found: " & pdfPath
    Exit Function
  End If
  If Not m_IsInitialized Then
    m_LastError = "WinRT PDF factory is not initialised."
    Exit Function
  End If

  m_pDoc = LoadDocument(pdfPath)
  If m_pDoc = 0 Then Exit Function

  InvokeInterface m_pDoc, IPdfDocument_GetPageCount, vbLong, VarPtr(m_PageCount)
  If m_PageCount <= 0 Then
    m_LastError = "PDF contains no pages."
    ReleaseInterface m_pDoc
    Exit Function
  End If

  m_FilePath = pdfPath
  DebugPrint "Loaded '" & pdfPath & "' - " & m_PageCount & " page(s)"
  LoadFile = True
End Function

Public Sub CloseFile()
  ReleaseInterface m_pDoc
  m_pDoc = 0
  m_PageCount = 0
  m_FilePath = vbNullString
  m_LastError = vbNullString
End Sub

' =======================================
'  ToFile: render selected pages to disk
' =======================================

Public Function ToFile(ByVal outputFolder As String, Optional ByVal pagesSelector As Variant, Optional ByVal lastPage As Variant) As Variant
  Dim pages() As Long, out() As String, count As Long, baseName As String, folder As String, idx As Variant
  ToFile = Empty
  m_LastError = vbNullString

  If m_pDoc = 0 Then m_LastError = "No document loaded.": Exit Function

  folder = IIf(Len(outputFolder) = 0, ParentFolder(m_FilePath), TrailingSlash(outputFolder))
  If Not FolderExists(folder) Then MkDir folder
  baseName = FileBasename(m_FilePath)
  pages = ResolvePages(pagesSelector, lastPage)

  If Not Not pages Then
    Dim PageNumberFormat As String
    PageNumberFormat = IIf(m_PageCount > 999, String(4, "0"), String(3, "0"))
    For Each idx In pages
      Dim pg As Long: pg = CLng(idx)
      Dim png() As Byte: png = RenderPageToMemory(m_pDoc, pg)
      If Not Not png Then
        Dim outPath As String: outPath = folder & baseName & "_p" & Format$(pg + 1, PageNumberFormat) & ".png"
        If WriteByteArrayToFile(outPath, png) Then
          ReDim Preserve out(0 To count): out(count) = outPath: count = count + 1
        End If
      End If
      DoEvents
    Next idx
  End If
  If count > 0 Then ToFile = out
End Function

' ==================================================
'  ToMemory: render selected pages to byte array(s)
' ==================================================

Public Function ToMemory(Optional ByVal pagesSelector As Variant, Optional ByVal lastPage As Variant) As Variant
  Dim pages() As Long, blobs() As Variant, count As Long, idx As Variant
  ToMemory = Empty
  m_LastError = vbNullString
  If m_pDoc = 0 Then m_LastError = "No document loaded.": Exit Function

  pages = ResolvePages(pagesSelector, lastPage)
  If Not Not pages Then
    For Each idx In pages
      Dim pg As Long: pg = CLng(idx)
      Dim png() As Byte: png = RenderPageToMemory(m_pDoc, pg)
      If Not Not png Then
        ReDim Preserve blobs(0 To count): blobs(count) = png: count = count + 1
      End If
      DoEvents
    Next idx
  End If
  If count > 0 Then ToMemory = blobs
End Function

' ================================================
'  ToPicture: render a single page to StdPicture.
' ================================================

Public Function ToPicture(ByVal pageNum As Long) As StdPicture
  m_LastError = vbNullString
  If m_pDoc = 0 Then m_LastError = "No document loaded.": Exit Function
  If pageNum < 1 Or pageNum > m_PageCount Then m_LastError = "Page out of range.": Exit Function

  Dim png() As Byte: png = RenderPageToMemory(m_pDoc, pageNum - 1)
  If Not Not png Then
    With CreateObject("WIA.Vector")
      .BinaryData = png
      Set ToPicture = .Picture
    End With
  End If
End Function

' =======================================================
'  ToStream: render a single page to IRandomAccessStream
'  NB: caller must CloseStream.
' =======================================================

#If Win64 Then
Public Function ToStream(ByVal pageNum As Long) As LongPtr
#Else
Public Function ToStream(ByVal pageNum As Long) As Long
#End If
  m_LastError = vbNullString
  If m_pDoc = 0 Then m_LastError = "No document loaded.": Exit Function
  If pageNum < 1 Or pageNum > m_PageCount Then m_LastError = "Page out of range.": Exit Function
  ToStream = RenderPageToStream(m_pDoc, pageNum - 1)
End Function

' ================================================
'  CloseStream: Public method to clean up streams
' ================================================

#If Win64 Then
Public Sub CloseStream(ByVal pStream As LongPtr)
#Else
Public Sub CloseStream(ByVal pStream As Long)
#End If
  If pStream = 0 Then Exit Sub
  Dim iid As GUID, pClosable As LongPtr
  GuidFromString IID_IClosable, iid
  InvokeInterface pStream, IUnknown_QueryInterface, vbLong, VarPtr(iid), VarPtr(pClosable)
  If pClosable <> 0 Then
    InvokeInterface pClosable, IClosable_Close, vbLong
    ReleaseInterface pClosable
  End If
  ReleaseInterface pStream
End Sub

' =================
'  Text Extraction
' =================

Public Function ExtractText(ByVal pageNum As Long) As String
  m_LastError = vbNullString
  ExtractText = vbNullString

  If m_PageCount = 0 Then m_LastError = "No document loaded.": Exit Function

  EnsureOcrReady
  If m_pIOcrEngine = 0 Then Exit Function

  Dim pStream As LongPtr
  pStream = ToStream(pageNum)
  If pStream = 0 Then Exit Function

  If RunRecognition(pStream) Then
    ExtractText = m_Results.Text
  Else
    m_LastError = "OCR failed: " & m_LastError
  End If

  CloseStream pStream
End Function

Public Function ExtractAllText(Optional ByVal pageSeparator As String = vbCrLf & "--- Page {n} ---" & vbCrLf) As String
  m_LastError = vbNullString
  If m_PageCount = 0 Then m_LastError = "No document loaded.": Exit Function

  Dim pages As Collection
  Set pages = New Collection

  Dim i As Long
  Dim pageText As String

  For i = 1 To m_PageCount
    pageText = ExtractText(i)
    If Len(pageText) > 0 Then
      pages.Add pageText
    End If
    DoEvents
  Next i

  If pages.count = 0 Then
    ExtractAllText = vbNullString
    Exit Function
  End If

  Dim arr() As String
  ReDim arr(0 To pages.count - 1)

  For i = 1 To pages.count
    arr(i - 1) = pages(i)
  Next i

  If InStr(pageSeparator, "{n}") > 0 Then
    Dim result As String
    result = arr(0)
    For i = 1 To pages.count - 1
      result = result & Replace(pageSeparator, "{n}", CStr(i + 1)) & arr(i)
    Next i
    ExtractAllText = result
  Else
    ExtractAllText = Join(arr, pageSeparator)
  End If
End Function

' ====================
'  OCR Public Methods
' ====================

' SetLanguage: Uses "Create then Swap" to prevent breaking the class on failure.
Public Function SetLanguage(ByVal index As Long) As Boolean
  EnsureOcrReady
  m_LastError = vbNullString
  If Not m_OcrReady Then m_LastError = "OCR engine not initialised.": Exit Function
  If index < 0 Or index >= m_LanguageCount Then m_LastError = "Language index out of range.": Exit Function

  Dim pLanguage As LongPtr
  If Not GetLanguageByIndex(index, pLanguage) Then m_LastError = "Could not retrieve language.": Exit Function

  Dim isSupported As Boolean
  InvokeInterface m_pIOcrEngineStatics, IOcrEngineStatics_IsLanguageSupported, vbLong, pLanguage, VarPtr(isSupported)

  If isSupported = False Then
    m_LastError = "Language not supported for OCR."
    ReleaseInterface pLanguage ' Must release before exiting!
    Exit Function
  End If

  ' Create new engine into temporary variable
  Dim pNewEngine As LongPtr
  InvokeInterface m_pIOcrEngineStatics, IOcrEngineStatics_TryCreateFromLanguage, vbLong, pLanguage, VarPtr(pNewEngine)
  ReleaseInterface pLanguage

  If pNewEngine = 0 Then
    m_LastError = "Failed to create OCR engine."
    Exit Function
  End If
  
  ReleaseInterface m_pIOcrEngine
  m_pIOcrEngine = pNewEngine
  SetLanguage = True
End Function

Public Function RecogniseFile(ByVal FilePath As String) As Boolean
  m_LastError = vbNullString
  If Not GuardReady() Then Exit Function

  Dim pStream As LongPtr, g As GUID
  GuidFromString IID_IRandomAccessStream, g
  If CreateRandomAccessStreamOnFile(StrPtr(FilePath), FILE_ACCESS_READ, g, pStream) <> S_OK _
    Or pStream = 0 Then
    m_LastError = "Could not open file: " & FilePath
    Exit Function
  End If

  GetUsedLanguage
  RecogniseFile = RunRecognition(pStream)
  CloseStream pStream
End Function

#If Win64 Then
Public Function RecogniseBitmap(ByVal hBitmap As LongPtr) As Boolean
#Else
Public Function RecogniseBitmap(ByVal hBitmap As Long) As Boolean
#End If
  m_LastError = vbNullString
  If Not GuardReady() Then Exit Function
  If hBitmap = 0 Then m_LastError = "HBITMAP handle is null.": Exit Function

  Dim pGdipBitmap As LongPtr
  If GdipCreateBitmapFromHBITMAP(hBitmap, 0, pGdipBitmap) <> 0 Or pGdipBitmap = 0 Then
    m_LastError = "GdipCreateBitmapFromHBITMAP failed.": Exit Function
  End If

  Dim pIStream As LongPtr: pIStream = SHCreateMemStream(0, 0)
  If pIStream = 0 Then m_LastError = "SHCreateMemStream failed.": GdipDisposeImage pGdipBitmap: Exit Function

  Dim encoderGuid As GUID: GuidFromString CLSID_GDIPPngEncoder, encoderGuid
  If GdipSaveImageToStream(pGdipBitmap, pIStream, encoderGuid, 0) <> 0 Then
    m_LastError = "GdipSaveImageToStream failed."
    ReleaseInterface pIStream
    GdipDisposeImage pGdipBitmap
    Exit Function
  End If
  GdipDisposeImage pGdipBitmap

  IStreamSeek pIStream, 0, STREAM_SEEK_SET

  Dim pRAS As LongPtr, g As GUID
  GuidFromString IID_IRandomAccessStream, g
  If CreateRandomAccessStreamOverStream(pIStream, BSOS_DEFAULT, g, pRAS) <> S_OK Or pRAS = 0 Then
    m_LastError = "CreateRandomAccessStreamOverStream failed."
    ReleaseInterface pIStream
    Exit Function
  End If

  GetUsedLanguage
  RecogniseBitmap = RunRecognition(pRAS)
  CloseStream pRAS
  ReleaseInterface pIStream
End Function

#If Win64 Then
Public Function RecogniseStream(ByVal pRandomAccessStream As LongPtr) As Boolean
#Else
Public Function RecogniseStream(ByVal pRandomAccessStream As Long) As Boolean
#End If
  m_LastError = vbNullString
  If Not GuardReady() Then Exit Function
  If pRandomAccessStream = 0 Then
    m_LastError = "Stream pointer is null."
    Exit Function
  End If
  GetUsedLanguage
  RecogniseStream = RunRecognition(pRandomAccessStream)
End Function

Public Function RecogniseStdPicture(ByVal pic As StdPicture) As Boolean
  m_LastError = vbNullString
  If pic Is Nothing Then
    m_LastError = "StdPicture is Nothing."
    Exit Function
  End If
  If pic.Type <> 1 Then ' vbPicTypeBitmap
    m_LastError = "Only bitmaps are supported."
    Exit Function
  End If
  EnsureOcrReady
  RecogniseStdPicture = RecogniseBitmap(pic.Handle)
End Function

' ====================================
'  Private - Lazy Init & OCR Pipeline
' ====================================

Private Sub EnsureOcrReady()
  If m_OcrReady Then Exit Sub

  m_pIOcrEngineStatics = ActivateFactory(CLS_OcrEngine, IID_IOcrEngineStatics)
  If m_pIOcrEngineStatics = 0 Then m_LastError = "Failed to activate OCR factory.": Exit Sub

  m_pIBitmapDecoderStatics = ActivateFactory(CLS_BitmapDecoder, IID_IBitmapDecoderStatics)
  If m_pIBitmapDecoderStatics = 0 Then m_LastError = "Failed to activate BitmapDecoder factory.": Exit Sub

  InvokeInterface m_pIOcrEngineStatics, IOcrEngineStatics_GetMaxImageDimension, vbLong, VarPtr(m_MaxImageDimension)
  LoadAvailableLanguages

  m_OcrReady = True

  If m_LanguageCount > 0 Then SetLanguage 0

  DebugPrint "OCR Engine lazy-initialised."
End Sub

Private Function GuardReady() As Boolean
  EnsureOcrReady
  If Not m_OcrReady Then m_LastError = "OCR engine not initialised.": Exit Function
  If m_pIOcrEngine = 0 Then m_LastError = "No recognition language set.": Exit Function
  GuardReady = True
End Function

Private Sub ResetResults()
  m_Results.Text = vbNullString
  m_Results.Angle = 0
  m_Results.LineCount = 0
  Erase m_Results.Lines
End Sub

Private Function RunRecognition(ByVal pStream As LongPtr) As Boolean
  Dim pSoftwareBitmap As LongPtr
  Dim pOcrResultAsync As LongPtr
  Dim pOcrResult      As LongPtr
  Dim pRefDouble      As LongPtr
  Dim hString         As LongPtr

  ResetResults

  If Not GetSoftwareBitmapFromStream(pStream, pSoftwareBitmap) Then Exit Function

  If InvokeInterface(m_pIOcrEngine, IOcrEngine_RecognizeAsync, vbLong, pSoftwareBitmap, VarPtr(pOcrResultAsync)) <> S_OK Or pOcrResultAsync = 0 Then
    m_LastError = "RecognizeAsync returned null."
    CloseStream pSoftwareBitmap
    Exit Function
  End If

  If Not AwaitAsync(pOcrResultAsync, pOcrResult) Then
    m_LastError = "RecognizeAsync failed or timed out."
    ReleaseInterface pOcrResultAsync
    CloseStream pSoftwareBitmap
    Exit Function
  End If
  ReleaseInterface pOcrResultAsync

  m_Results.Angle = 0
  If InvokeInterface(pOcrResult, IOcrResult_GetTextAngle, vbLong, VarPtr(pRefDouble)) = S_OK Then
    If pRefDouble <> 0 Then
      InvokeInterface pRefDouble, IReferenceDouble_GetValue, vbLong, VarPtr(m_Results.Angle)
      ReleaseInterface pRefDouble
    End If
  End If

  m_Results.Text = vbNullString
  If InvokeInterface(pOcrResult, IOcrResult_GetText, vbLong, VarPtr(hString)) = S_OK Then
    If hString <> 0 Then
      m_Results.Text = HStringToString(hString)
      WindowsDeleteString hString
    End If
  End If

  RunRecognition = ExtractLines(pOcrResult)
  ReleaseInterface pOcrResult
  CloseStream pSoftwareBitmap
End Function

Private Function GetSoftwareBitmapFromStream(ByVal pStream As LongPtr, ByRef pSoftwareBitmap As LongPtr) As Boolean
  Dim pDecoderAsync   As LongPtr
  Dim pDecoder        As LongPtr
  Dim pFrame          As LongPtr
  Dim pFrameSBM       As LongPtr
  Dim pSbmAsync       As LongPtr
  Dim frameW          As Long
  Dim frameH          As Long
  Dim iid             As GUID

  If InvokeInterface(m_pIBitmapDecoderStatics, IBitmapDecoderStatics_CreateAsync, vbLong, pStream, VarPtr(pDecoderAsync)) <> S_OK Then
    m_LastError = "BitmapDecoder CreateAsync failed."
    Exit Function
  End If
  If pDecoderAsync = 0 Then
    m_LastError = "BitmapDecoder CreateAsync returned null."
    Exit Function
  End If

  If Not AwaitAsync(pDecoderAsync, pDecoder) Then
    m_LastError = "BitmapDecoder async failed or timed out."
    ReleaseInterface pDecoderAsync
    Exit Function
  End If
  ReleaseInterface pDecoderAsync

  GuidFromString IID_IBitmapFrame, iid
  If InvokeInterface(pDecoder, IUnknown_QueryInterface, vbLong, VarPtr(iid), VarPtr(pFrame)) <> S_OK Or pFrame = 0 Then
    m_LastError = "QueryInterface for IBitmapFrame failed."
    ReleaseInterface pDecoder
    Exit Function
  End If

  InvokeInterface pFrame, IBitmapFrame_GetPixelWidth, vbLong, VarPtr(frameW)
  InvokeInterface pFrame, IBitmapFrame_GetPixelHeight, vbLong, VarPtr(frameH)
  ReleaseInterface pFrame

  If frameW > m_MaxImageDimension Or frameH > m_MaxImageDimension Then
    m_LastError = "Image (" & frameW & "x" & frameH & ") exceeds maximum OCR dimension."
    ReleaseInterface pDecoder
    Exit Function
  End If

  GuidFromString IID_IBitmapFrameWithSoftwareBitmap, iid
  If InvokeInterface(pDecoder, IUnknown_QueryInterface, vbLong, VarPtr(iid), VarPtr(pFrameSBM)) <> S_OK Or pFrameSBM = 0 Then
    m_LastError = "QueryInterface for IBitmapFrameWithSoftwareBitmap failed."
    ReleaseInterface pDecoder
    Exit Function
  End If

  If InvokeInterface(pFrameSBM, IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapAsync, vbLong, VarPtr(pSbmAsync)) <> S_OK Or pSbmAsync = 0 Then
    m_LastError = "GetSoftwareBitmapAsync returned null."
    ReleaseInterface pFrameSBM
    ReleaseInterface pDecoder
    Exit Function
  End If
  ReleaseInterface pFrameSBM
  ReleaseInterface pDecoder

  If Not AwaitAsync(pSbmAsync, pSoftwareBitmap) Then
    m_LastError = "GetSoftwareBitmapAsync failed or timed out."
    ReleaseInterface pSbmAsync
    Exit Function
  End If
  ReleaseInterface pSbmAsync

  GetSoftwareBitmapFromStream = True
End Function

Private Function ExtractLines(ByVal pOcrResult As LongPtr) As Boolean
  Dim pList As LongPtr, pLine As LongPtr, hString As LongPtr, i As Long
  InvokeInterface pOcrResult, IOcrResult_GetLines, vbLong, VarPtr(pList)
  If pList = 0 Then Exit Function

  m_Results.LineCount = 0
  InvokeInterface pList, IReadOnlyList_GetSize, vbLong, VarPtr(m_Results.LineCount)

  If m_Results.LineCount > 0 Then
    ReDim m_Results.Lines(0 To m_Results.LineCount - 1)
    For i = 0 To m_Results.LineCount - 1
      InvokeInterface pList, IReadOnlyList_GetAt, vbLong, CLng(i), VarPtr(pLine)
      If pLine <> 0 Then
        InvokeInterface pLine, IOcrLine_GetText, vbLong, VarPtr(hString)
        If hString <> 0 Then
          m_Results.Lines(i).Text = HStringToString(hString)
          WindowsDeleteString hString
        End If
        ExtractWords pLine, m_Results.Lines(i)
        ReleaseInterface pLine
      End If
    Next i
  End If
  ReleaseInterface pList
  ExtractLines = True
End Function

Private Sub ExtractWords(ByVal pLine As LongPtr, ByRef ocrLine As ocrLine)
  Dim pList As LongPtr, pWord As LongPtr, hString As LongPtr, tRect As RECTF, i As Long
  ocrLine.WordCount = 0
  InvokeInterface pLine, IOcrLine_GetWords, vbLong, VarPtr(pList)
  If pList = 0 Then Exit Sub

  InvokeInterface pList, IReadOnlyList_GetSize, vbLong, VarPtr(ocrLine.WordCount)
  If ocrLine.WordCount > 0 Then
    ReDim ocrLine.Words(0 To ocrLine.WordCount - 1)
    For i = 0 To ocrLine.WordCount - 1
      InvokeInterface pList, IReadOnlyList_GetAt, vbLong, CLng(i), VarPtr(pWord)
      If pWord <> 0 Then
        InvokeInterface pWord, IOcrWord_GetBoundingRect, vbLong, VarPtr(tRect)
        ocrLine.Words(i).Rect = tRect
        InvokeInterface pWord, IOcrWord_GetText, vbLong, VarPtr(hString)
        If hString <> 0 Then
          ocrLine.Words(i).Text = HStringToString(hString)
          WindowsDeleteString hString
        End If
        ReleaseInterface pWord
      End If
    Next i
  End If
  ReleaseInterface pList
End Sub

Private Sub GetUsedLanguage()
  Dim pLanguage As LongPtr, hString As LongPtr
  InvokeInterface m_pIOcrEngine, IOcrEngine_GetRecognizerLanguage, vbLong, VarPtr(pLanguage)
  If pLanguage = 0 Then Exit Sub

  InvokeInterface pLanguage, ILanguage_GetLanguageTag, vbLong, VarPtr(hString)
  If hString <> 0 Then
    m_Results.Language.LanguageTag = HStringToString(hString)
    WindowsDeleteString hString
  End If
  
  InvokeInterface pLanguage, ILanguage_GetDisplayName, vbLong, VarPtr(hString)
  If hString <> 0 Then
    m_Results.Language.DisplayName = HStringToString(hString)
    WindowsDeleteString hString
  End If
  
  InvokeInterface pLanguage, ILanguage_GetNativeName, vbLong, VarPtr(hString)
  If hString <> 0 Then
    m_Results.Language.NativeName = HStringToString(hString)
    WindowsDeleteString hString
  End If
  
  InvokeInterface pLanguage, ILanguage_GetScript, vbLong, VarPtr(hString)
  If hString <> 0 Then
    m_Results.Language.Script = HStringToString(hString)
    WindowsDeleteString hString
  End If
  
  ReleaseInterface pLanguage
End Sub

Private Sub LoadAvailableLanguages()
  Dim pList As LongPtr, pLanguage As LongPtr, hString As LongPtr, count As Long, i As Long
  InvokeInterface m_pIOcrEngineStatics, IOcrEngineStatics_GetAvailableRecognizerLanguages, vbLong, VarPtr(pList)
  If pList = 0 Then Exit Sub

  InvokeInterface pList, IReadOnlyList_GetSize, vbLong, VarPtr(count)
  If count > 0 Then
    ReDim m_Languages(0 To count - 1)
    For i = 0 To count - 1
      InvokeInterface pList, IReadOnlyList_GetAt, vbLong, CLng(i), VarPtr(pLanguage)
      If pLanguage <> 0 Then
        InvokeInterface pLanguage, ILanguage_GetLanguageTag, vbLong, VarPtr(hString)
        If hString <> 0 Then m_Languages(i).LanguageTag = HStringToString(hString): WindowsDeleteString hString
        InvokeInterface pLanguage, ILanguage_GetDisplayName, vbLong, VarPtr(hString)
        If hString <> 0 Then m_Languages(i).DisplayName = HStringToString(hString): WindowsDeleteString hString
        InvokeInterface pLanguage, ILanguage_GetNativeName, vbLong, VarPtr(hString)
        If hString <> 0 Then m_Languages(i).NativeName = HStringToString(hString): WindowsDeleteString hString
        InvokeInterface pLanguage, ILanguage_GetScript, vbLong, VarPtr(hString)
        If hString <> 0 Then m_Languages(i).Script = HStringToString(hString): WindowsDeleteString hString
        ReleaseInterface pLanguage
      End If
    Next i
    m_LanguageCount = count
  End If
  ReleaseInterface pList
End Sub

Private Function GetLanguageByIndex(ByVal index As Long, ByRef pLanguage As LongPtr) As Boolean
  Dim pList As LongPtr, count As Long
  InvokeInterface m_pIOcrEngineStatics, IOcrEngineStatics_GetAvailableRecognizerLanguages, vbLong, VarPtr(pList)
  If pList = 0 Then Exit Function
  InvokeInterface pList, IReadOnlyList_GetSize, vbLong, VarPtr(count)
  If index >= 0 And index < count Then
    InvokeInterface pList, IReadOnlyList_GetAt, vbLong, CLng(index), VarPtr(pLanguage)
    GetLanguageByIndex = (pLanguage <> 0)
  End If
  ReleaseInterface pList
End Function

Private Function RenderPageToMemory(ByVal pDoc As LongPtr, ByVal zeroIdx As Long) As Byte()
  Dim pPage As LongPtr, pIStream As LongPtr, pRAS As LongPtr, pAsync As LongPtr, g As GUID, cbSize As Long, png() As Byte
  RenderPageToMemory = png

  InvokeInterface pDoc, IPdfDocument_GetPage, vbLong, zeroIdx, VarPtr(pPage)
  If pPage = 0 Then m_LastError = "GetPage returned null.": Exit Function

  pIStream = SHCreateMemStream(0, 0)
  If pIStream = 0 Then m_LastError = "SHCreateMemStream failed.": GoTo Cleanup

  GuidFromString IID_IRandomAccessStream, g
  If CreateRandomAccessStreamOverStream(pIStream, BSOS_DEFAULT, g, pRAS) <> S_OK Or pRAS = 0 Then
    m_LastError = "CreateRandomAccessStreamOverStream failed.": GoTo Cleanup
  End If

  InvokeInterface pPage, IPdfPage_RenderToStreamAsync, vbLong, pRAS, VarPtr(pAsync)
  If pAsync = 0 Then m_LastError = "RenderToStreamAsync returned null.": GoTo Cleanup
  If Not AwaitAsync(pAsync, expectResult:=False) Then m_LastError = "Render async failed.": GoTo Cleanup

  Dim posEnd As Currency
  IStreamSeek pIStream, 0, STREAM_SEEK_END, posEnd
  cbSize = CLng(posEnd)
  If cbSize <= 0 Then m_LastError = "Rendered stream is empty.": GoTo Cleanup

  IStreamSeek pIStream, 0, STREAM_SEEK_SET
  ReDim png(0 To cbSize - 1)
  Dim bytesRead As Long
  If Not IStreamRead(pIStream, png(0), cbSize, bytesRead) Or bytesRead <> cbSize Then
    m_LastError = "IStream::Read failed.": Erase png
  End If

Cleanup:
  CloseStream pRAS
  ReleaseInterface pIStream
  ReleaseInterface pPage
  ReleaseInterface pAsync
  RenderPageToMemory = png
End Function

Private Function RenderPageToStream(ByVal pDoc As LongPtr, ByVal zeroIdx As Long) As LongPtr
  Dim pPage As LongPtr, pIStream As LongPtr, pRAS As LongPtr, pAsync As LongPtr, g As GUID
  RenderPageToStream = 0

  InvokeInterface pDoc, IPdfDocument_GetPage, vbLong, zeroIdx, VarPtr(pPage)
  If pPage = 0 Then m_LastError = "GetPage returned null.": Exit Function

  pIStream = SHCreateMemStream(0, 0)
  If pIStream = 0 Then m_LastError = "SHCreateMemStream failed.": GoTo Cleanup

  GuidFromString IID_IRandomAccessStream, g
  If CreateRandomAccessStreamOverStream(pIStream, BSOS_DEFAULT, g, pRAS) <> S_OK Or pRAS = 0 Then
    m_LastError = "CreateRandomAccessStreamOverStream failed.": GoTo Cleanup
  End If

  InvokeInterface pPage, IPdfPage_RenderToStreamAsync, vbLong, pRAS, VarPtr(pAsync)
  If pAsync = 0 Then m_LastError = "RenderToStreamAsync returned null.": GoTo Cleanup
  If Not AwaitAsync(pAsync, expectResult:=False) Then m_LastError = "Render async failed.": GoTo Cleanup

  RenderPageToStream = pRAS
  pRAS = 0 ' Don't release it, caller owns it now

Cleanup:
  If pRAS <> 0 Then CloseStream pRAS
  ReleaseInterface pIStream
  ReleaseInterface pPage
  ReleaseInterface pAsync
End Function

Private Function LoadDocument(ByVal pdfPath As String) As LongPtr
  Dim pStream As LongPtr, pAsync As LongPtr, pDoc As LongPtr
  pStream = OpenReadStream(pdfPath)
  If pStream = 0 Then m_LastError = "Could not open read stream.": Exit Function

  InvokeInterface m_Factory, IPdfDocStatics_LoadFromStreamAsync, vbLong, pStream, VarPtr(pAsync)
  If pAsync = 0 Then
    m_LastError = "LoadFromStreamAsync returned null."
  ElseIf AwaitAsync(pAsync, pDoc, expectResult:=True) Then
    LoadDocument = pDoc
  Else
    m_LastError = "Async document load failed."
  End If
  ReleaseInterface pStream
  ReleaseInterface pAsync
End Function

Private Function ResolvePages(ByVal pagesSelector As Variant, ByVal lastPage As Variant) As Long()
  Dim indices() As Long, count As Long, i As Long
  Erase indices
  If IsMissing(pagesSelector) Or IsEmpty(pagesSelector) Then
    ReDim indices(0 To m_PageCount - 1)
    For i = 0 To m_PageCount - 1
      indices(i) = i
    Next i
  ElseIf IsArray(pagesSelector) Then
    Dim v As Variant
    ReDim indices(0 To UBound(pagesSelector) - LBound(pagesSelector))
    For Each v In pagesSelector
      Dim pg As Long: pg = CLng(v)
      If pg >= 1 And pg <= m_PageCount Then
        indices(count) = pg - 1
        count = count + 1
      End If
    Next v
    If count = 0 Then
      m_LastError = "No valid page numbers supplied."
      Exit Function
    End If
    ReDim Preserve indices(0 To count - 1)
  Else
    Dim startPage As Long, endPage As Long
    startPage = CLng(pagesSelector)
    endPage = IIf(IsMissing(lastPage) Or IsEmpty(lastPage), startPage, CLng(lastPage))
    
    If startPage < 1 Then startPage = 1
    If endPage > m_PageCount Then endPage = m_PageCount
    If startPage > endPage Then
      m_LastError = "Invalid page range."
      Exit Function
    End If
      
    ReDim indices(0 To endPage - startPage)
    For i = startPage To endPage
      indices(count) = i - 1
      count = count + 1
    Next i
  End If
  ResolvePages = indices
End Function

' ================================
'  Private - COM & Stream Helpers
' ================================

Private Function AwaitAsync(ByVal pAsync As LongPtr, Optional ByRef pResult As LongPtr, Optional ByVal expectResult As Boolean = True) As Boolean
  If pAsync = 0 Then Exit Function
  Dim iid As GUID, pAsyncInfo As LongPtr, status As Long, spin As Long
  GuidFromString IID_IAsyncInfo, iid
  InvokeInterface pAsync, IUnknown_QueryInterface, vbLong, VarPtr(iid), VarPtr(pAsyncInfo)
  If pAsyncInfo = 0 Then Exit Function

  Do
    DoEvents
    InvokeInterface pAsyncInfo, IAsyncInfo_GetStatus, vbLong, VarPtr(status)
    spin = spin + 1
    If spin > ASYNC_MAX_SPIN Then Exit Do
  Loop While status = ASYNC_STATUS_STARTED

  If status = ASYNC_STATUS_COMPLETED Then
    If expectResult Then
      InvokeInterface pAsync, IAsyncOperation_GetResults, vbLong, VarPtr(pResult)
      AwaitAsync = (pResult <> 0)
    Else
      AwaitAsync = True
    End If
  Else
    InvokeInterface pAsyncInfo, IAsyncInfo_Close, vbLong
  End If
  ReleaseInterface pAsyncInfo
End Function

Private Sub IStreamSeek(ByVal pStream As LongPtr, ByVal Offset As Long, ByVal Origin As Long, Optional ByRef newPos As Currency)
  Dim result As Variant, posOut As Currency, offsetCY As Currency, types(2) As Integer, ptrs(2) As LongPtr, vars(2) As Variant
  offsetCY = CCur(Offset) / 10000
  types(0) = vbCurrency: vars(0) = offsetCY: ptrs(0) = VarPtr(vars(0))
  types(1) = vbLong: vars(1) = CLng(Origin): ptrs(1) = VarPtr(vars(1))
  types(2) = vbLongPtr: vars(2) = VarPtr(posOut): ptrs(2) = VarPtr(vars(2))
  If DispCallFunc(pStream, ISTREAM_SLOT_SEEK * PTR_SIZE, CC_STDCALL, vbLong, 3, VarPtr(types(0)), VarPtr(ptrs(0)), result) = S_OK Then
    newPos = posOut * 10000
  End If
End Sub

Private Function IStreamRead(ByVal pStream As LongPtr, ByRef buffer As Byte, ByVal cbRequest As Long, ByRef bytesRead As Long) As Boolean
  Dim result As Variant, types(2) As Integer, ptrs(2) As LongPtr, vars(2) As Variant
  types(0) = vbLongPtr: vars(0) = VarPtr(buffer): ptrs(0) = VarPtr(vars(0))
  types(1) = vbLong: vars(1) = cbRequest: ptrs(1) = VarPtr(vars(1))
  types(2) = vbLongPtr: vars(2) = VarPtr(bytesRead): ptrs(2) = VarPtr(vars(2))
  IStreamRead = (DispCallFunc(pStream, ISTREAM_SLOT_READ * PTR_SIZE, CC_STDCALL, vbLong, 3, VarPtr(types(0)), VarPtr(ptrs(0)), result) = S_OK)
End Function

Private Function OpenReadStream(ByVal FilePath As String) As LongPtr
  Dim g As GUID, pStream As LongPtr
  GuidFromString IID_IRandomAccessStream, g
  If CreateRandomAccessStreamOnFile(StrPtr(FilePath), FILE_ACCESS_READ, g, pStream) = S_OK Then OpenReadStream = pStream
End Function

Private Function WriteByteArrayToFile(ByVal FilePath As String, ByRef data() As Byte) As Boolean
  Dim hFile As LongPtr, written As Long, cbData As Long
  cbData = UBound(data) - LBound(data) + 1
  hFile = CreateFileW(StrPtr(FilePath), GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = 0 Then m_LastError = "CreateFileW failed.": Exit Function
  If WriteFile(hFile, data(LBound(data)), cbData, written, 0) = 0 Or written <> cbData Then
    m_LastError = "WriteFile failed."
    CloseHandle hFile
    Exit Function
  End If
  CloseHandle hFile
  WriteByteArrayToFile = True
End Function

Private Function ActivateFactory(ByVal className As String, ByVal sIID As String) As LongPtr
  Dim hStr As LongPtr, g As GUID, pFact As LongPtr
  WindowsCreateString StrPtr(className), Len(className), hStr
  If hStr = 0 Then Exit Function
  GuidFromString sIID, g
  RoGetActivationFactory hStr, g, pFact
  WindowsDeleteString hStr
  ActivateFactory = pFact
End Function

Private Sub GuidFromString(ByVal sGUID As String, ByRef g As GUID)
  CLSIDFromString StrPtr(sGUID), g
End Sub

Private Function HStringToString(ByVal hString As LongPtr) As String
  Dim pBuffer As LongPtr, Length As Long, bytes() As Byte
  pBuffer = WindowsGetStringRawBuffer(hString, Length)
  If pBuffer = 0 Or Length = 0 Then Exit Function
  ReDim bytes(0 To Length * 2 - 1)
  CopyMemory bytes(0), ByVal pBuffer, Length * 2
  HStringToString = bytes
End Function

Private Function InvokeInterface(ByVal pInterface As LongPtr, ByVal slotIndex As Long, ByVal returnType As Long, ParamArray args() As Variant) As Variant
  If pInterface = 0 Or slotIndex < 0 Then Exit Function
  If (returnType And &HFFFF0000) <> 0 Then Exit Function

  Dim byteOffset As Long, argCount As Long, result As Variant, hr As Long, i As Long
  byteOffset = slotIndex * PTR_SIZE

  Dim argVals() As Variant: argVals = args
  argCount = Abs(UBound(argVals) - LBound(argVals) + 1)

  Dim argPtrs() As LongPtr, argTypes() As Integer
  If argCount = 0 Then
    ReDim argPtrs(0): ReDim argTypes(0)
  Else
    ReDim argPtrs(0 To argCount - 1): ReDim argTypes(0 To argCount - 1)
    For i = 0 To argCount - 1
      argPtrs(i) = VarPtr(argVals(i))
      argTypes(i) = VarType(argVals(i))
    Next i
  End If

  hr = DispCallFunc(pInterface, byteOffset, CC_STDCALL, returnType, argCount, VarPtr(argTypes(0)), VarPtr(argPtrs(0)), result)

  If hr = S_OK Then
    InvokeInterface = result
  Else
    m_LastError = "InvokeInterface failed: slot " & slotIndex & ", HRESULT 0x" & Hex$(hr)
    DebugPrint m_LastError
  End If
End Function

Private Sub ReleaseInterface(ByRef p As LongPtr)
  If p <> 0 Then
    InvokeInterface p, IUnknown_Release, vbLong
    p = 0
  End If
End Sub

' ===========================
'  Private - File IO Helpers
' ===========================
Private Function FileExists(ByVal path As String) As Boolean
  Dim attrs As Long
  attrs = GetFileAttributesW(StrPtr(path))
  If attrs = INVALID_FILE_ATTRIBUTES Then Exit Function
  FileExists = ((attrs And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

Private Function FolderExists(ByVal path As String) As Boolean
  Dim attrs As Long
  attrs = GetFileAttributesW(StrPtr(path))
  If attrs = INVALID_FILE_ATTRIBUTES Then Exit Function
  FolderExists = ((attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function

Private Function FileBasename(ByVal FilePath As String) As String
  Dim parts() As String: parts = Split(FilePath, "\")
  Dim name As String: name = parts(UBound(parts))
  parts = Split(name, ".")
  If UBound(parts) > 0 Then ReDim Preserve parts(0 To UBound(parts) - 1)
  FileBasename = Join(parts, ".")
End Function

Private Function ParentFolder(ByVal FilePath As String) As String
  Dim parts() As String: parts = Split(FilePath, "\")
  ReDim Preserve parts(0 To UBound(parts) - 1)
  ParentFolder = Join(parts, "\") & "\"
End Function

Private Function TrailingSlash(ByVal folderPath As String) As String
  If Len(folderPath) = 0 Then Exit Function
  If Right$(folderPath, 1) <> "\" Then
    TrailingSlash = folderPath & "\"
  Else
    TrailingSlash = folderPath
  End If
End Function

Private Sub DebugPrint(ByVal msg As String)
  If m_DebugMode Then Debug.Print "[ChibiEx] " & msg
End Sub
