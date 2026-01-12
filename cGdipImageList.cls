VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdipImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cGdipImageList
' Author    : The vast majority from beededea but the basic class-boilerplate and some bug-fixing from chatGPT
' Date      : 04/01/2026
' Purpose   : Replacement of the RichClient imageList using the scripting.dictionary and GDI+ to store and pull VB6 native type images.
'             JPG, BMP, PNG tested so far. PNGs will display correctly with transparency if displayed using GDI+.
'             Alpha PNGs will display with a black background on any standard VB6 control but should display perfectly
'             using TwinBasic's native controls that have automatic PNG support built in.
'
'             Why is this ImageList useful?

'             * It is a Dictionary-backed ImageList, when using Cristian Buse's dictionary replacement, no scripting runtime dependency.
'             * It can store JPG, BMP, PNG files and others too.
'             * It can store images with alpha transparency as GDI+ provides this capability
'             * It can store images of varying size, not just small 16x16 or 32x32 icons as per the old VB6 imageList.
'             * Avoids RichClient dependency.
'             * Avoids runtime obsolescence.
'             * Uses dependable GDI+ to load and unload the images.
'             * You have full control as it is FOSS.
'             * The usage syntax is very similar to Rich Client's for easy drop-in replacement.
'             * It is quicker to pull images from the dictionary than directly from file.
'             * If used with Elroy's standard picture Ex project it can parse and render alpha images (PNGs &c) directly to VB6 picture/imageboxes.
'
'             Limitations?
'
'             * You can't currently use this imageList with RichClient as Olaf's code is designed specifically to work with his own imageList, eg. CC.RenderSurfaceContent
'             * VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project.
'
'             Usage:
'
'             Add a public or private variable to a module (BAS) in order to create a new GDI+ image list instance.
'             Public thisImageList As New cGdipImageList
'
'             Add a public variable to a module (BAS) to provide an instance counter for each usage of the class.
'             Public gGdipImageListInstanceCount As Long
'
'             thisImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
'             Set imgAbout.Picture = thisImageList.Picture("about-icon-dark")
'             thisImageList.Remove "about-icon-dark"
'             dictionaryCount = thisImageList.count
'             If thisImageList.Exists(thiskey) Then ...
'             thisImageList.ImageWidth = 150 ' sets the imageWidth, a value of zero causes the image's real width to be used
'             thisImageList.ImageHeight ' ditto
'             thisImageList.ImageOpacity = 100 ' 0-100% prior to loading will set the opacity
'             thisImageList.Clear  ' clears the imageList
'
'---------------------------------------------------------------------------------------

Option Explicit

' ===========================
' Win32 / GDI
' ===========================
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
     ByRef pbmi As BITMAPINFO, _
     ByVal iUsage As Long, _
     ByRef ppvBits As Long, _
     ByVal hSection As Long, _
     ByVal dwOffset As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObj As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObj As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" _
    (ByRef PicDesc As PICTDESC, _
     ByRef RefIID As GUID, _
     ByVal fPictureOwnsHandle As Long, _
     ByRef IPic As StdPicture) As Long

' ===========================
' GDI+
' ===========================
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" _
    (ByVal filename As Long, ByRef bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "GDIPlus" _
    (ByVal image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal pStream As Long, image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal image As Long, height As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal dx As Long, ByVal dy As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal pScanData As Long, image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal img As Long, Context As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GDIPlus" (ByVal Context As Long, ByVal PixOffsetMode As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMode As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "GDIPlus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "GDIPlus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As Any, grayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal Context As Long, ByVal image As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal CallbackData As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "GDIPlus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "GDIPlus" (ByVal numEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal Token As Long) As Long
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, gdipInput As GDIPLUS_STARTINPUT, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long

' APIs image cropping
Private Declare Function GdipGetImagePixelFormat Lib "GDIPlus" (ByVal image As Long, ByRef PixelFormat As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "GDIPlus" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal height As Long, ByVal PixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus

Private Declare Function GdipCreateFromHDC Lib "GDIPlus" _
    (ByVal hDC As Long, ByRef graphics As Long) As Long

Private Declare Function GdipDrawImageRectI Lib "GDIPlus" _
    (ByVal graphics As Long, _
     ByVal image As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal Width As Long, _
     ByVal height As Long) As Long
     
' Windows constants Start

Private Const PixelFormat32bppPARGB = &HE200B
Private Const PixelFormat32bppARGB = &H26200A

' global GDI+ Types START
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Private Type BLENDFUNCTION
'    BlendOp As Byte
'    BlendFlags As Byte
'    SourceConstantAlpha As Byte
'    AlphaFormat As Byte
'End Type

Private Type GDIPLUS_STARTINPUT
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 0) As Long
End Type

Private Type CLSID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
   ClassID As CLSID
   FormatID As CLSID
   CodecName As Long      ' String Pointer; const WCHAR*
   DllName As Long        ' String Pointer; const WCHAR*
   FormatDescription As Long ' String Pointer; const WCHAR*
   FilenameExtension As Long ' String Pointer; const WCHAR*
   MimeType As Long       ' String Pointer; const WCHAR*
   Flags As ImageCodecFlags   ' Should be a Long equivalent
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long      ' Byte Array Pointer; BYTE*
   SigMask As Long         ' Byte Array Pointer; BYTE*
End Type

Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type
' global GDI+ Types END



Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type



' global GDI+ Enums START
Private Enum GDIPLUS_UNIT
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

' NOTE: Enums evaluate to a Long
Private Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
   ProfileNotFound = 21
End Enum

' Information flags about image codecs
Private Enum ImageCodecFlags
   ImageCodecFlagsEncoder = &H1
   ImageCodecFlagsDecoder = &H2
   ImageCodecFlagsSupportBitmap = &H4
   ImageCodecFlagsSupportVector = &H8
   ImageCodecFlagsSeekableEncode = &H10
   ImageCodecFlagsBlockingDecode = &H20
   ImageCodecFlagsBuiltin = &H10000
   ImageCodecFlagsSystem = &H20000
   ImageCodecFlagsUser = &H40000
End Enum

Public Enum SmoothingModeEnum
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = 4&
    SmoothingModeAntiAlias8x8 = 5&
End Enum

' Enum vars for GDI+ colour matrix STARTS
Private Enum ColorAdjustType
   ColorAdjustTypeDefault
   ColorAdjustTypeBitmap
   ColorAdjustTypeBrush
   ColorAdjustTypePen
   ColorAdjustTypeText
   ColorAdjustTypeCount
   ColorAdjustTypeAny
End Enum
 
Private Enum ColorMatrixFlags
   ColorMatrixFlagsDefault = 0
   ColorMatrixFlagsSkipGrays = 1
   ColorMatrixFlagsAltGray = 2
End Enum
' Enum vars for GDI+ colour matrix ENDS

Private lngGDI As Long
Private gdipInit As GDIPLUS_STARTINPUT

'Private mDict As Object ' You will need a project reference to the Scripting.Dictionary

'-----------------------------------------------------------
' Declarations for the scripting.dictionary replacement START
'-----------------------------------------------------------

Private Type JustMyUDT
    dict As Dictionary
End Type

#Const x64StackIssues = ((Mac = 0) And Win64 And (twinbasic = 0))

#If twinbasic Then
    Private mDict As Collection ' std TB collection
#Else
    Private mDict As Dictionary ' scripting.dictionary
#End If
'-----------------------------------------------------------
' Declarations for the scripting.dictionary replacements END
'-----------------------------------------------------------

' member property variables
Private mImageWidth As Long
Private mImageHeight As Long
Private mImageOpacity As Integer

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Error
    
    'Set mDict = CreateObject("Scripting.Dictionary")
    
    Set mDict = New Dictionary ' if you are using Cristian Buse's Dictionary replacement for the Scripting.Dictionary
    
    If gGdipImageListInstanceCount = 0 Then Call initialiseGDIPStartup
    
    gGdipImageListInstanceCount = gGdipImageListInstanceCount + 1
    
    ' with image dimensions set to 0, the image's default dimensions will be extracted from the image itself
    mImageWidth = 0
    mImageHeight = 0
    ImageOpacity = 100
    
    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Error

    ' Clear the whole image list
    Clear
    
    Set mDict = Nothing
    
    gGdipImageListInstanceCount = gGdipImageListInstanceCount - 1
        
    ' Only Shutdown GDI+ when there are no users of this class left
    If lngGDI Then
        If gGdipImageListInstanceCount = 0 Then Call GdiplusShutdown(lngGDI)
    End If
    
    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module cGdipImageList"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Obtain the current count of items in the dictionary
'---------------------------------------------------------------------------------------
'
Public Property Get count() As Long
    On Error GoTo Count_Error

    count = mDict.count

    On Error GoTo 0
    Exit Property

Count_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Count of Class Module cGdipImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ImageWidth() As Long
    On Error GoTo ImageWidth_Error

    ImageWidth = mImageWidth

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cGdipImageList"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ImageWidth
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ImageWidth(ByVal newValue As Long)
    On Error GoTo ImageWidth_Error

    If mImageWidth <> newValue Then mImageWidth = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageWidth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageWidth of Class Module cGdipImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageHeight
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ImageHeight() As Long
    On Error GoTo ImageHeight_Error

    ImageHeight = mImageHeight

    On Error GoTo 0
    Exit Property

ImageHeight_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cGdipImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImageHeight
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ImageHeight(ByVal newValue As Long)
    On Error GoTo ImageHeight_Error

    If mImageHeight <> newValue Then mImageHeight = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageHeight_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageHeight of Class Module cGdipImageList"
End Property


'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not much use when using VB6 resulting in a black background
'             better for use with TwinBasic
'---------------------------------------------------------------------------------------
'
Public Property Get ImageOpacity() As Integer
    On Error GoTo ImageOpacity_Error

    ImageOpacity = mImageOpacity

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cGdipImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImageOpacity
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : image Opacity is not much use when using VB6 resulting in a black background
'             better for use with TwinBasic
'---------------------------------------------------------------------------------------
'
Public Property Let ImageOpacity(ByVal newValue As Integer)
    On Error GoTo ImageOpacity_Error

    If mImageOpacity <> newValue Then mImageOpacity = newValue Else Exit Property

    On Error GoTo 0
    Exit Property

ImageOpacity_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageOpacity of Class Module cGdipImageList"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Test to see if the item exists in the dictionary
'---------------------------------------------------------------------------------------
'
Public Function Exists(ByVal key As String) As Boolean
    On Error GoTo Exists_Error

    Exists = mDict.Exists(key)

    On Error GoTo 0
    Exit Function

Exists_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Exists of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Remove a single item from the dictionary
' Note      : Dictionary stores GDI+ Bitmap handles (GpBitmap)
'             Must be destroyed with GdipDisposeImage
'---------------------------------------------------------------------------------------
'
Public Sub Remove(ByVal key As String)
    On Error GoTo Remove_Error

    If mDict.Exists(key) Then
        ' GDI+ delete
        GdipDisposeImage mDict(key) 'NOT DeleteObject mDict(Key)
        mDict.Remove key
    End If

    On Error GoTo 0
    Exit Sub

Remove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Clear the whole image list
' Note      : Dictionary stores GDI+ Bitmap handles (GpBitmap)
'             Must be destroyed with GdipDisposeImage
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Dim k As Variant
    On Error GoTo Clear_Error

    For Each k In mDict.Keys
        ' GDI+ delete
        GdipDisposeImage mDict(k) 'NOT DeleteObject mDict(k)
    Next
    mDict.RemoveAll

    On Error GoTo 0
    Exit Sub

Clear_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clear of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetImage
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Get a specific image by key from the dictionary
' Returns   : GpBitmap (NOT HBITMAP)
'---------------------------------------------------------------------------------------
'
Public Function GetImage(ByVal key As String) As Long
    On Error GoTo GetImage_Error

    If Not mDict.Exists(key) Then Err.Raise 5, , "Image key not found"
    GetImage = mDict(key)

    On Error GoTo 0
    Exit Function

GetImage_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetImage of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddImage
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : Adds an image to the dictionary
'---------------------------------------------------------------------------------------
'
Public Sub AddImage(ByVal key As String, ByVal FilePath As String)

    On Error GoTo AddImage_Error

    If mDict.Exists(key) Then Err.Raise 457, , "Key already exists"
    
    If fFExists(FilePath) Then
        resizeAndLoadImgToDict mDict, key, FilePath, mImageWidth, mImageHeight, , mImageOpacity
    Else
        MsgBox " File not found " & FilePath
    End If

    On Error GoTo 0
    Exit Sub

AddImage_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddImage of Class Module cGdipImageList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Picture
' Author    : beededea
' Date      : 04/01/2026
' Purpose   : extracts the image from the dictionary using similar syntax to Rich Client
'---------------------------------------------------------------------------------------
'
Public Function Picture(ByVal key As String) As StdPicture
    Dim gpImage As Long: gpImage = 0
    
    On Error GoTo Picture_Error

    If Not mDict.Exists(key) Then Err.Raise 5, , "Image key not found"
    
    gpImage = GetImage(key)
    Set Picture = fBmpToStdPicture(gpImage, mImageWidth, mImageHeight)

    On Error GoTo 0
    Exit Function

Picture_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Picture of Class Module cGdipImageList"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fBmpToStdPicture
' Author    : beededea
' Date      : 05/01/2026
' Purpose   : Obtain a stdPicture handle to a bitmap image from the collection
'---------------------------------------------------------------------------------------
'
Public Function fBmpToStdPicture(ByVal lhImageBitmap As Long, ByVal Width As Long, ByVal height As Long) As StdPicture

    Dim hDC As Long: hDC = 0
    Dim hDibBitmap As Long: hDibBitmap = 0
    Dim hOld As Long: hOld = 0
    Dim lhGraphics As Long: lhGraphics = 0
    Dim dx As Long: dx = 0
    Dim dy As Long: dy = 0
    
    On Error GoTo fBmpToStdPicture_Error
    
    ' Initialises the hDC to draw upon, a handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    ' Get a device context compatible with the screen
    hDC = CreateCompatibleDC(ByVal 0&)

    'GDI+ API to determine image dimensions - note that using GDIP to resize can result in ghost boundaries or 'edging'
    Call GdipGetImageWidth(lhImageBitmap, dx) ' cairo_image_surface_get_width ' the width of the surface in pixels.
    If Width <= 0 Then Width = dx ' if no supplied width then use the original width
    
    Call GdipGetImageHeight(lhImageBitmap, dy) ' cairo_image_surface_get_height ()
    If height <= 0 Then height = dy

    ' create an alpha channel RGB DIB target bitmap
    hDibBitmap = fCreateNewGdipDIBsection(Width, height, hDC)

    If hDibBitmap <> 0 Then
        ' select the bitmap into the memory DC
        ' make the device context hDC use the bitmap.  hOld is a return value giving a handle which determines success and allows reverting later to release GDI handles
        hOld = SelectObject(hDC, hDibBitmap)
    Else
      DeleteDC hDC
      Err.Raise 1, "InitDC", "Bitmap creation failed"
    End If

    ' Create GDI+ graphics, wrap target DC in GDI+, creating a GDIP graphic object with a pointer 'lhGraphics' using a handle to the bitmap graphic section assigned to the device context
    GdipCreateFromHDC hDC, lhGraphics

    ' Draw image (scaled if required)
    GdipDrawImageRectI lhGraphics, lhImageBitmap, 0, 0, Width, height

    ' Convert hDibBitmap to StdPicture
    Set fBmpToStdPicture = fOleCreatePicFromHBitmap(hDibBitmap) ' return

    ' StdPicture now owns the bitmap
    
    GoTo fBmpToStdPicture_CleanUp

fBmpToStdPicture_Error:

    On Error GoTo 0
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fBmpToStdPicture of Class Module cGdipImageList"
     
fBmpToStdPicture_CleanUp:
    
    ' Cleanup GDI+
    GdipDeleteGraphics lhGraphics
    SelectObject hDC, hOld
    DeleteDC hDC

    Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : fCreateNewGdipDIBsection
' Author    : beededea
' Date      : 05/01/2026
' Purpose   : create an alpha channel RGB DIB bitmap
'---------------------------------------------------------------------------------------
'
Private Function fCreateNewGdipDIBsection(ByVal cx As Long, ByVal cy As Long, ByRef hDC As Long) As Long
    Dim bmi As BITMAPINFO
    Dim bits As Long: bits = 0
    Dim hBmpMemory As Long: hBmpMemory = 0
    
    On Error GoTo fCreateNewGdipDIBsection_Error

    'load the bitmap information with pertinent properties
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = cx
        .biHeight = -cy           ' top-down DIB
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0        ' BI_RGB
    End With
    
    ' create a device independent bitmap and return a handle, hBmpMemory, providing it a handle to device context allocated memory previously created with CreateCompatibleDC,
    ' providing size information in bmpInfo and setting any attributes to the new bitmap
    hBmpMemory = CreateDIBSection(hDC, bmi, 0, bits, 0, 0)
    
    fCreateNewGdipDIBsection = hBmpMemory ' return

    On Error GoTo 0
    Exit Function

fCreateNewGdipDIBsection_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fCreateNewGdipDIBsection of Class Module cGdipImageList"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fOleCreatePicFromHBitmap
' Author    : beededea
' Date      : 05/01/2026
' Purpose   : Creates a VB compatible stdPicture Object from a handle to a bitmap
'---------------------------------------------------------------------------------------
'
Private Function fOleCreatePicFromHBitmap(ByVal hBmp As Long) As StdPicture
    Dim pic As PICTDESC
    Dim IID_IPicture As GUID
    Dim oPicture As StdPicture

    On Error GoTo fOleCreatePicFromHBitmap_Error

    ' Initialize the PICTDESC structure
    With pic
        .cbSizeofStruct = Len(pic)
        .picType = 1 ' PICTYPE_BITMAP
        .hImage = hBmp
    End With

    ' Fill in OLE IDispatch Interface ID
    With IID_IPicture
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Creates a new picture object initialized according to a PICTDESC structure.
    OleCreatePictureIndirect pic, IID_IPicture, True, oPicture
    Set fOleCreatePicFromHBitmap = oPicture ' return
    
    On Error GoTo 0
    Exit Function

fOleCreatePicFromHBitmap_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fOleCreatePicFromHBitmap of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadBytesFromFile
' Author    : Credit to Olaf Schmidt
' Date      : 07/04/2020
' Purpose   : Reads the file as a stream of data using the ADODB.Stream object
'---------------------------------------------------------------------------------------
'
Private Function ReadBytesFromFile(ByVal filename As String) As Byte()
    Dim ab As Object
   
    On Error GoTo ReadBytesFromFile_Error
    
    ' COM object, which is used to represent a stream of data or text
    Set ab = CreateObject("ADODB.Stream")
    
    With ab
      .Open
        .Type = 1 'adTypeBinary
        .LoadFromFile filename
        ReadBytesFromFile = .Read
      .Close
    End With
  
   On Error GoTo 0
   Exit Function

ReadBytesFromFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadBytesFromFile of Class Module cGdipImageList"
End Function




'---------------------------------------------------------------------------------------
' Procedure : GetEncoderClsid
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'             The function calls GetImageEncoders to get an array of ImageCodecInfo objects. If one of the
'             ImageCodecInfo objects in that array represents the requested encoder, the function returns
'             the index of the ImageCodecInfo object and copies the CLSID into the variable pointed to by
'             pClsid. If the function fails, it returns –1.
'
' Built-in encoders for saving: (You can *try* to get other types also)
'   image/bmp
'   image/jpeg
'   image/gif
'   image/tiff
'   image/pngI received a
'
' Notes When Saving:
'The JPEG encoder supports the Transformation, Quality, LuminanceTable, and ChrominanceTable parameter categories.
'The TIFF encoder supports the Compression, ColorDepth, and SaveFlag parameter categories.
'The BMP, PNG, and GIF encoders no do not support additional parameters.
'
'---------------------------------------------------------------------------------------

Private Function GetEncoderClsid(strMimeType As String, ClassID As CLSID) As Long
   Dim num As Long: num = 0
   Dim Size As Long: Size = 0
   Dim i As Long: i = 0

   Dim ICI() As ImageCodecInfo
   Dim Buffer() As Byte
   
   On Error GoTo GetEncoderClsid_Error

   GetEncoderClsid = -1 'Failure flag

   ' Get the encoder array size
   Call GdipGetImageEncodersSize(num, Size)
   If Size = 0 Then Exit Function ' Failed!

   ' Allocate room for the arrays dynamically
   ReDim ICI(1 To num) As ImageCodecInfo
   ReDim Buffer(1 To Size) As Byte

   ' Get the array and string data
   Call GdipGetImageEncoders(num, Size, Buffer(1))
   
   ' Copy the class headers
   Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * num))

   ' Loop through all the codecs
   For i = 1 To num
      ' Must convert the pointer into a usable string
      If StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0 Then
         ClassID = ICI(i).ClassID   ' Save the class id
         GetEncoderClsid = i        ' return the index number for success
         Exit For
      End If
   Next
   
   ' Free the memory
   Erase ICI
   Erase Buffer

   On Error GoTo 0
   Exit Function

GetEncoderClsid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetEncoderClsid of Class Module cGdipImageList"
End Function



'---------------------------------------------------------------------------------------
' Procedure : PtrToStrW
' Author    : From www.mvps.org/vbnet...I think
' Date      : 21/08/2020
' Purpose   : Dereferences an ANSI or Unicode string pointer
'             and returns a normal VB BSTR used in GetEncoderClsid above
'---------------------------------------------------------------------------------------
'
Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim sOut As String: sOut = 0
    Dim lLen As Long: lLen = 0

    On Error GoTo PtrToStrW_Error

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If

    On Error GoTo 0
    Exit Function

PtrToStrW_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PtrToStrW of Class Module cGdipImageList"
End Function

'---------------------------------------------------------------------------------------
' Procedure : createGdipBitmap
' Author    : Credit to Olaf Schmidt for original idea
'             also to Joaquim https://www.vbforums.com/showthread.php?840601-RESOLVED-how-use-ColorMatrix
' Date      : 07/04/2020
' Purpose   : Creates the scaled image with quality and opacity attributes
'---------------------------------------------------------------------------------------
'
Private Function createGdipBitmap(SrcImg As Long, dxSrc As Long, dySrc As Long, dxDst As Long, dyDst As Long, Opacity As Integer) As Long
    Dim img As Long: img = 0
    Dim Ctx As Long: Ctx = 0
    Dim imgQuality As Long: imgQuality = 0
    Dim SmoothingMode As Long: SmoothingMode = 0
    Dim imgAttr As Long: imgAttr = 0
    Dim clrMatrix As ColorMatrix
    Dim graMatrix As ColorMatrix
    Dim sImageQuality As String: sImageQuality = vbNullString

    On Error GoTo createGdipBitmap_Error

    imgAttr = &H11
    sImageQuality = "2"
    
    'Setup the transformation matrix for alpha adjustment used in GdipSetImageAttributesColorMatrix below
    clrMatrix.m(0, 0) = 1
    clrMatrix.m(1, 1) = 1
    clrMatrix.m(2, 2) = 1
    clrMatrix.m(3, 3) = 1 * Opacity / 100 ' 0.5 'Alpha transform (50%) ' cairo_image_surface_create_for_data (BGRA = alpha)
    clrMatrix.m(4, 4) = 1

    ' Use the existing buffer pCairoBuf that has (iRenderWidth * iRenderHeight * 4) bytes, 4 bytes per pixel for ARGB32 and zero row padding
    ' cairo_surface_t *pSurface = cairo_image_surface_create_for_data(pCairoBuf, CAIRO_FORMAT_ARGB32, iRenderWidth, iRenderHeight, (iRenderWidth * 4))

    ' set the image quality and smoothing mode
    If sImageQuality = "0" Then
        imgQuality = &H1 '    ipmNearestNeighbor = &H5
        SmoothingMode = SmoothingModeNone
    End If
    If sImageQuality = "1" Then
        imgQuality = &H6 '    ipmHighQualityBiLinear = &H6
        SmoothingMode = SmoothingModeHighSpeed
    End If
    If sImageQuality = "2" Then
        imgQuality = &H7 '    ipmHighQualityBicubic = &H7
        SmoothingMode = SmoothingModeHighQuality
    End If
    
    'Creates an alpha RGB GDIP Bitmap object, ie. a surface (img) based on an array of bytes along with the destination size and format information img is the pointer to that bitmap object
    Call GdipCreateBitmapFromScan0(dxDst, dyDst, dxDst * 4, PixelFormat32bppPARGB, 0, img) ' Cairo.CreateSurface & Set_Device_Offset
    
    If img Then
        createGdipBitmap = img ' set the return value to the bitmap object at this point, we still have more to do to it.
        
        'Creates a Graphics object hw context - ctx, that is now associated with an Image bitmap object (surface).
        Call GdipGetImageGraphicsContext(img, Ctx)
    Else
        Err.Raise vbObjectError, , "unable to create scaled Img-Resource"
    End If
    
    ' now the hw context has been associated with the returned gdip bitmap, we can make changes to the context
    If Ctx Then
        ' set the quality of the context using three GDIP functions
        Call GdipSetPixelOffsetMode(Ctx, 3)            '     4=Half, 3=None
        Call GdipSetInterpolationMode(Ctx, imgQuality) ' three levels of quality
        Call GdipSetSmoothingMode(Ctx, SmoothingMode)  '          ditto
        
        ' Sets the compositing quality of this Graphics object when alpha blended. Speed vs quality. Used in conjunction with GdipSetCompositingMode
        'Call GdipSetCompositingQuality(Ctx, CompositingQualityHighQuality)  ' CompositingQualityHighSpeed
                                
        'Create storage for the image attributes struct used below
        Call GdipCreateImageAttributes(imgAttr)

        'Setup the image attributes into imgAttr struct using the transformation matrix and the enum values, ColorAdjustTypeBitmap and ColorMatrixFlagsDefault
        Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeBitmap, 1, clrMatrix, graMatrix, ColorMatrixFlagsDefault)

        ' draw the loaded source image scaled onto a generated image to the desired scale with the above image quality and opacity attributes
        If SrcImg <> 0 Then
            GdipDrawImageRectRectI Ctx, SrcImg, 0, 0, dxDst, dyDst, 0, 0, dxSrc, dySrc, 2, imgAttr, 0, 0 ' Cairo.Cairo_Surface /  cairo_image_surface_create_for_data (BGRA = alpha)
            ' Set CC = Cairo.CreateSurface(Me.ScaleWidth, Me.ScaleHeight).CreateContext ' sample code reference
        End If
        
        ' the image has now been drawn so we can now delete the now unwanted graphics context
        Call GdipDeleteGraphics(Ctx) ' cairo_destroy(cr) &  'cairo_surface_destroy(surface)
    End If

   On Error GoTo 0
   Exit Function

createGdipBitmap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createGdipBitmap of Class Module cGdipImageList"
End Function



'---------------------------------------------------------------------------------------
' Procedure : resizeAndLoadImgToDict
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : read the file as a series of bytes
'             Creates a stream object stored in global memory using the location address of the variable where the data resides
'             Creates a GDI+ Image object based on the stream, using GdipLoadImageFromStream
'             Finally, uses a function GdipCreateBitmapFromScan0 to both create and resize the image.
'---------------------------------------------------------------------------------------
'
Private Sub resizeAndLoadImgToDict(ByRef thisDictionary As Object, ByVal key As String, ByVal strFilename As String, ByVal Width As Long, ByVal height As Long, Optional ByVal fullStringKey As String = "", Optional ByVal ImageOpacity As Integer)

    Dim thiskey As String: thiskey = 0
    Dim encoderCLSID As CLSID
    Dim bytesFromFile() As Byte
    Dim Strm As stdole.IUnknown
    Dim img As Long: img = 0
    Dim dx As Long: dx = 0
    Dim dy As Long: dy = 0
    Dim picBitmap As Long: picBitmap = 0

    On Error GoTo resizeAndLoadImgToDict_Error
        
    ' Get the CLSID of the PNG encoder
    Call GetEncoderClsid("image/png", encoderCLSID)
    
    ' uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes
    bytesFromFile = ReadBytesFromFile(strFilename)

    ' creates a stream object stored in global memory using the location address of the variable where the data resides, Olaf Schmidt
    Call CreateStreamOnHGlobal(VarPtr(bytesFromFile(0)), 0, Strm)
    
    ' Creates a GDI+ Image object based on the stream, loads it into img - Olaf Schmidt
    Call GdipLoadImageFromStream(ObjPtr(Strm), img)       ' cairo_image_surface_create_from_png (const char *filename);
    If img = 0 Then Err.Raise vbObjectError, , "Could not load image " & strFilename & " with GDIPlus"

    'GDI+ API to determine image dimensions, Olaf Schmidt
    Call GdipGetImageWidth(img, dx) ' cairo_image_surface_get_width ' the width of the surface in pixels.
    If Width <= 0 Then Width = dx ' if no supplied width then use the original width
    
    Call GdipGetImageHeight(img, dy) ' cairo_image_surface_get_height ()
    If height <= 0 Then height = dy
        
    ' create a scaled GDI+ bitmap of the image surface using GdipCreateBitmapFromScan0 and context quality attributes
    picBitmap = createGdipBitmap(img, dx, dy, Width, height, ImageOpacity)

    'override any key
    If fullStringKey = "" Then
        ' create a unique key string
        thiskey = key
    Else
        thiskey = fullStringKey
    End If
    
    ' add the bitmap to the dictionary collection
    If thisDictionary.Exists(thiskey) Then
        thisDictionary.Remove thiskey
    End If
    thisDictionary.Add thiskey, picBitmap

   On Error GoTo 0
   Exit Sub

resizeAndLoadImgToDict_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure resizeAndLoadImgToDict with this file - " & strFilename & " of Class Module cGdipImageList"
        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : initialiseGDIPStartup
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : Initialises GDI Plus
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGDIPStartup()

   On Error GoTo initialiseGDIPStartup_Error
   
    gdipInit.GdiplusVersion = 1
    If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
        MsgBox "Error loading GDI+", vbCritical
    End If

   On Error GoTo 0
   Exit Sub

initialiseGDIPStartup_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGDIPStartup of Form dock"

End Sub



