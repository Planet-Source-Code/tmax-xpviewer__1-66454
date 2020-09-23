VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIMAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What does this do ?
'-------------------
'Used to read an image using gdi+
'can also resize that image and save it


'-----------------------------
'FOR MAKING THE PICTURE OBJECT
'-----------------------------
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PICTDESC
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CLSIDFromString Lib "ole32" ( _
   ByVal str As Long, _
   id As GUID) As Long

'--------------------------
'FOR MAKING SAVEAS FILEPATH
'--------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'-------
'MY VARS
'-------
Private m_hGDIplus As Long      'token to GDI engine
Private m_Image As Long         'original file loaded
Private m_Graphic As Long       'adjusted image
Private m_Bitmap As Long        'image to save
Enum Rot_Angle
 Deg90 = 1
 Deg180 = 2
 Deg270 = 3
 End Enum
'-----------------
'PUBLIC PROPERTIES
'-----------------
Public Property Get ImageWidth() As Single
    'the width of the active image
    Dim sngHeight As Single
    If m_Bitmap = 0 Then
        GdipGetImageDimension m_Image, ImageWidth, sngHeight
    Else
        GdipGetImageDimension m_Bitmap, ImageWidth, sngHeight
    End If
End Property
Public Property Get ImageHeight() As Single
    'the height of the active image
    Dim sngWidth As Single
    If m_Bitmap = 0 Then
        GdipGetImageDimension m_Image, sngWidth, ImageHeight
    Else
        GdipGetImageDimension m_Bitmap, sngWidth, ImageHeight
    End If
End Property
Public Property Get Initialised() As Boolean
    'used to check that gdi engine was started ok
    If m_hGDIplus <> 0 Then
        Initialised = True
    End If
End Property

'-----------------
'PUBLIC FUNCTIONS
'-----------------

Public Sub Thumbnail(FileName$, Optional ByVal LWidth As Long = 100, Optional ByVal LHeight As Long = 100)
   Dim img As Long
   Dim lngHeight As Long, lngWidth As Long
  
   'GdipLoadImageFromFile StrConv(Filename, vbUnicode), img  ' Load the image
   GdipLoadImageFromFile FileName, img  ' Load the image
    On Error GoTo LoadImgErr
   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth
    If lngHeight < lngWidth Then
   ' Create the thumbnail that is 100x100 in size
        GdipGetImageThumbnail img, LWidth, lngHeight / lngWidth * LHeight, m_Image
   Else
        GdipGetImageThumbnail img, lngWidth / lngHeight * LWidth, LHeight, m_Image
   End If
      GdipDrawImageRect m_Graphic, m_Image, 0, 0, LWidth, LHeight
   ' Cleanup
   
   GdipDisposeImage img ' Delete the image
   Exit Sub
LoadImgErr:
   GdipDisposeImage img ' Delete the image
End Sub

Public Function Load(sFileName As String) As Boolean
    Dim retval As Long
    Dispose
    retval = GdipLoadImageFromFile(sFileName, m_Image)
    If retval = 0 Then
        Load = True
    End If
End Function

Public Function PaintDC(ByVal hDC As Long, x As Integer, y As Integer)
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim graphic As Long
    
    'get a link to the DC we are going to paint onto
    GdipCreateFromHDC hDC, graphic
    
    'draw onto it
    If m_Bitmap = 0 Then
        GdipGetImageDimension m_Image, sngWidth, sngHeight
        GdipDrawImageRect graphic, m_Image, x, y, sngWidth, sngHeight
    Else
        GdipGetImageDimension m_Bitmap, sngWidth, sngHeight
        GdipDrawImageRect graphic, m_Bitmap, x, y, sngWidth, sngHeight
    End If
    
    GdipDeleteGraphics graphic
End Function

Public Function ReSize(ByVal W As Single, ByVal H As Single, Optional ByVal bGrow As Boolean) As Boolean
    Dim sngScale As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim lPixelFormat As Long
    Dim scaleMode As Long
    
'--clear up
    DisposeGraphic
    If (W <= 0) And (H <= 0) Then Exit Function
    If m_Image = 0 Then Exit Function

'--defaults
    scaleMode = InterpolationModeHighQualityBilinear
    
'--Establish resize measurements
    'find out what we have now
    GdipGetImageDimension m_Image, sngWidth, sngHeight
    'bail if its a bad image
    If (sngWidth <= 0) Or (sngHeight <= 0) Then Exit Function
    'get the existing scale
    sngScale = sngWidth / sngHeight
    
    'how do we change the new dimensions if at all?
    'if either one of the values is 0 then use the other to set the scale
    If W = 0 Then W = H * sngScale
    If H = 0 Then H = W / sngScale
    
    'images will look shite if they are grown
    'so check if we want this to happen
    If (H > sngHeight) Or (W > sngWidth) Then
        If bGrow Then
            'thats ok, let them grow the image
            scaleMode = InterpolationModeHighQualityBicubic
        Else
            'reset it to its original size
            H = sngHeight
            W = sngWidth
        End If
    End If
    
'--get pixel format to use
    'use the same pixel format as the image we are copying
    GdipGetImagePixelFormat m_Image, lPixelFormat
    'or hard code it to all be the same
    'lPixelFormat = PixelFormat32bppARGB

'--Resizing
    'create a new bitmap for our picture
    'PixelFormat32bppARGB
    GdipCreateBitmapFromScan0 W, H, 0, lPixelFormat, ByVal 0&, m_Bitmap
    
    'get a handle to the graphics object of our new bitmap
    GdipGetImageGraphicsContext m_Bitmap, m_Graphic
    'could use something like this if you want it to be displayed on screen
    'GdipCreateFromHDC formGDItest.Picture1.hDC, m_Graphic
   
    'set the method that we want to use to do this transformation
    GdipGetInterpolationMode m_Graphic, scaleMode
    
    'this isnt necessary, but may give a better quality.
    'need to look at it further
    'smoothing = antialising?
    'GdipGetSmoothingMode m_Graphic, SmoothingModeHighQuality
    
    'copy our image into the graphic object with the desired dimensions
    GdipDrawImageRectRectI m_Graphic, _
                            m_Image, _
                            0, 0, W, H, _
                            0, 0, sngWidth, sngHeight, _
                            UnitPixel
                    
    ReSize = True
End Function

Public Function GdiErrorString(ByVal lError As GpStatus) As String
  Dim S As String
'Private Enum GDIpStatus   ' aka Status
'   Ok = 0
'   GenericError = 1
'   InvalidParameter = 2
'   OutOfMemory = 3
'   ObjectBusy = 4
'   InsufficientBuffer = 5
'   NotImplemented = 6
'   Win32Error = 7
'   WrongState = 8
'   Aborted = 9
'   FileNotFound = 10
'   ValueOverflow = 11
'   AccessDenied = 12
'   UnknownImageFormat = 13
'   FontFamilyNotFound = 14
'   FontStyleNotFound = 15
'   NotTrueTypeFont = 16
'   UnsupportedGdiplusVersion = 17
'   GdiplusNotInitialized = 18
'   PropertyNotFound = 19
'   PropertyNotSupported = 20
'End Enum
  
  Select Case lError
    Case GenericError:              S = "Generic Error"
    Case InvalidParameter:          S = "Invalid Parameter"
    Case OutOfMemory:               S = "Out Of Memory"
    Case ObjectBusy:                S = "Object Busy"
    Case InsufficientBuffer:        S = "Insufficient Buffer"
    Case NotImplemented:            S = "Not Implemented"
    Case Win32Error:                S = "Win32 Error"
    Case WrongState:                S = "Wrong State"
    Case Aborted:                   S = "Aborted"
    Case FileNotFound:              S = "File Not Found"
    Case ValueOverflow:             S = "Value Overflow"
    Case AccessDenied:              S = "Access Denied"
    Case UnknownImageFormat:        S = "Unknown Image Format"
    Case FontFamilyNotFound:        S = "FontFamily Not Found"
    Case FontStyleNotFound:         S = "FontStyle Not Found"
    Case NotTrueTypeFont:           S = "Not TrueType Font"
'    Case UnsupportedGdiplusVersion: S = "Unsupported Gdiplus Version"
'    Case GdiplusNotInitialized:     S = "Gdiplus Not Initialized"
    Case PropertyNotFound:          S = "Property Not Found"
    Case PropertyNotSupported:      S = "Property Not Supported"
    Case Else:                      S = "Unknown GDI+ Error"
  End Select
  
  GdiErrorString = S
End Function


Public Function Picture() As StdPicture
  Dim retval        As Long
  Dim uPictDesc     As PICTDESC
  Dim aGuid(0 To 3) As Long
  Dim hImg As Long
  
    If m_Image = 0 Then Exit Function   'no image was loaded
    
    '-- Create bitmap
    If m_Bitmap = 0 Then
        retval = GdipCreateHBITMAPFromBitmap(m_Image, hImg, vbBlack)
    Else
        retval = GdipCreateHBITMAPFromBitmap(m_Bitmap, hImg, vbBlack)
    End If
    
    If retval = 0 Then
        '-- Fill struct
        With uPictDesc
            .Size = Len(uPictDesc)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hImg
            .hPal = 0
        End With
        
        '-- Fill in the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        
        '-- Create picture from bitmap handle
        OleCreatePictureIndirect uPictDesc, aGuid(0), -1, Picture
        
        'not sure if this is needed but putting it in incase
        'GdipDisposeImage hImg
    End If
End Function


'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Function ConvertBMPtoPicture(ByVal hBMP As Long, Optional ByVal hPal As Long) As Picture
    Dim r As Long
    Dim Pic As PICTDESC
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic)            ' Length of structure
        .Type = vbPicTypeBitmap     ' Type of Picture (bitmap)
        .hBmpOrIcon = hBMP          ' Handle to bitmap
        .hPal = hPal                ' Handle to palette (may be null)
    End With

    'Create the picture
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set ConvertBMPtoPicture = IPic
End Function

Private Function ReportError(ByVal lError As Long, Optional sTitle As String, Optional sText As String)
    'uncomment msgbox or do whatever with the error message
    Dim m As String
    Dim sT1 As String
    
    sT1 = "Error"
    If Len(sTitle) > 0 Then
        sT1 = sT1 & " : " & sTitle
    End If
    
    m = m & "Error code " & lError & vbCrLf
    m = m & GdiErrorString(lError) & vbCrLf
    
    If Len(sText) > 0 Then
        m = m & sText & vbCrLf
    End If
    
    'MsgBox M, vbCritical, sT1
End Function

Private Function GetEncoderClsid(sMimeType As String) As CLSID
   Dim lNumCoders       As Long
   Dim lSize            As Long
   Dim uInfo()          As ImageCodecInfo
   Dim lIdx             As Long
   Dim strEncoder       As String

   GdipGetImageEncodersSize lNumCoders, lSize
   If lSize > 0 Then
      ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As ImageCodecInfo
      GdipGetImageEncoders lNumCoders, lSize, uInfo(0)
      For lIdx = 0 To lNumCoders - 1
        strEncoder = PtrToStrW(uInfo(lIdx).MimeTypePtr)
         If StrComp(strEncoder, sMimeType, vbTextCompare) = 0 Then
            GetEncoderClsid = uInfo(lIdx).CLSID
            Exit For
         End If
      Next
   End If
End Function

'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
Private Function PtrToStrW(ByVal lpsz As Long) As String
   Dim sOut             As String
   Dim lLen             As Long

   lLen = lstrlenW(lpsz)

   If (lLen > 0) Then
      'was sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
      sOut = String$(lLen * 2, vbNullChar)
      CopyMemory ByVal sOut, ByVal lpsz, lLen * 2
      'PtrToStrW = StrConv(sOut, vbFromUnicode)
      PtrToStrW = TrimNulls(sOut)
   End If
End Function

Private Function TrimNulls(ByRef sText As String) As String
    Dim lPos As Long
   ' Locate the first instance of a Chr$(0) (NULL) character, and trim everything after it.
   lPos = InStr(1, sText, Chr$(0))
   If lPos <> 0 Then
      sText = Mid$(sText, 1, lPos - 1)
   End If
   TrimNulls = sText
End Function

Private Function SplitStrFromRight(ByVal searchStr, ByVal lookFor, Optional leftStr, Optional rightStr)
    'searches from right to left
    'for a single char within the searchString
    'when it finds it it returns the left and right sides
    Dim tempChar As String
    
    If Len(searchStr) = 0 Then Exit Function
    If IsMissing(rightStr) Then rightStr = ""
        
    tempChar = Right$(searchStr, 1)
    leftStr = Left(searchStr, Len(searchStr) - 1)
    If LCase(tempChar) <> LCase(lookFor) Then
        rightStr = tempChar & rightStr
        SplitStrFromRight leftStr, lookFor, leftStr, rightStr
    Else
        'we made a match, return values
    End If

End Function

Private Function pathExists(ByVal sPath As String, Optional bMakeIt As Boolean = False) As Boolean
    'checks to see if a path exists
    Dim retval As Long
    Dim Security As SECURITY_ATTRIBUTES
    
    retval = GetFileAttributes(sPath)
    If retval <> -1 Then
        pathExists = True
    Else
        If bMakeIt Then
            retval = CreateDirectory(sPath, Security)
            If retval <> 0 Then
                pathExists = True
            End If
        End If
    End If
End Function

Private Function EnsureBackslash(ByRef StrIn As String)
    'makes sure the last char is a backslash
    If Not Right(StrIn, 1) = "\" Then
        StrIn = StrIn & "\"
    End If
    EnsureBackslash = StrIn
End Function

Private Function Dispose()
    'clears any left overs
    DisposeGraphic
    DisposeImage
    
End Function
Private Function DisposeGraphic()
    GdipDeleteGraphics m_Graphic
    m_Graphic = 0
End Function
Private Function DisposeImage()
    GdipDisposeImage m_Image
    m_Image = 0
End Function


'---------------
'CLASS FUNCTIONS
'---------------
Private Sub Class_Initialize()
    'start the GDI engine
    On Error GoTo Handler
    
    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    GdiplusStartup m_hGDIplus, GpInput
    
    Exit Sub
Handler:
    ReportError Err.Number, "Initialising GDI+", Err.Description
    Resume Next

End Sub

Private Sub Class_Terminate()
    If Initialised Then
        DisposeImage
        Call GdiplusShutdown(m_hGDIplus)
    End If
End Sub

Public Function Rotate(ByVal angle As Rot_Angle) As Boolean
    Dim sngScale As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim lPixelFormat As Long
    Dim scaleMode As Long
    
'--clear up
    DisposeGraphic
    If m_Image = 0 Then Exit Function

'--defaults
    scaleMode = InterpolationModeHighQualityBilinear
    
'--Establish resize measurements
    'find out what we have now
    GdipGetImageDimension m_Image, sngWidth, sngHeight
    'bail if its a bad image
    If (sngWidth <= 0) Or (sngHeight <= 0) Then Exit Function
    'get the existing scale
    GdipGetImagePixelFormat m_Image, lPixelFormat
    If angle <> Deg180 Then
    GdipCreateBitmapFromScan0 sngHeight, sngWidth, 0, lPixelFormat, ByVal 0&, m_Bitmap
    Else
    GdipCreateBitmapFromScan0 sngWidth, sngHeight, 0, lPixelFormat, ByVal 0&, m_Bitmap
   End If
    'get a handle to the graphics object of our new bitmap
    GdipGetImageGraphicsContext m_Bitmap, m_Graphic
    'could use something like this if you want it to be displayed on screen
    'GdipCreateFromHDC formGDItest.Picture1.hDC, m_Graphic
   
    'set the method that we want to use to do this transformation
    GdipGetInterpolationMode m_Graphic, scaleMode
    Select Case angle
    Case Deg90:
        GdipRotateWorldTransform m_Graphic, 90, MatrixOrderAppend
        'GdipDrawImageRect m_Graphic, m_Image, 0, -sngHeight + 0.8, sngWidth, sngHeight
        GdipDrawImageRect m_Graphic, m_Image, 0, -sngHeight, sngWidth, sngHeight
    Case Deg180:
        GdipRotateWorldTransform m_Graphic, 180, MatrixOrderAppend
        GdipDrawImageRect m_Graphic, m_Image, -sngWidth, -sngHeight, sngWidth, sngHeight
    Case Deg270:
        GdipRotateWorldTransform m_Graphic, 270, MatrixOrderAppend
        'GdipDrawImageRect m_Graphic, m_Image, -sngWidth + 0.8, 0, sngWidth, sngHeight
         GdipDrawImageRect m_Graphic, m_Image, -sngWidth, 0, sngWidth, sngHeight
    End Select
                    
    Rotate = True
End Function

Private Function WriteFileToDisk(sFileName As String, uEncParams As EncoderParameters) As Long
    'does the actual saving of the image
    Dim sEncoder As String
    sEncoder = "image/jpeg"
  
    WriteFileToDisk = GdipSaveImageToFile(m_Image, sFileName, GetEncoderClsid(sEncoder), uEncParams)

End Function



