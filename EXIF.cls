VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEXIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'What does this do ?
'-------------------
'used for the sole purpose of accessing EXIF information
'each property item is loaded as it is needed.
'ie we dont do a mass load of all property items like we do in the FILE class
'this is because GDI+ places a lock on the file when it is opened

'How to use :
'------------
'.load a file using its full path name
'add / edit its meta data
'if you have made any changes then these can be saved back losslessly

'Example useage
'--------------
'dim x as new duncan_metadata.exif
'x.Load ("c:\myfile.jpg")
'x.EXIFdescription = "Hello baby its me"
'x.Save
'set x = nothing

'Requirements
'needs GDI+ TYPE LIBRARY from Dana Seaman located at http://www.cyberactivex.com

'Public Const PropertyTagResolutionXUnit = &H5001
'Public Const PropertyTagResolutionYUnit = &H5002
'Public Const PropertyTagExifExposureProg = &H8822
'Public Const PropertyTagExifShutterSpeed = &H9201
'Public Const PropertyTagExifAperture = &H9202
'Public Const PropertyTagExifBrightness = &H9203
'Public Const PropertyTagExifExposureBias = &H9204
'Public Const PropertyTagExifMaxAperture = &H9205

Option Explicit
'-----------------
'MY VARIABLES
'-----------------
Private m_hGDIplus As Long      'token to GDI engine
Private m_Image As Long         'handle to image
Private m_FileName As String    'remembers the file name we have open
'-----------------
'API FUNCTIONS
'-----------------
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'-----------------
'GDI+ FUNCTIONS
'-----------------
Private Declare Function LocalGdipGetPropertyItem Lib "gdiplus.dll" Alias "GdipGetPropertyItem" _
   (ByVal img As Long, _
   ByVal lId As Long, _
   ByVal lSize As Long, _
   ByVal lPtrBuff As Long) As Long
  Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'-----------------
'PUBLIC FUNCTIONS
'-----------------
Public Function Load(sFileName As String) As Boolean
    Dim retval As Long
    DisposeImage
    m_FileName = sFileName
    retval = GdipLoadImageFromFile(sFileName, m_Image)
    If retval = 0 Then
        Load = True
    Else
        ReportError retval, "Loading Image", "Image not loaded"
    End If
End Function

Public Function PaintDC(ByVal hDC As Long, x As Integer, y As Integer)
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim graphic As Long
    
    'get a link to the DC we are going to paint onto
    GdipCreateFromHDC hDC, graphic
    
    'draw onto it
    
        GdipGetImageDimension m_Image, sngWidth, sngHeight
        GdipDrawImageRect graphic, m_Image, x, y, sngWidth, sngHeight
    
    
    GdipDeleteGraphics graphic
End Function

Public Property Get PropertyCount() As Long
    Dim numProperties As Long
    If m_Image <> 0 Then
        GdipGetPropertyCount m_Image, numProperties
    End If
    PropertyCount = numProperties
End Property

Public Function Save() As Boolean
    'saves the file back to where it was loaded from
    'need to do a 90 then 270 transform so that we dont recompress the JPG
    Dim retval        As Long
    Dim encoderCLSID  As CLSID
    Dim uEncParams As EncoderParameters

    Const EncoderValueTransformRotate90 As Long = 13
    Const EncoderValueTransformRotate270 As Long = 15
  
    Dim OriginalFilePath As String
    Dim TempFileOriginal As String
    Dim TempFileRotated90 As String
  
    If Not Initialised Then Exit Function
    
    'get the values we will need for file swapping
    OriginalFilePath = m_FileName
    TempFileOriginal = GetATemporaryFileName
    TempFileRotated90 = GetATemporaryFileName
    
    'prepare the encoder for a rotation of 90 degrees
    uEncParams.Count = 1
    With uEncParams.Parameter
        .GUID = CLSIDFromString(EncoderTransformation)
        .NumberOfValues = 1
        .Type = EncoderParameterValueTypeLong
        .ValuePtr = VarPtr(EncoderValueTransformRotate90)
    End With
    
    'save the image to a temp location
    retval = WriteFileToDisk(TempFileRotated90, uEncParams)
    'if save went ok then proceed with file swap
    If retval = 0 Then
'release the image we have locked
'necessary?
DisposeImage
    'rename the original file to some temp name so we have a backup
        retval = CopyFile(OriginalFilePath, TempFileOriginal, 0)
        If retval > 0 Then  'backup succeeded
    'remove the original
             'retval = DeleteFile(OriginalFilePath)
    'now load the temp file we just saved
            If Load(TempFileRotated90) Then
    'rotate this back 270 degrees
                With uEncParams.Parameter
                    .GUID = CLSIDFromString(EncoderTransformation)
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong
                    .ValuePtr = VarPtr(EncoderValueTransformRotate270)
                End With
    'and save it
                retval = WriteFileToDisk(OriginalFilePath, uEncParams)
                If retval = 0 Then
                    'saving of the new image went successfully
                    Save = True
                Else
                    'we couldnt complete the transformation
                    'restore the backup
                    ReportError retval, "Saving", "Writing transform 2."
                    retval = CopyFile(TempFileOriginal, OriginalFilePath, 0)
                End If
    'load the original again so the new data is in memory
                Load OriginalFilePath
            Else
                ReportError retval, "Saving", "Loading transform 1."
            End If
        Else
            ReportError retval, "Saving", "Backing up original file."
        End If
    Else
        'save failed, exit out
        ReportError retval, "Saving", "Writing transform 1."
    End If
    
    'remove the tempory files
    retval = DeleteFile(TempFileOriginal)
    retval = DeleteFile(TempFileRotated90)
    
End Function


'-----------------
'EXIF PROPERTIES
'-----------------
'since there are only a handfull of EXIF tags I want to edit I have written these wrappers
'to get and then set the values back
'its basically just a fast and easy way for me to access the info I want

'PropertyTagImageTitle = &H320 = 800
'PropertyTagImageDescription = &H10E = 270
'PropertyTagEquipMake = &H10F = 271
'PropertyTagEquipModel = &H110 = 272
'PropertyTagArtist = &H13B = 315
'PropertyTagCopyright = &H8298 =  33432
'PropertyTagExifUserComment = &H9286 = 37510
'PropertyTagExifDTOrig = &H9003 = 36867
'PropertyTagExifDTDigitized = &H9004 = 36868
'PropertyTagDateTime = &H132 = 306

Public Property Get EXIFtitle() As String
    Const id = 800      'PropertyTagImageTitle = &H320 = 800
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFtitle = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFtitle(inVal As String)
    Const id = 800      'PropertyTagImageTitle = &H320 = 800
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFdescription() As String
    Const id = 270      'PropertyTagImageDescription = &H10E = 270
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFdescription = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFdescription(inVal As String)
    Const id = 270      'PropertyTagImageDescription = &H10E = 270
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFartist() As String
    Const id = 315      'PropertyTagArtist = &H13B = 315
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFartist = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFartist(inVal As String)
    Const id = 315      'PropertyTagArtist = &H13B = 315
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFcomment() As String
    Const id = 37510      'PropertyTagExifUserComment = &H9286 = 37510
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFcomment = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFcomment(inVal As String)
    Const id = 37510      'PropertyTagExifUserComment = &H9286 = 37510
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
      
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFcopyright() As String
    Const id = 33432      'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFcopyright = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFcopyright(inVal As String)
    Const id = 33432      'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFmodified() As Date
    Const id = 306      'PropertyTagDateTime = &H132 = 306
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFmodified = ISODateTimeToDate(pi.ParseString)
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFmodified(inVal As Date)
    Const id = 306      'PropertyTagDateTime = &H132 = 306
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(DateToISODateTime(inVal), id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        'should be a date value
        'convert it to
        SetPropertyItem pi
    End If
    Set pi = Nothing
End Property


'Public Const PropertyTagExifExposureTime = &H829A
'Public Const PropertyTagExifFNumber = &H829D
Public Property Get ExifExposureTime() As String
    Const id = 33434 '&H829D       'PropertyTagSoftwareUsed= &h131 =
    Dim j As Long
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
      ExifExposureTime = sItem 'Mid(sItem, 1, 1) & "." & Mid(sItem, 2)  '''test only for f1 to f9  not f11 and above
    End If
    Set pi = Nothing
End Property
Public Property Let ExifExposureTime(inVal As String)
    Const id = 33434    'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        'EXIFsoftware = inVal
    End If
    Set pi = Nothing
End Property

Public Property Get ExifShutterSpeed() As String
    Const id = 37377 '&H829D       'PropertyTagSoftwareUsed= &h131 =
    Dim j As Long
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         '  Debug.Print "exp", j, pi.ParseString, pi.ParseLong(j)
         Next j
        'retreival was successful
      ExifShutterSpeed = "1/" & sItem & " s" 'Mid(sItem, 1, 1) & "." & Mid(sItem, 2)  '''test only for f1 to f9  not f11 and above
    End If
    Set pi = Nothing
End Property
Public Property Let ExifShutterSpeed(inVal As String)
    Const id = 37377    'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        'EXIFsoftware = inVal
    End If
    Set pi = Nothing
End Property
'ExifExposureProg = &H8822
Public Property Get ExifExposureProg() As String
    Const id = 34850 '&H829D       'PropertyTagSoftwareUsed= &h131 =
    Dim j As Long
    Dim sItem As String
    Dim sDescription As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
       For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
         Select Case sItem
         Case 0
            sDescription = "Normal Mode"
         Case 1
            sDescription = "Manual Mode"
        Case 2
            sDescription = "Auto Mode"
         Case Else
            sDescription = sItem
         End Select
      ExifExposureProg = sDescription 'Mid(sItem, 1, 1) & "." & Mid(sItem, 2)  '''test only for f1 to f9  not f11 and above
    End If
    Set pi = Nothing
End Property
Public Property Let ExifExposureProg(inVal As String)
    Const id = 34850  'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        'EXIFsoftware = inVal
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFFNumber() As String
    Const id = 33437 '&H829D       'PropertyTagSoftwareUsed= &h131 =
    Dim j As Long
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
           ' Debug.Print "exif", j, pi.ParseString, pi.ItemType
         Next j
        'retreival was successful
       EXIFFNumber = Mid(sItem, 1, 1) & "." & Mid(sItem, 2)  '''test only for f1 to f9  not f11 and above
    End If
    Set pi = Nothing
End Property
Public Property Let EXIFFNumber(inVal As String)
    Const id = 33437    'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        'EXIFsoftware = inVal
    End If
    Set pi = Nothing
End Property


Public Property Get EXIFsoftware() As String
    Const id = 305        'PropertyTagSoftwareUsed= &h131 =
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFsoftware = pi.ParseString
    End If
    Set pi = Nothing
End Property
Public Property Let EXIFsoftware(inVal As String)
    Const id = 305      'PropertyTagCopyright = &H8298 = 33432
    Dim pi As New GDIPPropertyItem
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        'EXIFsoftware = inVal
    End If
    Set pi = Nothing
End Property
'read only from here down - these shouldnt need changing
'Public Const PropertyTagResolutionXUnit = &H5001
'Public Const PropertyTagResolutionYUnit = &H5002
'Public Const PropertyTagExifExposureProg = &H8822
'Public Const PropertyTagExifShutterSpeed = &H9201
'Public Const PropertyTagExifAperture = &H9202
'Public Const PropertyTagExifBrightness = &H9203
'Public Const PropertyTagExifExposureBias = &H9204
'Public Const PropertyTagExifMaxAperture = &H9205
'Public Const PropertyTagExifExposureProg = &H8822
Public Property Get EXIFmake() As String
    Const id = 271      'PropertyTagEquipMake = &H10F = 271
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFmake = pi.ParseString
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFmodel() As String
    Const id = 272      'PropertyTagEquipModel = &H110 = 272
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFmodel = pi.ParseString
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFtaken() As String
    Const id = 36867      'PropertyTagExifDTOrig = &H9003 = 36867
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFtaken = ISODateTimeToDate(pi.ParseString)
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFflash() As String
    Const id = 37385      'PropertyTagExifFlash = &H9209 = 37385
    Dim j As Long
    Dim sItem As String
    Dim sDescription As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
         Select Case sItem
         Case 0
            sDescription = "Not Fired"
         Case 1
            sDescription = "Fired."
         Case 5
            sDescription = "Strobe return not detected"
         Case 7
            sDescription = "Strobe return detected"
         Case 9
            sDescription = "Flash fired; Compulsory flash mode"
         Case 13
            sDescription = "Flash fired; Compulsory flash mode; Return light not detected"
         Case 15
            sDescription = "Flash fired; Compulsory flash mode; Return light detected"
         Case 16
            sDescription = "Flash not fired; Compulsory flash mode"
         Case 24
            sDescription = "Flash not fired; Auto mode"
         Case 25
            sDescription = "Flash fired; Auto mode"
         Case 29
            sDescription = "Flash fired; Auto mode; Return light not detected"
         Case 31
            sDescription = "Flash fired; Auto mode; Return light detected"
         Case 32
            sDescription = "No flash function"
         Case 65
            sDescription = "Flash fired; Red-eye reduction mode"
         Case 69
            sDescription = "Flash fired; Red-eye reduction mode; Return light not detected"
         Case 71
            sDescription = "Flash fired; Red-eye reduction mode; Return light detected"
         Case 73
            sDescription = "Flash fired; Compulsory flash mode; Red-eye reduction mode"
         Case 77
            sDescription = "Flash fired; Compulsory flash mode; Red-eye reduction mode; Return light not detected"
         Case 79
            sDescription = "Flash fired; Compulsory flash mode; Red-eye reduction mode; Return light detected"
         Case 89
            sDescription = "Flash fired; Auto mode; Red-eye reduction mode"
         Case 93
            sDescription = "Flash fired; Auto mode; Red-eye reduction mode; Return light not detected"
         Case 95
            sDescription = "Flash fired; Auto mode; Red-eye reduction mode; Return light detected"
         Case Else
            sDescription = sItem
         End Select
     
         EXIFflash = sDescription
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFiso() As String
    Const id = 34855      'PropertyTagExifISOSpeed  = &H8827 = 34855
    Dim j As Long
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
       EXIFiso = sItem
    End If
    Set pi = Nothing
' iso-speed                 8827  ISO Speed
' Values: An integer in the range 0 to 65535
    
End Property

Public Property Get EXIFprogram() As String
    Const id = 34850      'PropertyTagExifExposureProg = &H8822 = 34850
    Dim j As Long
    Dim sItem As String
    Dim sDescription As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
         Select Case sItem
         Case 1
            sDescription = "Manual"
         Case 2
            sDescription = "Program Normal"
         Case 3
            sDescription = "Aperture Priority"
         Case 4
            sDescription = "Shutter Priority"
         Case 5
            sDescription = "Program Creative"
         Case 6
            sDescription = "Program Action"
         Case 7
            sDescription = "Portrait Mode"
         Case 8
            sDescription = "Landscape Mode"
         End Select

         EXIFprogram = sDescription
    End If
    Set pi = Nothing
    
End Property
Public Property Get EXIFscene() As String
    Const id = 41990      '&HA406
    Dim j As Long
    Dim sItem As String
    Dim sDescription As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
         Select Case sItem
         Case 0
            sDescription = "Standard"
         Case 1
            sDescription = "Landscape"
         Case 2
            sDescription = "Portrait"
         Case 3
            sDescription = "Night Scene"
         End Select

         EXIFscene = sDescription
    End If
    Set pi = Nothing
    
End Property

Public Property Get EXIFlightsource() As String
    Const id = 37384      'PropertyTagExifLightSource = &H9208
    Dim j As Long
    Dim sItem As String
    Dim sDescription As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
         Select Case sItem
         Case 0
            sDescription = "Unknown"
         Case 1
            sDescription = "Daylight"
         Case 2
            sDescription = "Fluorescent"
         Case 3
            sDescription = "Tungsten"
         Case 4
            sDescription = "Flash"
         Case 9
            sDescription = "Fine Weather"
         Case 10
            sDescription = "Cloudy Weather"
         Case 11
            sDescription = "Shade"
         Case 12
            sDescription = "Daylight Fluorescent"
         Case 13
            sDescription = "Day White Fluorescent"
         Case 14
            sDescription = "Cool White Fluorescent"
         Case 15
            sDescription = "White Fluorescent"
         Case 17
            sDescription = "Standard Light A"
         Case 18
            sDescription = "Standard Light B"
         Case 19
            sDescription = "Standard Light C"
         Case 20
            sDescription = "D55"
         Case 21
            sDescription = "D65"
         Case 22
            sDescription = "D75"
         Case 23
            sDescription = "D50"
         Case 24
            sDescription = "ISO Studio Tungsten "
         Case 255
            sDescription = "Other Light Source"
         End Select

         EXIFlightsource = sDescription
    End If
    Set pi = Nothing
    
End Property

Public Property Get EXIFxpTitle() As String
    Const id = 40091      '&H9C9B
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFxpTitle = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFxpTitle(inVal As String)
    Const id = 40091      '&H9C9B
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFxpComment() As String
    Const id = 40092      '&H9C9C
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFxpComment = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFxpComment(inVal As String)
    Const id = 40092      '&H9C9C
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFxpAuthor() As String
    Const id = 40093      'Author = &H9C9D
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFxpAuthor = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFxpAuthor(inVal As String)
    Const id = 40093      'Author = &H9C9D
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFxpKeywords() As String
    Const id = 40094      '&H9C9E
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFxpKeywords = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFxpKeywords(inVal As String)
    Const id = 40094      '&H9C9E
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property
Public Property Get EXIFxpSubject() As String
    Const id = 40095      '&H9C9F
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFxpSubject = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFxpSubject(inVal As String)
    Const id = 40095      '&H9C9F
    Dim pi As New GDIPPropertyItem
    'limited to 999 bytes
    If Len(inVal) > 999 Then inVal = Left(inVal, 999)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFdistance() As String
'not yet verified - no camera i have access to shows distance

    Const id = 37382      'PropertyTagExifSubjectDist = &H9206
    Dim j As Long
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
         For j = 1 To pi.ValueCount
            If (j > 1) Then
               sItem = sItem & ", "
            End If
            sItem = sItem & pi.ParseShort(j)
         Next j
        'retreival was successful
       EXIFdistance = sItem
    Else
        EXIFdistance = "Unknown"
    End If
    Set pi = Nothing
    
End Property

Public Property Get Height() As Long
    Dim retval As Long

    retval = GdipGetImageHeight(m_Image, Height)
    If retval Then
        ReportError retval, "Height evaluation", "Unable to determine image height"
    End If
   
End Property

Public Property Get Width() As Long
    Dim retval As Long

    retval = GdipGetImageWidth(m_Image, Width)
    If retval Then
        ReportError retval, "Width evaluation", "Unable to determine image width"
    End If
End Property

Public Function PaletteSize() As Long
    Dim lngSize As Long
    Dim retval As Long
    retval = GdipGetImagePaletteSize(m_Image, lngSize)
    PaletteSize = lngSize
End Function

Public Property Get PixelFormat() As String
    'returns colour depth
    Dim retval As Long
    Dim pf As Long 'PixelFormat
    Dim sDescription As String
    retval = GdipGetImagePixelFormat(m_Image, pf)
    If retval Then
        ReportError retval, "Pixel format evaluation", "Unable to determine image colour depth"
    Else
        Select Case pf
        Case PixelFormat1bppIndexed
            sDescription = "1bppIndexed"
        Case PixelFormat4bppIndexed
            sDescription = "4bppIndexed"
        Case PixelFormat8bppIndexed
            sDescription = "8bppIndexed"
        Case PixelFormat16bppARGB1555
            sDescription = "16bppARGB1555"
        Case PixelFormat16bppGrayScale
            sDescription = "16bppGrayScale"
        Case PixelFormat16bppRGB555
            sDescription = "16bppRGB555"
        Case PixelFormat16bppRGB565
            sDescription = "16bppRGB565"
        Case PixelFormat24bppRGB
            sDescription = "24bppRGB"
        Case PixelFormat32bppARGB
            sDescription = "32bppARGB"
        Case PixelFormat32bppPARGB
            sDescription = "32bppPARGB"
        Case PixelFormat32bppRGB
            sDescription = "32bppRGB"
        Case PixelFormat48bppRGB
            sDescription = "48bppRGB"
        Case PixelFormat64bppARGB
            sDescription = "64bppARGB"
        Case PixelFormat64bppPARGB
            sDescription = "64bppPARGB"
        Case Else
            sDescription = "Unknown"
        End Select
        
    End If
    PixelFormat = sDescription
End Property

Public Property Get EXIFversion() As String
    Const id = 36864      '&H9000
    Dim sItem As String
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        sItem = pi.ParseString
        If IsNumeric(sItem) Then
            sItem = sItem / 100
        End If
        EXIFversion = sItem
    End If
    Set pi = Nothing
End Property

Public Property Get EXIFid() As String
    Const id = 42016      'Unique Image ID = &HA420
    Dim pi As New GDIPPropertyItem
    'get the property from the image
    Set pi = PropertyItemForID(id)
    'is it a valid object ?
    If pi.id = id Then
        'retreival was successful
        EXIFid = pi.ParseString
    End If
    Set pi = Nothing
End Property

Public Property Let EXIFid(inVal As String)
    Const id = 42016      'Unique Image ID = &HA420
    Dim pi As New GDIPPropertyItem
    'limited to 32 bytes
    If Len(inVal) > 32 Then inVal = Left(inVal, 32)
    'create a properly formatted PropertyItem for adding
    If pi.SetPropertyStringValue(inVal, id) Then
        'use GDI+ to add it to the image (in memory - still will need physical save to disk)
        SetPropertyItem pi
        'change updated value
        EXIFmodified = Now
    End If
    Set pi = Nothing
End Property


Public Property Get FileType() As String
    Const ImageFormatSuffix        As String = "-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatUndefined     As String = "{B96B3CA9" & ImageFormatSuffix
    Const ImageFormatMemoryBMP     As String = "{B96B3CAA" & ImageFormatSuffix
    Const ImageFormatBMP           As String = "{B96B3CAB" & ImageFormatSuffix
    Const ImageFormatEMF           As String = "{B96B3CAC" & ImageFormatSuffix
    Const ImageFormatWMF           As String = "{B96B3CAD" & ImageFormatSuffix
    Const ImageFormatJPEG          As String = "{B96B3CAE" & ImageFormatSuffix
    Const ImageFormatPNG           As String = "{B96B3CAF" & ImageFormatSuffix
    Const ImageFormatGIF           As String = "{B96B3CB0" & ImageFormatSuffix
    Const ImageFormatTIFF          As String = "{B96B3CB1" & ImageFormatSuffix
    Const ImageFormatEXIF          As String = "{B96B3CB2" & ImageFormatSuffix
    Const ImageFormatIcon          As String = "{B96B3CB5" & ImageFormatSuffix
    
    Dim retval As Long
    Dim FormatID As CLSID
    
    retval = GdipGetImageRawFormat(m_Image, FormatID)
    
    Select Case GetGuidString(FormatID)
           Case ImageFormatUndefined
             FileType = "Undefined"
           Case ImageFormatMemoryBMP
             FileType = "Memory BMP"
           Case ImageFormatBMP
             FileType = "BMP"
           Case ImageFormatEMF
             FileType = "EMF"
           Case ImageFormatWMF
             FileType = "WMF"
           Case ImageFormatJPEG
             FileType = "JPEG"
           Case ImageFormatPNG
             FileType = "PNG"
           Case ImageFormatGIF
             FileType = "GIF"
           Case ImageFormatTIFF
             FileType = "TIFF"
           Case ImageFormatEXIF
             FileType = "EXIF"
           Case ImageFormatIcon
             FileType = "Icon"
    End Select

End Property


'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Function DisposeImage()
    Dim retval As Long
    retval = GdipDisposeImage(m_Image)
    If retval = 0 Then
        m_Image = 0
    Else
        'cant release image
    End If
End Function

Private Function PropertyItemForID(ByVal lId As Long) As GDIPPropertyItem
    'gets a property from the image and stores it in the property item structure
    Dim lSize As Long
    If m_Image <> 0 Then
        GdipGetPropertyItemSize m_Image, lId, lSize
        If (lSize > 0) Then
           ReDim b(0 To lSize - 1) As Byte
           Dim lPtrBuff As Long
           lPtrBuff = VarPtr(b(0))
           LocalGdipGetPropertyItem m_Image, lId, lSize, lPtrBuff
           Dim p As PropertyItem
           Dim cItem As New GDIPPropertyItem
           Dim lDataSize As Long
           If Not (lPtrBuff = 0) And (lSize > 0) Then
              CopyMemory p, ByVal lPtrBuff, Len(p)
              cItem.fInit p.id, p.Length, p.Type, p.ValuePtr, lSize
           End If
           Set PropertyItemForID = cItem
        End If
    End If
End Function

Private Sub SetPropertyItem(item As GDIPPropertyItem)
    Dim retval As Long
    Dim p As GdiPlus.PropertyItem
    'make sure we have an image loaded
    If m_Image <> 0 Then
        'and a valid property item assigned
        If item.id <> 0 Then
            p.id = item.id
            p.Length = item.Length
            p.Type = item.ItemType
            ReDim b(0 To item.DataBufferSize - 1) As Byte
            item.GetData b()
            p.ValuePtr = VarPtr(b(0))
            retval = GdipSetPropertyItem(m_Image, p)
        End If
    End If
End Sub




'------------------------
'PRIVATE HELPER FUNCTIONS
'------------------------
Private Property Get Initialised() As Boolean
    If m_hGDIplus <> 0 Then
        Initialised = True
    End If
End Property
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
      sOut = String$(lLen * 2, vbNullChar)
      CopyMemory ByVal sOut, ByVal lpsz, lLen * 2
      PtrToStrW = Replace(sOut, Chr$(0), "", 1)
   End If
End Function
Private Function ISODateTimeToDate( _
      ByVal isoDateTime As String _
   ) As Date
   'to test
   'example = "2003:12:03 18:24:55"
On Error Resume Next
Dim dDate As Date
   dDate = DateSerial( _
         Mid(isoDateTime, 1, 4), _
         Mid(isoDateTime, 6, 2), _
         Mid(isoDateTime, 9, 2)) + _
      TimeSerial( _
         Mid(isoDateTime, 12, 2), _
         Mid(isoDateTime, 15, 2), _
         Mid(isoDateTime, 18, 2))
   ISODateTimeToDate = dDate
End Function

Private Function DateToISODateTime(ByVal dDateTime As Date) As String
    'converts a datetime into ISO date time format
    'example = "2003:12:03 18:24:55"
    Dim retval As String
    Dim dp As String
    
    If IsDate(dDateTime) Then
        retval = Year(dDateTime)
        retval = retval & ":"
        dp = Month(dDateTime)
        If Len(dp) = 1 Then dp = "0" & dp
        retval = retval & dp
        retval = retval & ":"
        dp = Day(dDateTime)
        If Len(dp) = 1 Then dp = "0" & dp
        retval = retval & dp
        retval = retval & " "
        dp = Hour(dDateTime)
        If Len(dp) = 1 Then dp = "0" & dp
        retval = retval & dp
        retval = retval & ":"
        dp = Minute(dDateTime)
        If Len(dp) = 1 Then dp = "0" & dp
        retval = retval & dp
        retval = retval & ":"
        dp = Second(dDateTime)
        If Len(dp) = 1 Then dp = "0" & dp
        retval = retval & dp
    End If
    DateToISODateTime = retval
End Function

Private Function GetATemporaryFileName() As String
    'used to create swap file for lossless saving
    On Error Resume Next
    Dim sTempDir As String
    Dim sTempFileName As String
    
    'Create buffers
    sTempDir = String(100, Chr$(0))
    sTempFileName = String(260, 0)
    'Get the temporary path
    GetTempPath 100, sTempDir
    'Strip the 0's off the end
    sTempDir = Left$(sTempDir, InStr(sTempDir, Chr$(0)) - 1)
    'backup in case none found
    If Len(sTempDir) = 0 Then
        sTempDir = "C:\"
    End If
    'get file name
    GetTempFileName sTempDir, "IMG", 0, sTempFileName
    'Strip the 0's off the end
    sTempFileName = Left$(sTempFileName, InStr(sTempFileName, Chr$(0)) - 1)
    GetATemporaryFileName = sTempFileName
End Function

Private Function ReportError(ByVal lError As Long, Optional sTitle As String, Optional sText As String)
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

Private Function GdiErrorString(ByVal lError As Long) As String
  Dim S As String
  
  Select Case lError
    Case 0: S = "No Error"
    Case 1: S = "Generic Error"
    Case 2: S = "Invalid Parameter"
    Case 3: S = "Out Of Memory"
    Case 4: S = "Object Busy"
    Case 5: S = "Insufficient Buffer"
    Case 6: S = "Not Implemented"
    Case 7: S = "Win32 Error"
    Case 8: S = "Wrong State"
    Case 9: S = "Aborted"
    Case 10: S = "File Not Found"
    Case 11: S = "Value Overflow"
    Case 12: S = "Access Denied"
    Case 13: S = "Unknown Image Format"
    Case 14: S = "FontFamily Not Found"
    Case 15: S = "FontStyle Not Found"
    Case 16: S = "Not TrueType Font"
    Case 17: S = "Unsupported Gdiplus Version"
    Case 18: S = "Gdiplus Not Initialized"
    Case 19: S = "Property Not Found"
    Case 20: S = "Property Not Supported"
    Case Else: S = "Unknown Error"
  End Select
  
  GdiErrorString = S
End Function


Private Function WriteFileToDisk(sFileName As String, uEncParams As EncoderParameters) As Long
    'does the actual saving of the image
    Dim sEncoder As String
    sEncoder = "image/jpeg"
  
    WriteFileToDisk = GdipSaveImageToFile(m_Image, sFileName, GetEncoderClsid(sEncoder), uEncParams)

End Function

Private Function hexPad(ByVal Value As Long, ByVal padSize As Long) As String
    'used below
    Dim sRet As String
    Dim lMissing As Long
    sRet = Hex$(Value)
    lMissing = padSize - Len(sRet)
    If (lMissing > 0) Then
        sRet = String$(lMissing, "0") & sRet
    ElseIf (lMissing < 0) Then
    End If
        sRet = Mid$(sRet, -lMissing + 1)
    hexPad = sRet
End Function

Private Function GetGuidString(GUID As CLSID) As String
    'used to help determine an Image Type
    Dim i As Long
    Dim sGuid As String

    sGuid = "{" & hexPad(GUID.Data1, 8) & "-" & hexPad(GUID.Data2, 4) & "-" & hexPad(GUID.Data3, 4) & "-"
    sGuid = sGuid & hexPad(GUID.Data4(0), 2) & hexPad(GUID.Data4(1), 2) & "-"
    For i = 2 To 7
        sGuid = sGuid & hexPad(GUID.Data4(i), 2)
    Next i
    sGuid = sGuid & "}"
    GetGuidString = sGuid
End Function

'-----------------
'CLASS FUNCTIONS
'-----------------
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

Public Function SaveAs(Degree As EncoderValue) As Boolean
    Dim retval        As Long
    Dim encoderCLSID  As CLSID
    Dim uEncParams As EncoderParameters
    Dim OriginalFilePath As String
    Dim TempFileRotated As String
  
    If Not Initialised Then Exit Function
    
    'get the values we will need for file swapping
    OriginalFilePath = m_FileName
    TempFileRotated = GetATemporaryFileName
    
    'prepare the encoder for a rotation of ? degrees
    uEncParams.Count = 1
    With uEncParams.Parameter
        .GUID = CLSIDFromString(EncoderTransformation)
        .NumberOfValues = 1
        .Type = EncoderParameterValueTypeLong
        .ValuePtr = VarPtr(Degree)
    End With
    
    'save the image to a temp location
    retval = WriteFileToDisk(TempFileRotated, uEncParams)
    'if save went ok then proceed with file swap
    If retval = 0 Then
'release the image we have locked
'necessary?
    DisposeImage
    End If
    'rename the original file to some temp name so we have a backup
    retval = CopyFile(TempFileRotated, OriginalFilePath, 0)
    'remove the tempory files
    retval = DeleteFile(TempFileRotated)
    
End Function
