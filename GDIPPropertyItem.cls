VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPPropertyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This file originally sourced from http://www.vbaccelerator.com
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Enum PropertyTagType
    PropertyTagTypeByte = 1
    PropertyTagTypeASCII = 2
    PropertyTagTypeShort = 3
    PropertyTagTypeLong = 4
    PropertyTagTypeRational = 5
    PropertyTagTypeUndefined = 7
    PropertyTagTypeSLONG = 9
    PropertyTagTypeSRational = 10
End Enum

Private m_lId As Long                   'the property id
Private m_lLength As Long               'the length of the data
Private m_itemType As PropertyTagType   'the type of the tag eg Ascii
Private m_bData() As Byte               'the data

'for reference
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

'-----------------
'PUBLIC PROPERTIES
'-----------------
Public Property Get id() As Long
   id = m_lId
End Property

Public Property Get ItemType() As PropertyTagType
   ItemType = m_itemType
End Property

Public Property Get Length() As Long
   Length = m_lLength
End Property

Public Property Get Name() As String
    Select Case m_lId
    Case PropertyTagExifIFD
        Name = "ExifIFD"
    Case PropertyTagsIFD
        Name = "IFD"
    Case PropertyTagNewSubfileType
        Name = "NewSubfileType"
    Case PropertyTagSubfileType
        Name = "TagSubFileType"
    Case PropertyTagImageWidth
        Name = "ImageWidth"
    Case PropertyTagImageHeight
        Name = "ImageHeight"
    Case PropertyTagBitsPerSample
        Name = "BitsPerSample"
    Case PropertyTagCompression
        Name = "Compression"
    Case PropertyTagPhotometricInterp
        Name = "PhotometricInterp"
    Case PropertyTagThreshHolding
        Name = "ThreshHolding"
    Case PropertyTagCellWidth
        Name = "CellWidth"
    Case PropertyTagCellHeight
        Name = "CellHeight"
    Case PropertyTagFillOrder
        Name = "FillOrder"
    Case PropertyTagDocumentName
        Name = "DocumentName"
    Case PropertyTagImageDescription
        Name = "ImageDescription"
    Case PropertyTagEquipMake
        Name = "EquipMake"
    Case PropertyTagEquipModel
        Name = "EquipModel"
    Case PropertyTagStripOffsets
        Name = "StripOffsets"
    Case PropertyTagOrientation
        Name = "Orientation"
    Case PropertyTagSamplesPerPixel
        Name = "SamplesPerPixel"
    Case PropertyTagRowsPerStrip
        Name = "RowsPerStrip"
    Case PropertyTagStripBytesCount
        Name = "StripBytesCount"
    Case PropertyTagMinSampleValue
        Name = "MinSampleValue"
    Case PropertyTagMaxSampleValue
        Name = "MaxSampleValue"
    Case PropertyTagXResolution
        Name = "XResolution"
    Case PropertyTagYResolution
        Name = "YResolution"
    Case PropertyTagPlanarConfig
        Name = "PlanarConfig"
    Case PropertyTagPageName
        Name = "PageName"
    Case PropertyTagXPosition
        Name = "XPosition"
    Case PropertyTagYPosition
        Name = "YPosition"
    Case PropertyTagFreeOffset
        Name = "FreeOffset"
    Case PropertyTagFreeByteCounts
        Name = "FreeByteCounts"
    Case PropertyTagGrayResponseUnit
        Name = "GrayResponseUnit"
    Case PropertyTagGrayResponseCurve
        Name = "GrayResponseCurve"
    Case PropertyTagT4Option
        Name = "T4Option"
    Case PropertyTagT6Option
        Name = "T6Option"
    Case PropertyTagResolutionUnit
        Name = "ResolutionUnit"
    Case PropertyTagPageNumber
        Name = "PageNumber"
    Case PropertyTagTransferFuncition
        Name = "TransferFuncition"
    Case PropertyTagSoftwareUsed
        Name = "SoftwareUsed"
    Case PropertyTagDateTime
        Name = "DateTime"
    Case PropertyTagArtist
        Name = "Artist"
    Case PropertyTagHostComputer
        Name = "HostComputer"
    Case PropertyTagPredictor
        Name = "Predictor"
    Case PropertyTagWhitePoint
        Name = "WhitePoint"
    Case PropertyTagPrimaryChromaticities
        Name = "PrimaryChromaticities"
    Case PropertyTagColorMap
        Name = "ColorMap"
    Case PropertyTagHalftoneHints
        Name = "HalftoneHints"
    Case PropertyTagTileWidth
        Name = "TileWidth"
    Case PropertyTagTileLength
        Name = "TileLength"
    Case PropertyTagTileOffset
        Name = "TileOffset"
    Case PropertyTagTileByteCounts
        Name = "TileByteCounts"
    Case PropertyTagInkSet
        Name = "InkSet"
    Case PropertyTagInkNames
        Name = "InkNames"
    Case PropertyTagNumberOfInks
        Name = "NumberOfInks"
    Case PropertyTagDotRange
        Name = "DotRange"
    Case PropertyTagTargetPrinter
        Name = "TargetPrinter"
    Case PropertyTagExtraSamples
        Name = "ExtraSamples"
    Case PropertyTagSampleFormat
        Name = "SampleFormat"
    Case PropertyTagSMinSampleValue
        Name = "SMinSampleValue"
    Case PropertyTagSMaxSampleValue
        Name = "SMaxSampleValue"
    Case PropertyTagTransferRange
        Name = "TransferRange"
    Case PropertyTagJPEroc
        Name = "JPEroc"
    Case PropertyTagJPEGInterFormat
        Name = "JPEGInterFormat"
    Case PropertyTagJPEGInterLength
        Name = "JPEGInterLength"
    Case PropertyTagJPEGRestartInterval
        Name = "JPEGRestartInterval"
    Case PropertyTagJPEGLosslessPredictors
        Name = "JPEGLosslessPredictors"
    Case PropertyTagJPEointTransforms
        Name = "JPEointTransforms"
    Case PropertyTagJPEGQTables
        Name = "JPEGQTables"
    Case PropertyTagJPEGDCTables
        Name = "JPEGDCTables"
    Case PropertyTagJPEGACTables
        Name = "JPEGACTables"
    Case PropertyTagYCbCrCoefficients
        Name = "YCbCrCoefficients"
    Case PropertyTagYCbCrSubsampling
        Name = "YCbCrSubsampling"
    Case PropertyTagYCbCrPositioning
        Name = "YCbCrPositioning"
    Case PropertyTagREFBlackWhite
        Name = "REFBlackWhite"
    Case PropertyTagICCProfile
        Name = "ICCProfile"
    Case PropertyTagGamma
        Name = "Gamma"
    Case PropertyTagICCProfileDescriptor
        Name = "ICCProfileDescriptor"
    Case PropertyTagSRGBRenderingIntent
        Name = "SRGBRenderingIntent"
    Case PropertyTagImageTitle
        Name = "ImageTitle"
    Case PropertyTagCopyright
        Name = "Copyright"
    Case PropertyTagResolutionXUnit
        Name = "ResolutionXUnit"
    Case PropertyTagResolutionYUnit
        Name = "ResolutionYUnit"
    Case PropertyTagResolutionXLengthUnit
        Name = "ResolutionXLengthUnit"
    Case PropertyTagResolutionYLengthUnit
        Name = "ResolutionYLengthUnit"
    Case PropertyTagPrintFlags
        Name = "PrintFlags"
    Case PropertyTagPrintFlagsVersion
        Name = "PrintFlagsVersion"
    Case PropertyTagPrintFlagsCrop
        Name = "PrintFlagsCrop"
    Case PropertyTagPrintFlagsBleedWidth
        Name = "PrintFlagsBleedWidth"
    Case PropertyTagPrintFlagsBleedWidthScale
        Name = "PrintFlagsBleedWidthScale"
    Case PropertyTagHalftoneLPI
        Name = "HalftoneLPI"
    Case PropertyTagHalftoneLPIUnit
        Name = "HalftoneLPIUnit"
    Case PropertyTagHalftoneDegree
        Name = "HalftoneDegree"
    Case PropertyTagHalftoneShape
        Name = "HalftoneShape"
    Case PropertyTagHalftoneMisc
        Name = "HalftoneMisc"
    Case PropertyTagHalftoneScreen
        Name = "HalftoneScreen"
    Case PropertyTagJPEGQuality
        Name = "JPEGQuality"
    Case PropertyTagGridSize
        Name = "GridSize"
    Case PropertyTagThumbnailFormat
        Name = "ThumbnailFormat"
    Case PropertyTagThumbnailWidth
        Name = "ThumbnailWidth"
    Case PropertyTagThumbnailHeight
        Name = "ThumbnailHeight"
    Case PropertyTagThumbnailColorDepth
        Name = "ThumbnailColorDepth"
    Case PropertyTagThumbnailPlanes
        Name = "ThumbnailPlanes"
    Case PropertyTagThumbnailRawBytes
        Name = "ThumbnailRawBytes"
    Case PropertyTagThumbnailSize
        Name = "ThumbnailSize"
    Case PropertyTagThumbnailCompressedSize
        Name = "ThumbnailCompressedSize"
    Case PropertyTagColorTransferFunction
        Name = "ColorTransferFunction"
    Case PropertyTagThumbnailData
        Name = "ThumbnailData"
    Case PropertyTagThumbnailImageWidth
        Name = "ThumbnailImageWidth"
    Case PropertyTagThumbnailImageHeight
        Name = "ThumbnailImageHeight"
    Case PropertyTagThumbnailBitsPerSample
        Name = "ThumbnailBitsPerSample"
    Case PropertyTagThumbnailCompression
        Name = "ThumbnailCompression"
    Case PropertyTagThumbnailPhotometricInterp
        Name = "ThumbnailPhotometricInterp"
    Case PropertyTagThumbnailImageDescription
        Name = "ThumbnailImageDescription"
    Case PropertyTagThumbnailEquipMake
        Name = "ThumbnailEquipMake"
    Case PropertyTagThumbnailEquipModel
        Name = "ThumbnailEquipModel"
    Case PropertyTagThumbnailStripOffsets
        Name = "ThumbnailStripOffsets"
    Case PropertyTagThumbnailOrientation
        Name = "ThumbnailOrientation"
    Case PropertyTagThumbnailSamplesPerPixel
        Name = "ThumbnailSamplesPerPixel"
    Case PropertyTagThumbnailRowsPerStrip
        Name = "ThumbnailRowsPerStrip"
    Case PropertyTagThumbnailStripBytesCount
        Name = "ThumbnailStripBytesCount"
    Case PropertyTagThumbnailResolutionX
        Name = "ThumbnailResolutionX"
    Case PropertyTagThumbnailResolutionY
        Name = "ThumbnailResolutionY"
    Case PropertyTagThumbnailPlanarConfig
        Name = "ThumbnailPlanarConfig"
    Case PropertyTagThumbnailResolutionUnit
        Name = "ThumbnailResolutionUnit"
    Case PropertyTagThumbnailTransferFunction
        Name = "ThumbnailTransferFunction"
    Case PropertyTagThumbnailSoftwareUsed
        Name = "ThumbnailSoftwareUsed"
    Case PropertyTagThumbnailDateTime
        Name = "ThumbnailDateTime"
    Case PropertyTagThumbnailArtist
        Name = "ThumbnailArtist"
    Case PropertyTagThumbnailWhitePoint
        Name = "ThumbnailWhitePoint"
    Case PropertyTagThumbnailPrimaryChromaticities
        Name = "ThumbnailPrimaryChromaticities"
    Case PropertyTagThumbnailYCbCrCoefficients
        Name = "ThumbnailYCbCrCoefficients"
    Case PropertyTagThumbnailYCbCrSubsampling
        Name = "ThumbnailYCbCrSubsampling"
    Case PropertyTagThumbnailYCbCrPositioning
        Name = "ThumbnailYCbCrPositioning"
    Case PropertyTagThumbnailRefBlackWhite
        Name = "ThumbnailRefBlackWhite"
    Case PropertyTagThumbnailCopyRight
        Name = "ThumbnailCopyRight"
    Case PropertyTagLuminanceTable
        Name = "LuminanceTable"
    Case PropertyTagChrominanceTable
        Name = "ChrominanceTable"
    Case PropertyTagFrameDelay
        Name = "FrameDelay"
    Case PropertyTagLoopCount
        Name = "LoopCount"
    Case PropertyTagPixelUnit
        Name = "PixelUnit"
    Case PropertyTagPixelPerUnitX
        Name = "PixelPerUnitX"
    Case PropertyTagPixelPerUnitY
        Name = "PixelPerUnitY"
    Case PropertyTagPaletteHistogram
        Name = "PaletteHistogram"
    Case PropertyTagExifExposureTime
        Name = "ExifExposureTime"
    Case PropertyTagExifFNumber
        Name = "ExifFNumber"
    Case PropertyTagExifExposureProg
        Name = "ExifExposureProg"
    Case PropertyTagExifSpectralSense
        Name = "ExifSpectralSense"
    Case PropertyTagExifISOSpeed
        Name = "ExifISOSpeed"
    Case PropertyTagExifOECF
        Name = "ExifOECF"
    Case PropertyTagExifVer
        Name = "ExifVer"
    Case PropertyTagExifDTOrig
        Name = "ExifDTOrig"
    Case PropertyTagExifDTDigitized
        Name = "ExifDTDigitized"
    Case PropertyTagExifCompConfig
        Name = "ExifCompConfig"
    Case PropertyTagExifCompBPP
        Name = "ExifCompBPP"
    Case PropertyTagExifShutterSpeed
        Name = "ExifShutterSpeed"
    Case PropertyTagExifAperture
        Name = "ExifAperture"
    Case PropertyTagExifBrightness
        Name = "ExifBrightness"
    Case PropertyTagExifExposureBias
        Name = "ExifExposureBias"
    Case PropertyTagExifMaxAperture
        Name = "ExifMaxAperture"
    Case PropertyTagExifSubjectDist
        Name = "ExifSubjectDist"
    Case PropertyTagExifMeteringMode
        Name = "ExifMeteringMode"
    Case PropertyTagExifLightSource
        Name = "ExifLightSource"
    Case PropertyTagExifFlash
        Name = "ExifFlash"
    Case PropertyTagExifFocalLength
        Name = "ExifFocalLength"
    Case PropertyTagExifMakerNote
        Name = "ExifMakerNote"
    Case PropertyTagExifUserComment
        Name = "ExifUserComment"
    Case PropertyTagExifDTSubsec
        Name = "ExifDTSubsec"
    Case PropertyTagExifDTOrigSS
        Name = "ExifDTOrigSS"
    Case PropertyTagExifDTDigSS
        Name = "ExifDTDigSS"
    Case PropertyTagExifFPXVer
        Name = "ExifFPXVer"
    Case PropertyTagExifColorSpace
        Name = "ExifColorSpace"
    Case PropertyTagExifPixXDim
        Name = "ExifPixXDim"
    Case PropertyTagExifPixYDim
        Name = "ExifPixYDim"
    Case PropertyTagExifRelatedWav
        Name = "ExifRelatedWav"
    Case PropertyTagExifInterop
        Name = "ExifInterop"
    Case PropertyTagExifFlashEnergy
        Name = "ExifFlashEnergy"
    Case PropertyTagExifSpatialFR
        Name = "ExifSpatialFR"
    Case PropertyTagExifFocalXRes
        Name = "ExifFocalXRes"
    Case PropertyTagExifFocalYRes
        Name = "ExifFocalYRes"
    Case PropertyTagExifFocalResUnit
        Name = "ExifFocalResUnit"
    Case PropertyTagExifSubjectLoc
        Name = "ExifSubjectLoc"
    Case PropertyTagExifExposureIndex
        Name = "ExifExposureIndex"
    Case PropertyTagExifSensingMethod
        Name = "ExifSensingMethod"
    Case PropertyTagExifFileSource
        Name = "ExifFileSource"
    Case PropertyTagExifSceneType
        Name = "ExifSceneType"
    Case PropertyTagExifCfaPattern
        Name = "ExifCfaPattern"
    Case PropertyTagGpsVer
        Name = "GpsVer"
    Case PropertyTagGpsLatitudeRef
        Name = "GpsLatitudeRef"
    Case PropertyTagGpsLatitude
        Name = "GpsLatitude"
    Case PropertyTagGpsLongitudeRef
        Name = "GpsLongitudeRef"
    Case PropertyTagGpsLongitude
        Name = "GpsLongitude"
    Case PropertyTagGpsAltitudeRef
        Name = "GpsAltitudeRef"
    Case PropertyTagGpsAltitude
        Name = "GpsAltitude"
    Case PropertyTagGpsGpsTime
        Name = "GpsGpsTime"
    Case PropertyTagGpsGpsSatellites
        Name = "GpsGpsSatellites"
    Case PropertyTagGpsGpsStatus
        Name = "GpsGpsStatus"
    Case PropertyTagGpsGpsMeasureMode
        Name = "GpsGpsMeasureMode"
    Case PropertyTagGpsGpsDop
        Name = "GpsGpsDop"
    Case PropertyTagGpsSpeedRef
        Name = "GpsSpeedRef"
    Case PropertyTagGpsSpeed
        Name = "GpsSpeed"
    Case PropertyTagGpsTrackRef
        Name = "GpsTrackRef"
    Case PropertyTagGpsTrack
        Name = "GpsTrack"
    Case PropertyTagGpsImgDirRef
        Name = "GpsImgDirRef"
    Case PropertyTagGpsImgDir
        Name = "GpsImgDir"
    Case PropertyTagGpsMapDatum
        Name = "GpsMapDatum"
    Case PropertyTagGpsDestLatRef
        Name = "GpsDestLatRef"
    Case PropertyTagGpsDestLat
        Name = "GpsDestLat"
    Case PropertyTagGpsDestLongRef
        Name = "GpsDestLongRef"
    Case PropertyTagGpsDestLong
        Name = "GpsDestLong"
    Case PropertyTagGpsDestBearRef
        Name = "GpsDestBearRef"
    Case PropertyTagGpsDestBear
        Name = "GpsDestBear"
    Case PropertyTagGpsDestDistRef
        Name = "GpsDestDistRef"
    Case PropertyTagGpsDestDist
        Name = "GpsDestDist"
    Case Else
        Name = "unknown (" & m_lId & ")"
   End Select
End Property

Public Function ValueCount() As Long
   Select Case ItemType
   Case PropertyTagTypeASCII
      ' each item is 1 byte:
      ValueCount = 1
   
   Case PropertyTagTypeUndefined, PropertyTagTypeByte, 6 ' sbyte
      ValueCount = m_lLength
      
   Case PropertyTagTypeShort, 8 ' schar
      ' each item is 2 bytes:
      ValueCount = m_lLength / 2
      
   Case PropertyTagTypeRational, PropertyTagTypeSRational, 12 ' double
      ' each item is 8 bytes:
      ValueCount = m_lLength / 8
      
   Case PropertyTagTypeLong, PropertyTagTypeSLONG, 11 ' float
      ' each item is 4 bytes:
      ValueCount = m_lLength / 4
      
   End Select

End Function

Public Property Get ParseString() As String
    'returns the value of an Ascii type entry
    ParseString = TrimNulls(StrConv(m_bData, vbUnicode))
End Property
Public Function SetPropertyStringValue(inVal As String, Optional id As Long) As Boolean
    'used to write a new property or to update an existing one
    'property must be of ASCII type

    If id = 0 Then
        'if they didnt supply an id then assume they wanted to update the
        'currently loaded property, if of course there is one
        If m_lId <> 0 Then
            id = m_lId
        Else
            Exit Function
        End If
    End If

    'does this property id already exist?
    'if it does we need to confirm it is Ascii before we try and write to it
    If Not IsArrayEmpty(Not m_bData) Then
        'The array already contains some data
        'check to see that this item is ok to write ascii to
        If ItemType <> PropertyTagTypeASCII Then
            Exit Function
        End If
    Else
        'a new entry will be created
    End If

    'initialise our item
    CreateProperty id, Len(inVal), PropertyTagTypeASCII, StrConv(inVal, vbFromUnicode)

    SetPropertyStringValue = True
End Function


Public Property Get ParseRational(ByVal lItem As Long) As Variant
Dim lStart As Long
   If (lItem > 0) And (lItem <= ValueCount) Then
      lStart = (lItem - 1) * 8
      ReDim lValue(1 To 2) As Long
      CopyMemory lValue(1), m_bData(lStart), 4
      CopyMemory lValue(2), m_bData(lStart + 4), 4
      ParseRational = lValue
   Else
      'SetStatusHelper InvalidParameter
   End If
End Property
Public Property Get ParseShort(ByVal lItem As Long) As Integer
Dim lStart As Long
   If (lItem > 0) And (lItem <= ValueCount) Then
      lStart = (lItem - 1) * 2
      Dim iRet As Integer
      CopyMemory iRet, m_bData(lStart), 2
      ParseShort = iRet
   Else
      'SetStatusHelper InvalidParameter
      lStart = lStart
   End If
End Property
Public Property Get ParseLong(ByVal lItem As Long) As Long
Dim lStart As Long
   If (lItem > 0) And (lItem <= ValueCount) Then
      lStart = (lItem - 1) * 4
      Dim iRet As Long
      CopyMemory iRet, m_bData(lStart), 4
      ParseLong = iRet
   Else
      'SetStatusHelper InvalidParameter
   End If
End Property

Public Sub GetData(ByRef b() As Byte)
On Error Resume Next
Dim i As Long
   If (m_lLength > 0) Then
      For i = LBound(b) To UBound(b)
         b(i) = m_bData(i - LBound(b))
      Next i
   End If
End Sub

Public Property Get DataBufferSize() As Long
   DataBufferSize = ElementDataSize() * m_lLength
End Property

Public Property Get ElementDataSize() As Long
Dim lSize As Long
   Select Case ItemType
   Case PropertyTagTypeASCII, PropertyTagTypeUndefined, PropertyTagTypeByte, 6 ' sbyte
      ' each item is 1 byte:
      lSize = 1
   Case PropertyTagTypeShort, 8 ' schar
      ' each item is 2 bytes:
      lSize = 2
   Case PropertyTagTypeRational, PropertyTagTypeSRational, 12 ' double
      ' each item is 8 bytes:
      lSize = 8
   Case PropertyTagTypeLong, PropertyTagTypeSLONG, 11 ' float
      ' each item is 4 bytes:
      lSize = 4
   End Select
   ElementDataSize = lSize
End Property


'-----------------
'PRIVATE FUNCTIONS
'-----------------
Friend Sub fInit( _
      ByVal lId As Long, _
      ByVal lLength As Long, _
      ByVal eItemType As PropertyTagType, _
      ByVal lPtr As Long, _
      ByVal lSize As Long _
   )
    'used by cGDIP to create a property item
   m_lId = lId
   m_lLength = lLength
   m_itemType = eItemType
   If Not (lPtr = 0) And (lLength > 0) Then
      Dim lDataSize As Long
      lDataSize = lSize - 16
      If (lDataSize > 0) Then
         ReDim m_bData(0 To lDataSize - 1) As Byte
         CopyMemory m_bData(0), ByVal lPtr, lDataSize
      End If
   End If
End Sub

Private Sub CreateProperty( _
      ByVal lId As Long, _
      ByVal lLength As Long, _
      ByVal eItemType As PropertyTagType, _
      ByRef b() As Byte _
   )
    'used by us to create a property item
        Dim i As Long
        m_lId = lId
        m_lLength = lLength
        m_itemType = eItemType
        If (m_lLength > 0) Then
           ReDim m_bData(LBound(b) To UBound(b))
           For i = LBound(b) To UBound(b)
              m_bData(i - LBound(b)) = b(i)
           Next i
        End If
End Sub

Private Function TrimNulls(ByRef sText As String) As String
    Dim lPos As Long
    Dim sLeft As String
    Dim sRight As String
    'removes all the nulls out of a string
    lPos = InStr(1, sText, Chr$(0))
    Do While lPos > 0
        sLeft = Mid$(sText, 1, lPos - 1)
        sRight = Mid$(sText, lPos + 1, Len(sText) - lPos)
        sText = sLeft & sRight
        
       lPos = InStr(1, sText, Chr$(0))
    Loop
    TrimNulls = sText
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Long
  'use like this
  '--Dim b() As Byte
  '--If IsArrayEmpty(Not b) Then
  '--  Debug.Print "The array has not been Redimmed."
  '--Else
  '--  Debug.Print "The array has been Redimmed."
  '--End If
  
  'lets you know if your array has any data in it
  IsArrayEmpty = (lArrayPointer = -1)
End Function
