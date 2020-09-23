Attribute VB_Name = "ModPic"
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub StorePic(objPic As PictureBox)
Dim lDC As Long, lBMP As Long
With objPic
    .Visible = False
    '- create a Device Context to store container section
    lDC = CreateCompatibleDC(.hDC)
    lBMP = CreateCompatibleBitmap(.hDC, .Width, .Height)
    SelectObject lDC, lBMP
    BitBlt lDC, 0, 0, .Width, .Height, .hDC, 0, 0, vbSrcCopy
    .Visible = True
End With
    DeleteDC lDC
    DeleteObject lBMP
    
End Sub


