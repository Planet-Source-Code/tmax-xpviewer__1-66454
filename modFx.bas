Attribute VB_Name = "modFx"

Option Explicit
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Enum WallPaperMode
    Stretch = 0
    Tile = 1
    Center = 2
End Enum
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
' NOTE: Oringinal enum was unnamed
Public Enum SHFolders
    CSIDL_Desktop = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_templates = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum
Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * 260
    cAlternate       As String * 14
End Type
Public Const DT_CENTER = &H1
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_TOP = &H0
Public Const DT_BOTTOM = &H8

Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1      'Destination specifies multiple files
Const FOF_SILENT = &H4              'Don't display progress dialog
Const FOF_RENAMEONCOLLISION = &H8   'Rename if destination already exists
Const FOF_NOCONFIRMATION = &H10     'Don't prompt user
Const FOF_WANTMAPPINGHANDLE = &H20  'Fill in hNameMappings member
Const FOF_ALLOWUNDO = &H40          'Store undo information if possible
Const FOF_FILESONLY = &H80          'On *.*, don't copy directories
Const FOF_SIMPLEPROGRESS = &H100    'Don't show name of each file
Const FOF_NOCONFIRMMKDIR = &H200    'Don't confirm making any needed dirs


Private Const MAX_PATH = 260
Public buffer As String * MAX_PATH
Public Message$, Title$, ReturnMsg$
Public Enum MsgStyle
vbOKOnly = 0
vbOKCancel = 1
vbYesNo = 2
End Enum
Public Enum FileOpt
  F_CopySmart = 1 'FO_COPY
  F_COPY = 2
  F_Move = 4
  F_Delete = 8
  F_DelUndo = 16
  F_Vault = 64
  F_Rename = 128
End Enum
         
      
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function AlphaBlend Lib "MSImg32.dll" (ByVal hDcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) As Long
' Drag Form Declaration
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef S As SHELLEXECUTEINFO) As Long
Public Declare Function fCreateShellLink Lib "Vb5stkit.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public FilePath As String
Public FileSelect As String
Public FileSelectIndex As Integer
Public LastPathSelect As String
Public fsys As New FileSystemObject
Public WallTiles$
Public WallFileName$
Public WallBackColor$
Public WallBackPicture$
Public AddSoftwareName As Boolean
Public Shadow As String
Public ChkCRC As Boolean
Public Type_JPEG As String
Public FavoritePath As String
Public PlugInSoftware(5) As String  ' let 6 plugIns software
Public PluginCount As Integer
Public DateAttach As String
Public DFontName$
Public DFontSize$
Public DFontColor$
Public DFormat$
Public OffsetX$
Public OffsetY$
Public lwFontAlign As Long
Public ThumbnailSize As String
Public SlideTimer As String
Public ChkDateAttach As String
Public CDBurnPath As String
Public StartPath As String
Public SmartPath As String
Public CompanyName As String
Public BurnSoftware As String
Public FolderOper As Boolean

Public FileDrag As New Collection
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Function DiskOps(Source As String, Dest As String, Flavor As FileOpt, Success As Long) As Long
'On Error GoTo DiskOpsErr
Dim result As Long
Dim FileOp As SHFILEOPSTRUCT
With FileOp
.hwnd = 0
   Select Case Flavor
      Case 1                           ' SmartCopy
         .wFunc = FO_COPY
         .fFlags = FOF_NOCONFIRMATION
      Case 2                           ' Copy
         .wFunc = FO_COPY
      Case 4                           ' Move
         .wFunc = FO_MOVE
      Case 8                           ' Delete
         .wFunc = FO_DELETE
      Case 16                          ' Delete (Recycle bin)
         .wFunc = FO_DELETE
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      Case 64                          ' Vault
         .wFunc = FO_COPY
         .fFlags = FOF_MULTIDESTFILES
      Case 128                         ' Rename
         .wFunc = FO_RENAME
   End Select
  '.lpszProgressTitle = ""
   .pFrom = Source & vbNullChar & vbNullChar    ' The files to copy separated by Nulls and terminated by 2 nulls
   .pTo = Dest & vbNullChar & vbNullChar                ' The directory or filename(s) to copy into terminated in 2 nulls
End With
   result = SHFileOperation(FileOp)
   DiskOps = result
   If result <> 0 Then 'Operation failed
      'Msgbox the error that occurred in the API.
   '''  MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & Err.LastDllError, vbCritical Or vbOKOnly
       ShowMsg "Cannot " & FileOp.wFunc & Source & vbCrLf & "to " & Dest, vbOKOnly, "Critical Error"
   Else
      If FileOp.fAnyOperationsAborted <> 0 Then
        MsgBox "Operation Failed", vbCritical Or vbOKOnly
         Success = -1
      End If
   End If
End Function

Public Function bFileExists(FileName As String) As Boolean
      Dim TempAttr         As Integer
   On Error GoTo ExitFileExist 'any errors show that the file doesnt exist, so goto this label
   TempAttr = GetAttr(FileName)  'get the attributes of the files
   bFileExists = ((TempAttr And vbDirectory) = 0) 'check if its a directory and not a file
ExitFileExist:
   On Error GoTo 0 'clear all errors
End Function

Public Function GetATemporaryFileName() As String
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
    GetTempFileName sTempDir, "DEK", 0, sTempFileName
    'Strip the 0's off the end
    sTempFileName = Left$(sTempFileName, InStr(sTempFileName, Chr$(0)) - 1)
    GetATemporaryFileName = sTempFileName
End Function

Public Sub SearchJpgType()
Dim JpgDefault$
JpgDefault = QueryValue(HKEY_CLASSES_ROOT, ".jpg", "")
Type_JPEG = QueryValue(HKEY_CLASSES_ROOT, JpgDefault, "")
'search registry for .jpg  default value
' search for that value
' search for defaut value -->type jpg
End Sub

Public Function FolderLocation(lFolder As SHFolders) As String
   Dim lp As Long
   'Get the PIDL for this folder
   SHGetSpecialFolderLocation 0&, lFolder, lp
   SHGetPathFromIDList lp, buffer
   FolderLocation = StripNull(buffer)
   'Free the PIDL
   CoTaskMemFree lp
End Function

Public Function StripNull(ByVal StrIn As String) As String
On Error Resume Next
Dim nul As Long
nul = InStr(StrIn, vbNullChar)
    Select Case nul
        Case Is > 1
            StripNull = Left$(StrIn, nul - 1)
        Case 1
            StripNull = ""
        Case 0
            StripNull = Trim$(StrIn)
   End Select
End Function

Public Function GetFName(FileName As String) As String
Dim fs, F
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFile(FileName)
    GetFName = F.Name
    Set F = Nothing
    Set fs = Nothing
End Function

Public Function GetFPath(FileName As String) As String
Dim fs, F
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFile(FileName)
    GetFPath = F.ParentFolder + IIf(Right(F.ParentFolder, 1) <> "\", "\", "")
    Set F = Nothing
    Set fs = Nothing
End Function

Function DragForm(frm As Form)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function

Public Function CheckDiskette() As Boolean
On Error GoTo CheckDiskette_Error
'set default for error flag

'this will generate a trappable error if card not inserted
ChDir "a:"
'go back to original hard drive
ChDir "c:"
On Error GoTo 0
CheckDiskette = True
Exit Function
CheckDiskette_Error:
If Err.Number = 75 Then
   ShowMsg "Please check that the diskette " & vbCrLf & "is properly inserted and try again.", vbOKOnly, "Xpress Viewer"
   CheckDiskette = False
    Exit Function
Else
    ShowMsg "Error " & Err.Number & vbCrLf & " (" & Err.Description & ") in procedure CheckDiskette", vbOKOnly, "Xpress Viewer"
Err.Clear
End If
End Function

Sub FormatDrive(strDrive As String)
Dim strDriveLetter As String
Dim lngDriveNumber As Long
Dim lngRetVal As Long
Dim lngDriveType As Long
Dim lngRet As Long

strDriveLetter = UCase(strDrive)
lngDriveNumber = (Asc(strDriveLetter) - 65) ' Change letter to Number: A=0
lngDriveType = GetDriveType(strDriveLetter)

If lngDriveType = 2 Then 'Floppies, etc
lngRetVal = SHFormatDrive(FrmMdi.hwnd, lngDriveNumber, 0&, 0&)
Else
lngRet = MsgBox("This drive is NOT a removeable" & vbCrLf & _
"drive! Format this drive?", 276, "SHFormatDrive Example")
Select Case lngRet
Case 6 'Yes
lngRetVal = SHFormatDrive(FrmMdi.hwnd, lngDriveNumber, 0&, 0&)
Case 7 'No
' Do nothing
End Select
End If
End Sub

Public Sub FileProperties(FileName$)
Dim shInfo As SHELLEXECUTEINFO

    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = FileName$
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With
    ShellExecuteEx shInfo

End Sub

Public Sub SetWallPaper(PathName$, Mode As WallPaperMode)
On Error Resume Next
Dim ximg As cIMAGE
Dim WallPaperPath As String
Dim wStyle, wTile As String
Dim ScreenHeight As Integer, ScreenWidth As Integer
Dim ImgRatio As Single, ScreenRatio As Single
Dim PicWall As PictureBox
    WallPaperPath = WindowsFolder & GetFName(PathName)
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    ScreenRatio = Screen.Width / Screen.Height
    If Mode = Stretch Then
        'Get monitor display settings.
        With PicWall
            .Width = ScreenWidth
            .Height = ScreenHeight
            .BackColor = WallBackColor
            .Picture = LoadPicture(WallBackPicture)
            Set ximg = New cIMAGE
            ximg.Load PathName
            If ximg.ImageWidth > ximg.ImageHeight Then
                ximg.ReSize .Width, 0, True
            Else
                ximg.ReSize 0, .Height, True
            End If
            ximg.PaintDC .hDC, (ScreenWidth - ximg.ImageWidth) / 2, (ScreenHeight - ximg.ImageHeight) / 2
            Set ximg = Nothing
            .Picture = .Image
            SavePicture .Picture, WallPaperPath
            .Picture = LoadPicture
            .Cls
        End With
    Else
            Set ximg = New cIMAGE
            ximg.Load PathName
            If ximg.ImageWidth > ximg.ImageHeight Then
                ximg.ReSize ScreenWidth \ Val(WallTiles), 0, 0
            Else
                ximg.ReSize 0, ScreenHeight \ Val(WallTiles), 0
            End If
            'set the path to the wallpaper bitmap
            ' save as bitmap , so desktop will automatically replace wallpaper & refresh
            ' If save as Jpeg , it will not refresh desktop
            'SavePicture LoadPicture(PathName), WallPaperPath
            SavePicture ximg.Picture, WallPaperPath
            Set ximg = Nothing
    End If
Select Case Mode
    Case 0    'Stretch wallpaper
        wStyle = "2"
        wTile = "0"
    Case 1 'Tile wallpaper
        wStyle = "0"
        wTile = "1"
    Case 2 'Center wallpaper
        wStyle = "0"
        wTile = "0"
End Select
Set PicWall = Nothing
    SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", wStyle, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", wTile, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", WallPaperPath, REG_SZ
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, WallPaperPath, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub

Public Function WindowsFolder() As String
Dim buffer  As String * MAX_PATH
GetWindowsDirectory buffer, 255
WindowsFolder = StripNull(buffer)
WindowsFolder = WindowsFolder + IIf(Right(WindowsFolder, 1) <> "\", "\", "")
End Function

Sub DropShadow(hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Function ShowMsg(Msg As String, style As MsgStyle, TitleMsg As String) As Boolean
ShowMsg = False
With FrmMessage
Select Case style
    Case vbOKOnly:
        .CmdYes.Caption = "Ok"
        .CmdYes.Left = (.ScaleWidth - .CmdYes.Width) / 2
        .CmdNo.Visible = False
    Case vbOKCancel:
        .CmdYes = "Ok"
        .CmdNo.Caption = "Cancel"
    Case vbYesNo:
        .CmdYes.Caption = "Yes"
        .CmdNo.Caption = "No"
End Select
    .LblMsg = Msg
    .LblMsg2 = Msg
    .LblTitle = TitleMsg
    .LblTitle2 = TitleMsg
    .Show 1
End With
If ReturnMsg = "Yes" Or ReturnMsg = "Ok" Then
    ShowMsg = True
End If
End Function


