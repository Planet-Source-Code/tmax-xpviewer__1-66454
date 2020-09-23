Attribute VB_Name = "RegFunc"
Option Explicit
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0

Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)

    Dim lRetVal As Long
    Dim hKey As Long

    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)

End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long
       Dim hKey As Long

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
       
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select

End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function

Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    
    Dim hNewKey As Long
    Dim lRetVal As Long
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function

Sub Example()
    'Examples of each function:
    'CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test", "Testing, Testing", REG_SZ
    'MsgBox QueryValue(HKEY_CURRENT_USER, "TestKey\SubKey1", "Test")
    'DeleteKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'DeleteValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test"
End Sub

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)

       Dim lRetVal As Long
       Dim hKey As Long

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = StripNull(vValue)           '*** Must Strip Null for vValue
       RegCloseKey (hKey)
End Function

Public Sub Check1stTime()
Dim ret
ret = QueryValue(HKEY_CURRENT_USER, "XPViewer", "")  ' Check for 1st time user & Application Version
'InitReg
If ret = "" Then
    ShowMsg "Initialize registry...", vbOKOnly, "Check Version":
    ReturnMsg = "Yes"
    Unload FrmMessage
    InitReg
     frmAbout.Show 1
End If
End Sub

'First Time Use
'HKEY_CURRENT_USER\XPViewer
'run AppPATH + "XPViewer.reg"
Public Sub InitReg()
SearchJpgType
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Control"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Control\Picture"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Control\WallPaper"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Folder"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Prefix"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\PlugIn"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Application"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Attach"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Attach\Font"
CreateNewKey HKEY_CURRENT_USER, "XPViewer\Attach\Offset"
SetKeyValue HKEY_CURRENT_USER, "XPViewer", "", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "LastPathSelect", FolderLocation(CSIDL_Desktop), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "CdBurnPath", FolderLocation(CSIDL_Desktop) & "\CDBurn", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "StartPath", FolderLocation(CSIDL_Desktop), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "SmartPath", FolderLocation(CSIDL_Desktop), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "FavoritePath", FolderLocation(CSIDL_Desktop), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "Dest 1", FolderLocation(CSIDL_Desktop), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "DestCount", 1, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "Picture", App.Path + "\Picture\", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "BackGroundColor", "0", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "TypeJpeg", Type_JPEG, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "Favorite", FolderLocation(CSIDL_PERSONAL), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailSize", "96", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailShadow", "True", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "SlideTimer", "1000", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugInCount", 1, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugIn 1", WindowsFolder & "system32\MsPaint.exe", REG_SZ
'SetKeyValue HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugIn 2", "D:\Program Files\Adobe\Photoshop CS\Photoshop.exe", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "PreCount", 5, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre 1", "Img##", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre 2", "Image##", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre 3", "Pic##", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre 4", "< Original >", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre 5", "< Time Stamp >", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Application", "BurnSoftware", "C:\Program Files\Ahead\Nero\nero.exe /w", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach", "Date", Format$(Now, "dd-mm-yyyy"), REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach", "Check", "TRUE", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Name", "Verdana", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Size", "10", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Color", "8388863", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "x", "100", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "y", "100", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "Filename", "", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackColor", "12632256", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackPicture", "", REG_SZ
End Sub



