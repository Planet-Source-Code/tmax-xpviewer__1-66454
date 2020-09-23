Attribute VB_Name = "BrowseforSpecificFolder"
Option Explicit

Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type


'****************
'API declarations
'****************
 Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

 Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                          (lpBrowseInfo As BROWSEINFO) As Long

 Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                          (ByVal pidl As Long, _
                          ByVal pszPath As String) As Long
    
 Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
 
 Public Declare Sub MoveMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
    Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
  lParam As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)

Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Const MAX_PATH = 255
'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_STATUSTEXT As Long = &H4


Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

Select Case uMsg
    Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)
    Case Else
End Select
End Function

Public Function GetWindowHandle(strClassName As String, strWindowName As String) As Long
'as VBA does not support a Hwnd(window handle)property, we have to
'use this function to get the hwnd
'"ThunderDFrame" is the classname for VBA forms, but "ThunderFormDC"
'is the classname for VB forms, although this function is not needed
'for VB. The windowname is always the form's caption property.
GetWindowHandle = FindWindow(strClassName, strWindowName)

End Function
Public Function AddressOfCallBack(Address As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
 
   AddressOfCallBack = Address

End Function
'---------------------------------------------
' Function: BrowseForFolderDlg
' Action: Invokes the Windows Browse for Folder dialog
' Return: If successful, returns the selected folder's full path,
' returns an empty string otherwise.
' -------------------------------------------------
Public Function BrowseForFolderDlg(strInitialFolder As String, strDialogPrompt As String, hwnd As Long, Optional IncludeFiles As Boolean = False) As String
    Dim BI As BROWSEINFO
    Dim lngPidlRtn As Long
    Dim strPath As String * MAX_PATH ' buffer
    Dim lpPath As Long
    
    On Error GoTo ErrHandler
    strInitialFolder = strInitialFolder + IIf(Right$(strInitialFolder, 1) <> "\", "\", "")
    With BI
        'verify that the directory is valid
        If strInitialFolder <> "" Then
            If GetAttr(strInitialFolder) And vbDirectory Then
                'allocate memory for our string
                lpPath = LocalAlloc(lPtr, Len(strInitialFolder))
                'fill the memory with the contents of the string
                MoveMemory ByVal lpPath, ByVal strInitialFolder, Len(strInitialFolder)
                .lpfn = AddressOfCallBack(AddressOf BrowseCallbackProc)
                .lParam = lpPath
            End If
        End If
        .ulFlags = BIF_RETURNONLYFSDIRS + IIf(IncludeFiles, BIF_BROWSEINCLUDEFILES, 0)
        
   
'        .ulFlags = 0 '1
        ' Whoever owns the handle that we pass will own the dialog
        ' The desktop folder will be the dialog's root folder if this
        'is initialized to 0.
        .hOwner = hwnd
        
        'SHSimpleIDListFromPath can also be used to set this value.
        .pidlRoot = 0
        ' Set the dialog's prompt string
        .lpszTitle = strDialogPrompt
    End With
    
    ' Shows the browse dialog and doesn't return until the dialog is
    ' closed. lngpidlRtn will contain the pidl of the selected folder if the dialog is not cancelled.
    lngPidlRtn = SHBrowseForFolder(BI)
    
    If lngPidlRtn Then
    ' Get the path from the selected folder's pidl returned
    ' from the SHBrowseForFolder call (rtns True on success,
    ' strPath must be pre-allocated!)
        If SHGetPathFromIDList(lngPidlRtn, strPath) Then
      ' Return the path
            BrowseForFolderDlg = Left$(strPath, InStr(strPath, vbNullChar) - 1)
        End If
    ' Free the memory the shell allocated for the selected folder's pidl.
        Call CoTaskMemFree(lngPidlRtn)
        
    End If
    'free the memory that we allocated for the pre-selected folder's pidl
    Call LocalFree(BI.lParam)
    
    Exit Function
ErrHandler:
    If lngPidlRtn Then
       Call CoTaskMemFree(lngPidlRtn)
    End If
    If lpPath Then
        Call LocalFree(lpPath)
    End If
    BrowseForFolderDlg = ""
End Function

