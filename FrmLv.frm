VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLv 
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11400
   Begin VB.PictureBox PicExif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   10800
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   5
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      Picture         =   "FrmLv.frx":0CCE
      ScaleHeight     =   360
      ScaleWidth      =   11400
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      Begin ProXPViewer.TMcmdbutton TMcmdClose 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   40
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   400
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   480
         Picture         =   "FrmLv.frx":585D
         Stretch         =   -1  'True
         Top             =   40
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox PicThumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      ScaleHeight     =   75.077
      ScaleMode       =   0  'User
      ScaleWidth      =   75.077
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ListView Lv1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10186
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Attributes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CRC32"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
FrmMdi.sbStatusBar.Panels(3).Text = Lv1.ListItems.Count & " object(s)"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FrmMdi.PicPreview.Picture = LoadPicture
FrmMdi.PicPreview.Cls
End Sub

Private Sub Form_Resize()
On Error Resume Next
Lv1.Top = PicTop.ScaleHeight
Lv1.Height = Me.ScaleHeight - PicTop.ScaleHeight
pb1.Width = PicTop.ScaleWidth - pb1.Left - 20
Lv1.Width = Me.ScaleWidth
Lv1.Arrange = lvwAutoTop
End Sub

Private Sub Lv1_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim retval As Long
retval = DiskOps(Lv1.SelectedItem.Key, Caption & NewString, F_Rename, 1)
If retval = 183 Then
    Cancel = -1
Else
Lv1.SelectedItem.Text = NewString
Lv1.SelectedItem.Key = Caption & NewString
End If
End Sub

Private Sub Lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Lv1.SortKey = ColumnHeader.Index - 1
If (Lv1.Sorted = False) Then
    Lv1.Sorted = True
    Lv1.SortOrder = 0
Else
    If (Lv1.SortOrder = 0) Then
        Lv1.SortOrder = 1
    Else
        Lv1.SortOrder = 0
    End If
End If
End Sub

Private Sub Lv1_DblClick()
Dim fpreview As FrmPreview
Set fpreview = New FrmPreview
FileSelect = Lv1.SelectedItem.Key
fpreview.Show
With fpreview
    .LoadPV FileSelect
    .CurrentPic = Lv1.SelectedItem.Index
    .LoadLv
End With
Set fpreview = Nothing
End Sub

Private Sub Lv1_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim i%, j%
For i% = 1 To Lv1.ListItems.Count
    If Lv1.ListItems(i%).Selected Then j% = j% + 1
Next i%
FrmMdi.sbStatusBar.Panels(1).Text = ""
FrmMdi.sbStatusBar.Panels(2).Text = ""
FrmMdi.sbStatusBar.Panels(3).Text = j% & " object(s)"
If j% = 1 Then
    FrmMdi.sbStatusBar.Panels(1).Text = item.Text
    FrmMdi.sbStatusBar.Panels(2).Text = SetBytes(item.Tag)
    FrmMdi.LoadPreview item.Key
    FrmMdi.ShowExif item.Key
End If
End Sub

Sub RefreshLv(npath As String)
On Error Resume Next
Me.MousePointer = 11
Set Lv1.Icons = Nothing
CreateThumbs npath, Val(ThumbnailSize) '96
AddFiles npath
Me.MousePointer = 0
End Sub

Sub AddFiles(Path As String)
On Error Resume Next
Dim itmx As ListItem
Dim fso, F, fc, fj, f1
Dim i%
Dim xcrc As clsCRC
If ChkCRC Then
    Set xcrc = New clsCRC
    xcrc.Algorithm = CRC32
End If
Set fso = CreateObject("Scripting.FileSystemObject")
Lv1.ListItems.Clear
Set Lv1.Icons = ImgList1
If fso.FolderExists(Path) Then
    Set F = fso.GetFolder(Path)
    Set fj = F.Files
    For Each f1 In fj
        If f1.Type = Type_JPEG Then
            i% = i% + 1
            Set itmx = Lv1.ListItems.Add(i, f1.Path, f1.Name, i)
            itmx.SubItems(1) = KBytes(f1.Size)
            itmx.SubItems(2) = f1.Type
            itmx.SubItems(3) = Format$(f1.DateLastModified, "d/m/yyyy h:m AM/PM")
            itmx.SubItems(4) = splitAttr(f1.Attributes)     '1-R 2-H 32-A
            If ChkCRC Then
                xcrc.CalculateFile f1.Path
                itmx.SubItems(5) = Hex(xcrc.Value)
            End If
            itmx.Tag = f1.Size
        End If
    Next
    Set F = Nothing
    Set fj = Nothing
    Set f1 = Nothing
End If
Set fso = Nothing
If ChkCRC Then Set xcrc = Nothing
FrmMdi.sbStatusBar.Panels(3).Text = Lv1.ListItems.Count & " object(s)"
FrmMdi.sbStatusBar.Panels(4).Text = SetBytes(GetDiskSpaceFree(Mid(Path, 1, 2)))
End Sub

Public Function KBytes(Bytes) As String
On Error GoTo KBerror
    If Bytes >= 1073741824 Then
        KBytes = Format(Bytes / 1024, "#0,000,000") & " KB"
    ElseIf Bytes >= 1048576 Then
        KBytes = Format(Bytes / 1024, "#0,000") & " KB"
    ElseIf Bytes >= 1024 Then
        KBytes = (Format(Bytes / 1024, "#0") + 1) & " KB"
    ElseIf Bytes < 1024 Then
        KBytes = "1 KB"
    End If
    Exit Function
KBerror:
    KBytes = "0 Bytes"
End Function

Public Function SetBytes(Bytes) As String
On Error GoTo SBerror
    If Bytes >= 1073741824 Then
        SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") & " GB"
    ElseIf Bytes >= 1048576 Then
        SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
    ElseIf Bytes >= 1024 Then
        SetBytes = (Format(Bytes / 1024, "#0.00") + 1) & " KB"
    ElseIf Bytes < 1024 Then
        SetBytes = Bytes
    End If
    Exit Function
SBerror:
    SetBytes = "0 Bytes"
End Function

Private Sub Lv1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i%
If KeyCode = vbKeyDelete Then
    FrmMdi.OperDelete
End If
'Ctrl + a
If KeyCode = 65 And Shift = 2 And Lv1.ListItems.Count Then
    For i% = 1 To Lv1.ListItems.Count
        Lv1.ListItems(i%).Selected = True
    Next i%
    FrmMdi.sbStatusBar.Panels(1).Text = ""
    FrmMdi.sbStatusBar.Panels(2).Text = ""
    FrmMdi.sbStatusBar.Panels(3).Text = Lv1.ListItems.Count & " object(s)"
End If
' Ctrl + r
If KeyCode = 82 And Shift = 2 And Lv1.ListItems.Count Then
    RefreshLv Me.Caption
End If
End Sub

Private Sub Lv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu FrmMdi.MnuListView
End Sub

Private Sub Lv1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Data.GetFormat(vbCFText) Then
    Effect = vbDropEffectMove And Effect
Else
    Effect = vbDropEffectMove
End If
End Sub

Private Sub Lv1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Dim itmx As ListItem
Dim i%
Set itmx = Lv1.SelectedItem
    Data.SetData , vbCFFiles
For Each itmx In Lv1.ListItems
    If itmx.Selected Then
        Data.Files.Add itmx.Key
    End If
Next
End Sub

Private Sub Lv1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim fs, F
Dim i%
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
If Data.GetFormat(vbCFFiles) Then
    For i% = 1 To Data.Files.Count
        Set F = fs.GetFile(Data.Files(i%))
        DiskOps F.Path, Me.Caption & F.Name, F_CopySmart, 1
        Set F = Nothing
    Next i%
End If
RefreshLv Me.Caption
Set fs = Nothing
End Sub

Function splitAttr(attr As Integer) As String     '1-R 2-H 32-A
If attr < 1 Then splitAttr = "": Exit Function
    Select Case attr
        Case 1
            splitAttr = "R"
        Case 2
            splitAttr = "H"
        Case 3
            splitAttr = "RH"
        Case 32
            splitAttr = "A"
        Case 33
            splitAttr = "AR"
        Case 34
            splitAttr = "AH"
        Case 35
            splitAttr = "ARH"
    End Select
End Function

Private Sub PicTop_DblClick()
If Me.WindowState <> 2 Then
    Me.WindowState = 2
Else
    Me.WindowState = 0
End If
End Sub

Private Sub PicTop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub TMcmdClose_Click()
Unload Me
End Sub

Public Function GetDiskSpaceFree(ByVal strDrive As String) As Long
    Dim lRet As Long
    Dim lBytes As Long
    Dim lSect As Long
    Dim lClust As Long
    Dim lTot As Long
    On Error Resume Next
    GetDiskSpaceFree = -1
    If True Then 'GetDrive(strDrive, strDrive) Then
        lRet = GetDiskFreeSpace(strDrive, lSect, lBytes, lClust, lTot)
        If Err.Number = 0 Then
            If lRet <> 0 Then
                GetDiskSpaceFree = lBytes * lSect * lClust
                If Err.Number <> 0 Then
                    GetDiskSpaceFree = &H7FFFFFFF
                End If
            End If
        End If
    End If
    Err.Clear
End Function

Sub CreateThumbs(Path As String, Size As Integer)
On Error Resume Next
Dim ximg As cIMAGE
Dim W%, H%
Dim fso, F, fc, fj, f1
Dim i%, j%
Dim c%
Dim uRect As RECT
Dim hBrush As Long
Me.scaleMode = 3
PicThumb.Width = Size + 8
PicThumb.Height = Size + 8
Me.scaleMode = 1
If Shadow Then hBrush = CreateSolidBrush(RGB(200, 200, 200))
ImgList1.ListImages.Clear
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(Path) Then
    Set F = fso.GetFolder(Path)
    Set fj = F.Files
    Me.scaleMode = 3
    pb1.Max = F.Files.Count
    Image1.Visible = True
    For Each f1 In fj
        j% = j% + 1
        'If f1.Type = "ACDSee JPEG Image" Then
        If f1.Type = Type_JPEG Then
            i% = i% + 1
            Set ximg = New cIMAGE
            With PicThumb
                .Cls
                ximg.Thumbnail f1.Path, Size, Size
                W = ximg.ImageWidth
                H = ximg.ImageHeight
                If Shadow Then
                    SetRect uRect, (.Width - W) / 2, (.Height - H) / 2, (.Width + W) / 2, (.Height + H) / 2
                    FillRect .hDC, uRect, hBrush
                End If
                ximg.PaintDC .hDC, (.Width - W) / 2 - 3, (.Height - H) / 2 - 3
                ImgList1.ListImages.Add i%, f1.Path, .Image
            End With
            Set ximg = Nothing
        End If
        pb1.Value = j
        Image1.Width = pb1.Value / pb1.Max * (PicTop.ScaleWidth - (Image1.Left * 1.3))
    Next
    Image1.Visible = False
    Set F = Nothing
    Set fj = Nothing
    Set f1 = Nothing
    Me.scaleMode = 1
End If
Set fso = Nothing
End Sub

Sub TCreateThumbs(Path As String, Size As Integer)
On Error Resume Next
Dim ximg As cIMAGE
Dim W%, H%
Dim fso, F, fc, fj, f1
Dim i%, j%
Dim c%
Dim pnts(2) As POINTAPI
Me.scaleMode = 3
PicThumb.Width = Size + 8
PicThumb.Height = Size + 8
Me.scaleMode = 1
ImgList1.ListImages.Clear
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(Path) Then
    Set F = fso.GetFolder(Path)
    Set fj = F.Files
    Me.scaleMode = 3
    pb1.Max = F.Files.Count
    Image1.Visible = True
    For Each f1 In fj
        j% = j% + 1
        'If f1.Type = "ACDSee JPEG Image" Then
        If f1.Type = Type_JPEG Then
            i% = i% + 1
            Set ximg = New cIMAGE
            With PicThumb
                .Cls
                ximg.Thumbnail f1.Path, Size, Size
                W = ximg.ImageWidth
                H = ximg.ImageHeight
                ximg.PaintDC .hDC, (.Width - W) / 2 - 3, (.Height - H) / 2 - 3
                If Shadow = "True" Then
                    If H > W Then
                        For c% = 0 To 2
                            .Forecolor = RGB(220 - 20 * c%, 220 - 20 * c%, 220 - 20 * c%)
                            pnts(0).x = (.Width + W) / 2 - 3 + c%
                            pnts(0).y = (.Height - H) / 2
                            pnts(1).x = (.Width + W) / 2 - 3 + c%
                            pnts(1).y = (.Height + H) / 2
                            Polyline .hDC, pnts(0), 2
                            pnts(0).x = (.Width - W) / 2
                            pnts(0).y = .Height - 7 + c%
                            pnts(1).x = (.Width + W) / 2
                            pnts(1).y = .Height - 7 + c%
                            Polyline .hDC, pnts(0), 2
                        Next c%
                    Else
                        For c% = 0 To 2
                            .Forecolor = RGB(220 - 20 * c%, 220 - 20 * c%, 220 - 20 * c%)
                            pnts(0).x = .Width - 7 + c%
                            pnts(0).y = (.Height - H) / 2
                            pnts(1).x = .Width - 7 + c%
                            pnts(1).y = (.Height + H) / 2
                            Polyline .hDC, pnts(0), 2
                            pnts(0).x = (.Width - W) / 2
                            pnts(0).y = (.Height + H) / 2 - 3 + c%
                            pnts(1).x = (.Width + W) / 2
                            pnts(1).y = (.Height + H) / 2 - 3 + c%
                            Polyline .hDC, pnts(0), 2
                        Next c%
                    End If
                End If
                ImgList1.ListImages.Add i%, f1.Path, .Image
            End With
            Set ximg = Nothing
        End If
        pb1.Value = j
        Image1.Width = pb1.Value / pb1.Max * (PicTop.ScaleWidth - (Image1.Left * 1.3))
    Next
    Image1.Visible = False
    Set F = Nothing
    Set fj = Nothing
    Set f1 = Nothing
    Me.scaleMode = 1
End If
Set fso = Nothing
End Sub
