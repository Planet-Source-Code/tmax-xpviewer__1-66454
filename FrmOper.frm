VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOper 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      Picture         =   "FrmOper.frx":0000
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   484
      TabIndex        =   17
      Top             =   0
      Width           =   7260
      Begin VB.Image Image2 
         Height          =   315
         Index           =   1
         Left            =   480
         Picture         =   "FrmOper.frx":4B8F
         Top             =   -15000
         Width           =   345
      End
      Begin VB.Image Image3 
         Height          =   315
         Index           =   1
         Left            =   0
         Picture         =   "FrmOper.frx":7283
         Top             =   -15000
         Width           =   345
      End
      Begin VB.Image Image1 
         Height          =   315
         Index           =   1
         Left            =   840
         Picture         =   "FrmOper.frx":9887
         Top             =   -15000
         Width           =   345
      End
      Begin VB.Image Image2 
         Height          =   315
         Index           =   2
         Left            =   480
         Picture         =   "FrmOper.frx":BFC4
         Top             =   -15000
         Width           =   345
      End
      Begin VB.Image Image3 
         Height          =   315
         Index           =   2
         Left            =   0
         Picture         =   "FrmOper.frx":E63C
         Top             =   -15000
         Width           =   345
      End
      Begin VB.Image Image1 
         Height          =   315
         Index           =   2
         Left            =   840
         Picture         =   "FrmOper.frx":10C96
         Top             =   -15000
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000CCFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   6975
      Begin ProXPViewer.TMcmdbutton TMcmdFolder 
         Height          =   375
         Left            =   6120
         TabIndex        =   16
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "..."
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
      Begin VB.ListBox LstProcess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         ItemData        =   "FrmOper.frx":13328
         Left            =   240
         List            =   "FrmOper.frx":1332A
         TabIndex        =   15
         Top             =   1200
         Width           =   5535
      End
      Begin ProXPViewer.TMcmdbutton TMcmdSelect 
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "Select"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin ProXPViewer.TMcmdbutton TMcmdRename 
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "Rename"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin VB.TextBox TxtStart 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Text            =   "1"
         Top             =   840
         Width           =   375
      End
      Begin VB.ComboBox cboPrefix 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cboPercent 
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   -10000
         Width           =   1095
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Top             =   -10000
         Width           =   1095
      End
      Begin VB.ComboBox cboDestFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
      Begin ProXPViewer.TMcmdbutton TMcmdCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin ProXPViewer.TMcmdbutton TMcmdOk 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "Ok"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView LvRename 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ExifDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LvSelect 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Origin"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FrmOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DestFolder$

Private Sub cboDestFolder_Click()
DestFolder = cboDestFolder.Text
DestFolder = DestFolder + IIf(Right(DestFolder, 1) <> "\", "\", "")
End Sub

Private Sub cboDestFolder_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i%
Dim Gotit As Boolean
On Error GoTo DestfolderErr
Gotit = False
If KeyCode = 13 Then
    If Dir(cboDestFolder.Text, vbDirectory) = "" Then MkDir cboDestFolder.Text
    cboDestFolder.Text = cboDestFolder.Text + IIf(Right(cboDestFolder.Text, 1) <> "\", "\", "")
    DestFolder = cboDestFolder.Text
    For i% = 0 To cboDestFolder.ListCount - 1
        If cboDestFolder.List(i%) = cboDestFolder.Text Then Gotit = True
    Next i
    If Not Gotit Then cboDestFolder.AddItem cboDestFolder.Text: cboDestFolder.ListIndex = cboDestFolder.ListCount - 1
End If
Exit Sub
DestfolderErr:
MsgBox cboDestFolder.Text & vbCrLf & Err.description
Resume Next
End Sub

Private Sub cboPrefix_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i%
Dim Gotit As Boolean
Gotit = False
If KeyCode = 13 Then
    If Right(cboPrefix.Text, 1) <> "#" Then Exit Sub
    For i% = 0 To cboPrefix.ListCount - 1
        If cboPrefix.List(i%) = cboPrefix.Text Then Gotit = True
    Next i
    If Not Gotit Then cboPrefix.AddItem cboPrefix.Text: cboPrefix.ListIndex = cboPrefix.ListCount - 1
    End If
End Sub

Sub Convert()
If InStr(1, cboPrefix.Text, "Time Stamp") > 0 Then
    TimeStamp
Else
    If InStr(1, cboPrefix.Text, "Original") > 0 Then
        CopyOriginal
    Else
        'MoveSelectedFile
        ConvertFName
    End If
End If
End Sub

Sub ConvertFName()
Dim i%, j%
Dim PrefixLen As Integer
Dim itmx As ListItem
Dim formatLen As String
Dim Start As Integer
If LvRename.ListItems.Count = 0 Then
    MsgBox "No file select!"
    Exit Sub
End If
If InStr(1, cboPrefix.Text, "#") = 0 Then
    MsgBox "Prefix must be :'Img#' or 'image###'"
    Exit Sub
End If
If Val(TxtStart.Text) < 0 Then
    MsgBox "Start must be more than '0'"
    TxtStart.SelStart = 0
    TxtStart.SelLength = Len(TxtStart)
    TxtStart.SetFocus
    Exit Sub
End If
PrefixLen = Len(Mid(cboPrefix.Text, InStr(1, cboPrefix.Text, "#")))
Dim PreFix$
PreFix$ = Mid(cboPrefix.Text, 1, Len(cboPrefix.Text) - PrefixLen)
formatLen = "0"
For i% = 1 To PrefixLen - 1
    formatLen = formatLen & "0"
Next i%
LvSelect.ListItems.Clear
With LvRename
    For i% = 1 To .ListItems.Count
        j% = i% + Val(TxtStart.Text) - 1
        Set itmx = LvSelect.ListItems.Add(, DestFolder & PreFix & Format$(Trim(j%), formatLen) & ".JPG", PreFix & Format$(Trim(j%), formatLen) & ".JPG")
        itmx.SubItems(1) = .ListItems(i%).Key
    Next
End With
End Sub

Sub CopyOriginal()
Dim i%
Dim itmx As ListItem
If LvRename.ListItems.Count = 0 Then
    MsgBox "No file select!"
    Exit Sub
End If
LvSelect.ListItems.Clear
With LvRename
    For i% = 1 To .ListItems.Count
        Set itmx = LvSelect.ListItems.Add(, DestFolder & .ListItems(i).Text, .ListItems(i).Text)
        itmx.SubItems(1) = .ListItems(i%).Key
    Next
End With
End Sub

Sub TimeStamp()
Dim i%
Dim itmx As ListItem
Dim MsgErr As String
Dim ErrLoad As Boolean
On Error GoTo TErr
ErrLoad = False
If LvRename.ListItems.Count = 0 Then
    MsgBox "No file select!"
    Exit Sub
End If
With LvRename
    For i% = 1 To .ListItems.Count
        Set itmx = LvSelect.ListItems.Add(, DestFolder & .ListItems(i).Text, .ListItems(i).SubItems(2) & ".JPG")
        itmx.SubItems(1) = .ListItems(i%).Key
    Next
End With
Set itmx = Nothing
If ErrLoad Then MsgBox "File :" & vbCrLf & MsgErr & vbCrLf & "Error TimeStamp", vbOKOnly, "Time Stamping"
Exit Sub
TErr:
MsgErr = MsgErr & "PreFix" & vbCrLf
ErrLoad = True
Resume Next
End Sub

Private Sub Form_Load()
LoadAllValue
End Sub

Private Sub Image3_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then Image3(3).Picture = Image3(1).Picture
End Sub

Private Sub Image3_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then Image3(3).Picture = Image3(0).Picture
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub TMcmdFolder_Click()
Dim Folder As String
Folder = BrowseForFolderDlg(cboDestFolder.Text, "Select Source Folder", Me.hwnd)
If Folder <> "" Then
    cboDestFolder.Text = Folder
End If
'tempfolder = cboDestFolder.Text
cboDestFolder_KeyUp 13, 1
End Sub

Private Sub TMcmdOK_Click()
SaveAllValue
Unload Me
End Sub

Sub LoadAllValue()
Dim DestCount%, PreCount%, i%
DestCount = Val(QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "DestCount"))
For i% = 1 To DestCount
    cboDestFolder.AddItem QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "Dest " & Trim$(i%))
Next i%
PreCount = Val(QueryValue(HKEY_CURRENT_USER, "XPViewer\Prefix", "PreCount"))
For i% = 1 To PreCount
    cboPrefix.AddItem QueryValue(HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre " & Trim$(i%))
Next
cboPrefix.ListIndex = 0
End Sub

Sub SaveAllValue()
Dim i%
Dim ret
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "PreCount", cboPrefix.ListCount, REG_SZ
For i% = 1 To cboPrefix.ListCount
    SetKeyValue HKEY_CURRENT_USER, "XPViewer\Prefix", "Pre " & Trim$(i%), cboPrefix.List(i - 1), REG_SZ
Next i%
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "DestCount", cboDestFolder.ListCount, REG_SZ
For i% = 1 To cboDestFolder.ListCount
    SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder\DestFolder", "Dest " & Trim$(i%), cboDestFolder.List(i% - 1), REG_SZ
Next i%
End Sub

Private Sub TMcmdCancel_Click()
Unload Me
End Sub

Private Sub TMcmdRename_Click()
On Error GoTo FCopyErr
Dim i%
With LvSelect
    For i% = 1 To .ListItems.Count
        FileCopy .ListItems(i%).SubItems(1), DestFolder & .ListItems(i%).Text
        AddMsg "Copy " & .ListItems(i%).Text & "..."
    Next
End With
AddMsg "Done."
Exit Sub
FCopyErr:
If Err.Number = 70 Then ShowMsg "Source and Destination filepath are the same!", vbOKOnly, "Rename Error"
Resume Next
End Sub

Private Sub TMcmdSelect_Click()
Convert
End Sub

Sub MoveSelectedFile()
Dim formatLen As String
Dim PreFix$, PrefixLen%
Dim i%, j%
Dim itmx As ListItem
PrefixLen = Len(Mid(cboPrefix.Text, InStr(1, cboPrefix.Text, "#")))
PreFix$ = Mid(cboPrefix.Text, 1, Len(cboPrefix.Text) - PrefixLen)
formatLen = "0"
For i% = 1 To PrefixLen - 1
    formatLen = formatLen & "0"
Next i%
With LvRename
    For i% = 1 To .ListItems.Count
        If .ListItems(i%).Selected = True Then
            j% = i% + Val(TxtStart.Text) - 1
            Set itmx = LvSelect.ListItems.Add(, DestFolder & PreFix & Format$(Trim(i%), formatLen) & ".JPG", PreFix & Format$(Trim(j%), formatLen) & ".JPG")
            itmx.SubItems(1) = .ListItems(i%).Key
        End If
    Next
End With
End Sub

Sub AddMsg(Msg As String)
LstProcess.AddItem Msg
LstProcess.ListIndex = LstProcess.ListCount - 1
End Sub
