VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9630
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPreviews.frx":0000
   ScaleHeight     =   642
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   5325
      ScaleHeight     =   1200
      ScaleWidth      =   1845
      TabIndex        =   9
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6435
      Top             =   1350
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   740
      TabIndex        =   0
      Top             =   0
      Width           =   11100
      Begin ProXPViewer.TMcmdbutton TMcmdbutton2 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "x"
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
      Begin MSComctlLib.ListView Lv1 
         Height          =   9615
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   16960
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "ImgList1"
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "|<"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "<<"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   ">>"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   4
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   ">|"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Index           =   4
         Left            =   2610
         TabIndex        =   8
         Top             =   60
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   661
         Enabled3D       =   0   'False
         XpButton        =   1
         Caption         =   "Slide Show"
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
      Begin VB.Label LblFilename 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3810
         TabIndex        =   6
         Top             =   135
         Width           =   60
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   0
         Picture         =   "FrmPreviews.frx":601C0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15450
      End
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CurrentPic As Integer

Sub LoadLv()
Dim i%
Lv1.ListItems.Clear
With FrmMdi.ActiveForm
    For i% = 1 To .ImgList1.ListImages.Count
        Lv1.ListItems.Add i%, .ImgList1.ListImages(i%).Key
    Next i%
End With
End Sub

Sub ChkButton()
TMcmdbutton1(0).Enabled = True
TMcmdbutton1(1).Enabled = True
TMcmdbutton1(2).Enabled = True
TMcmdbutton1(3).Enabled = True
If CurrentPic = 1 Then TMcmdbutton1(0).Enabled = False
If CurrentPic = Lv1.ListItems.Count Then TMcmdbutton1(3).Enabled = False
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Timer1.Interval = Val(SlideTimer)
End Sub

Private Sub Timer1_Timer()
TMcmdbutton1_Click (2)
If CurrentPic = Lv1.ListItems.Count Then TMcmdbutton1_Click (0)
End Sub

Private Sub TMcmdbutton2_Click()
Timer1.Enabled = False
Unload Me
End Sub

Public Sub LoadPV(FileName As String)
Dim ximg As cIMAGE
Dim rc As RECT
Dim PicLeft%, PicTop%
Set ximg = New cIMAGE
ximg.Load FileName
LblFilename = FileName
If ximg.ImageHeight < ximg.ImageWidth Then
    ximg.ReSize (ScaleWidth) * 15 / 16, 0, False
Else
    ximg.ReSize 0, (ScaleHeight - Picture1.ScaleHeight) * 15 / 16, False
End If
PicTop = (ScaleHeight - ximg.ImageHeight + Picture1.ScaleHeight) / 2
PicLeft = (ScaleWidth - ximg.ImageWidth) / 2
SetRect rc, PicLeft - 6, PicTop - 6, ximg.ImageWidth + 6, ximg.ImageHeight + 6
Me.Cls
ximg.PaintDC Me.hDC, PicLeft, PicTop
makeEdge rc
Set ximg = Nothing
End Sub

Private Sub TMcmdbutton1_Click(Index As Integer)
Select Case Index
    Case 0: CurrentPic = 1
    Case 1: If CurrentPic > 1 Then CurrentPic = CurrentPic - 1
    Case 2: If CurrentPic < Lv1.ListItems.Count Then CurrentPic = CurrentPic + 1
    Case 3: CurrentPic = Lv1.ListItems.Count
    Case 4: SlideShow
End Select
LoadPV Lv1.ListItems(CurrentPic).Key
ChkButton
End Sub

Sub SlideShow()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub makeEdge(rc As RECT)
Dim i%
Dim pnts(2) As POINTAPI
For i% = 6 To 1 Step -1
    Me.Forecolor = RGB(165 - i% * 15, 165 - i% * 15, 165 - i% * 15)
    'bottom
    pnts(0).x = rc.Left
    pnts(0).y = rc.Top + rc.Bottom - i% + 1
    pnts(1).x = rc.Left + rc.Right
    pnts(1).y = rc.Top + rc.Bottom - i% + 1
    Polyline Me.hDC, pnts(0), 2
    'right
    pnts(0).x = rc.Left + rc.Right - i%
    pnts(0).y = rc.Top + rc.Bottom
    pnts(1).x = rc.Left + rc.Right - i%
    pnts(1).y = rc.Top
    Polyline Me.hDC, pnts(0), 2
    Me.Forecolor = RGB(245 - i% * 15, 245 - i% * 15, 245 - i% * 15)
    'left
    pnts(0).x = rc.Left + i% - 1
    pnts(0).y = rc.Top
    pnts(1).x = rc.Left + i% - 1
    pnts(1).y = rc.Top + rc.Bottom
    Polyline Me.hDC, pnts(0), 2
    'top
    pnts(0).x = rc.Left
    pnts(0).y = rc.Top + i% - 1
    pnts(1).x = rc.Left + rc.Right
    pnts(1).y = rc.Top + i% - 1
    Polyline Me.hDC, pnts(0), 2
Next i%
End Sub

