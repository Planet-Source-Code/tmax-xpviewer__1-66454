VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMdi 
   BackColor       =   &H8000000C&
   Caption         =   "XPress Viewer"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   315
   ClientWidth     =   14730
   Icon            =   "FrmMdis.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FrmMdis.frx":0CCE
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10260
      Left            =   0
      ScaleHeight     =   684
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   9
      Top             =   480
      Width           =   5175
      Begin VB.PictureBox PicClosed 
         BorderStyle     =   0  'None
         Height          =   10215
         Left            =   5040
         ScaleHeight     =   10215
         ScaleWidth      =   135
         TabIndex        =   31
         Top             =   0
         Width           =   135
         Begin VB.Image ImgClosed 
            Height          =   10200
            Left            =   0
            Picture         =   "FrmMdis.frx":57A26
            Stretch         =   -1  'True
            Top             =   0
            Width           =   120
         End
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   5070
      End
      Begin VB.PictureBox Pic5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   337
         TabIndex        =   15
         Top             =   3480
         Width           =   5055
         Begin VB.TextBox TxtNewFolder 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   2895
         End
         Begin ProXPViewer.TMcmdbutton TMcmdOption 
            Height          =   375
            Left            =   4200
            TabIndex        =   16
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            XpButton        =   1
            Caption         =   "Option"
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
         Begin ProXPViewer.TMcmdbutton TMcmdNewFolder 
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            XpButton        =   1
            Caption         =   "New Folder"
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
         Begin VB.Image Image2 
            Height          =   495
            Left            =   0
            Picture         =   "FrmMdis.frx":5C09F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5040
         End
      End
      Begin VB.PictureBox PicExif 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   0
         ScaleHeight     =   4095
         ScaleWidth      =   5055
         TabIndex        =   11
         Top             =   8520
         Width           =   5055
         Begin VB.Label LblExifdata 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1400
            Left            =   1440
            TabIndex        =   21
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label LblExifTitle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1400
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   1215
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00E0E0E0&
            Height          =   375
            Index           =   11
            Left            =   1560
            Top             =   4800
            Width           =   3300
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00E0E0E0&
            Height          =   375
            Index           =   11
            Left            =   120
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label LblData 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ExifData"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1680
            TabIndex        =   13
            Top             =   4920
            Width           =   600
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exposure Time"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   12
            Top             =   4920
            Width           =   1050
         End
         Begin VB.Image Image1 
            Height          =   1605
            Left            =   -240
            Picture         =   "FrmMdis.frx":60C2E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5280
         End
      End
      Begin VB.PictureBox PicPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         Height          =   4575
         Left            =   0
         ScaleHeight     =   301
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   10
         Top             =   3960
         Width           =   5055
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   4515
            TabIndex        =   14
            Top             =   1.50000e5
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.Image ImgPreview 
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Left            =   960
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Image Image4 
            Height          =   4560
            Left            =   0
            Picture         =   "FrmMdis.frx":652A7
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5040
         End
      End
      Begin VB.DirListBox Dir1 
         Height          =   3240
         Left            =   0
         TabIndex        =   30
         Top             =   300
         Width           =   5055
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   0
         Picture         =   "FrmMdis.frx":69920
         Stretch         =   -1  'True
         Top             =   -1.50000e5
         Width           =   5040
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "FrmMdis.frx":6E4AF
      ScaleHeight     =   480
      ScaleWidth      =   14730
      TabIndex        =   0
      Top             =   0
      Width           =   14730
      Begin ProXPViewer.TMcmdbutton TMcmdIndPrint 
         Height          =   375
         Left            =   -10000
         TabIndex        =   7
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "Index Print"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton6 
         Height          =   375
         Left            =   6435
         TabIndex        =   6
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "WC"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton5 
         Height          =   375
         Left            =   5835
         TabIndex        =   5
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "WH"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton4 
         Height          =   375
         Left            =   5235
         TabIndex        =   4
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "WV"
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
      Begin VB.ComboBox cbPath 
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
         IntegralHeight  =   0   'False
         Left            =   10080
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
      Begin ProXPViewer.TMcmdbutton TMcmdbutton8 
         Height          =   375
         Left            =   7035
         TabIndex        =   8
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "view"
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -240
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   300
         ImageHeight     =   57
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMdis.frx":7303E
               Key             =   "Up"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMdis.frx":735E4
               Key             =   "Over"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMdis.frx":73BD7
               Key             =   "Down"
            EndProperty
         EndProperty
      End
      Begin ProXPViewer.TMcmdbutton TMcmdRefresh 
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   45
         Width           =   975
         _ExtentX        =   1296
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "Refresh"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton9 
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "A:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ProXPViewer.TMcmdbutton TMcmdbutton3 
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "E:"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton2 
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "D:"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton1 
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "C:"
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
      Begin ProXPViewer.TMcmdbutton TMcmdDesktop 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   45
         Width           =   855
         _ExtentX        =   1720
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "Desktop"
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
      Begin ProXPViewer.TMcmdbutton TMcmdbutton7 
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "Favorites"
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
      Begin ProXPViewer.TMcmdbutton TMcmdAbout 
         Height          =   375
         Left            =   7995
         TabIndex        =   32
         Top             =   45
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "About"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRC32"
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
         Left            =   -10000
         TabIndex        =   19
         Top             =   120
         Width           =   600
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   -500
         Width           =   1500
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   0
         Picture         =   "FrmMdis.frx":7418E
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   19320
      End
      Begin VB.Label LblAdress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   570
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   10740
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12621
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "2006-9-4"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            TextSave        =   "19:51"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuListView 
      Caption         =   "ListView"
      Visible         =   0   'False
      Begin VB.Menu MnuRotate90 
         Caption         =   "Rotate 90"
      End
      Begin VB.Menu MnuRotate180 
         Caption         =   "Rotate 180"
      End
      Begin VB.Menu MnuRotate270 
         Caption         =   "Rotate 270"
      End
      Begin VB.Menu MnuSpace0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFlipVertical 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu MnuFlipHorizontal 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu Mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu MnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu MnuExifDate 
         Caption         =   "Exif Date"
      End
      Begin VB.Menu MnuAddExifDate 
         Caption         =   "Add ExifDate"
      End
      Begin VB.Menu MnuSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopyTo 
         Caption         =   "Copy To"
      End
      Begin VB.Menu MnuMoveTo 
         Caption         =   "Move To"
      End
      Begin VB.Menu MNuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSmartMove 
         Caption         =   "Smart Move"
      End
      Begin VB.Menu MnuSendTo 
         Caption         =   "Send To"
         Begin VB.Menu MnuADrive 
            Caption         =   "A Drive"
         End
         Begin VB.Menu MnuSendtoFolder 
            Caption         =   "Folder"
         End
      End
      Begin VB.Menu MnuResize 
         Caption         =   "Resize"
      End
      Begin VB.Menu MnuSpace7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpenWith 
         Caption         =   "Open With"
         Begin VB.Menu MnuPaint 
            Caption         =   "EZPaint"
         End
         Begin VB.Menu MnuPlugIn 
            Caption         =   "PlugIn"
            Begin VB.Menu MnuPlugIn1 
               Caption         =   "PlugIn 1"
               Index           =   0
            End
         End
         Begin VB.Menu MnuBurn 
            Caption         =   "Burn"
         End
      End
      Begin VB.Menu MnuDeskShortCut 
         Caption         =   "Desktop ShortCut"
      End
      Begin VB.Menu MnuSetWallPaper 
         Caption         =   "Set Wallpaper"
         Begin VB.Menu MnuWallStretch 
            Caption         =   "Stretch"
         End
         Begin VB.Menu MnuWallTile 
            Caption         =   "Tile"
         End
         Begin VB.Menu MnuWallCenter 
            Caption         =   "Center"
         End
      End
      Begin VB.Menu MnuBatchRename 
         Caption         =   "Batch Rename"
      End
      Begin VB.Menu MnuProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Visible         =   0   'False
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu MnuFolder 
      Caption         =   "Folder"
      Visible         =   0   'False
      Begin VB.Menu MnuFCreate 
         Caption         =   "Create"
      End
      Begin VB.Menu MnuFDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuFRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu MnuFRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "FrmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum ImgOperation
    Rotate_90 = 1
    Rotate_180 = 2
    Rotate_270 = 3
    Flip_Vertical = 4
    Flip_Horizontal = 5
    Exif_Date = 6
    moveto = 7
    SmartMove = 8
    copyto = 9
    AddExifDate = 10
End Enum

Dim ShowDelPic As Boolean
Const DT_BOTTOM = &H8
Const DT_LEFT = &H0
Const DT_RIGHT = &H2
Dim FontAlign As Long
Dim Software$, ModiDate$
Dim CFile$, cFileName$

Private Sub cbPath_Click()
GoPaths cbPath.Text
End Sub

Private Sub Dir1_Change()
Dim NewPath$
NewPath = Dir1.Path + IIf(Right(Dir1.Path, 1) <> "\", "\", "")
Dir1.Refresh
LoadNewWin NewPath
FilePath = NewPath + IIf(Right(NewPath, 1) <> "\", "\", "")
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err Then
    Drive1.Drive = Dir1.Path
End If

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MnuFolder
End Sub

Private Sub Image6_Click()
Static T As Boolean
If T Then
    Image6.Picture = ImageList1.ListImages(1).Picture
    ChkCRC = True
Else
    Image6.Picture = ImageList1.ListImages(3).Picture
    ChkCRC = False
End If
T = Not T
End Sub

Private Sub ImgClosed_Click()
If PicLeft.Tag = "closed" Then
PicLeft.Width = 5175
PicLeft.Tag = "Open"
PicClosed.Left = 336
'PicClosed.ToolTipText = "Clicked to closed"
ImgClosed.ToolTipText = "Clicked to closed"
Else
PicLeft.Width = 120
PicLeft.Tag = "closed"
PicClosed.Left = 0
'PicClosed.ToolTipText = "Clicked to open"
ImgClosed.ToolTipText = "Clicked to open"
End If
End Sub

Private Sub MDIForm_Load()
Dim i%, PluginCount%
''On Error Resume Next
    Check1stTime
    SearchJpgType
    LoadPlugIns
    LoadPrintDate
    LoadFolderReg
    LoadMiscReg
    LoadWallPaper
    GoPaths StartPath
    PicExif.Left = PicPreview.Left
    LoadTitle
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim retval As Boolean
retval = ShowMsg("Are You Sure ?", vbYesNo, "Exit Program")
If retval = True Then
    ClosedAll
Else
    Cancel = -1
End If
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
cbPath.Width = Picture1.ScaleWidth - cbPath.Left
End Sub

Private Sub MnuAddExifDate_Click()
Me.MousePointer = 11
Oper AddExifDate
Me.MousePointer = 0
End Sub

Private Sub MnuADrive_Click()
Me.MousePointer = 11
SaveResize
Me.MousePointer = 0
End Sub

Private Sub MnuBatchRename_Click()
'Open BatchRename Form
Dim foper As FrmOper
Dim X1exif As cEXIF
Dim fs, F, S, FExt
Dim i%
Set foper = New FrmOper
Dim itmx As ListItem
Set itmx = Me.ActiveForm.Lv1.SelectedItem
Me.MousePointer = 11
Set fs = CreateObject("Scripting.FileSystemObject")
For Each itmx In Me.ActiveForm.Lv1.ListItems
    If itmx.Selected Then
      Set F = fs.GetFile(itmx.Key)
        i% = i% + 1
        SetAttr itmx.Key, vbArchive
        foper.LvRename.ListItems.Add i%, itmx.Key, itmx.Text ', itmx.Icon
        foper.LvRename.ListItems(i%).SubItems(1) = F.ParentFolder
        Set X1exif = New cEXIF
            If X1exif.Load(F.Path) = True Then
                FExt = X1exif.EXIFmodified
                If X1exif.EXIFmodified = "12:00:00 AM" Then
                    foper.LvRename.ListItems(i%).SubItems(2) = Format$(F.DateLastModified, "yyyy-mm-dd hh-nn-ss")
                Else
                    foper.LvRename.ListItems(i%).SubItems(2) = Format$(FExt, "yyyy-mm-dd hh-nn-ss")
                End If
            End If
            Set X1exif = Nothing
    End If
Next
Me.MousePointer = 0
foper.Show 1
Set foper = Nothing
Set F = Nothing
Set fs = Nothing
End Sub

Private Sub MnuBurn_Click()
 ShellExecute Me.hwnd, "open", BurnSoftware, " /w", vbNullString, SW_SHOWNORMAL
End Sub

Private Sub mnuCopy_Click()
CFile = Me.ActiveForm.Lv1.SelectedItem.Key
cFileName = Me.ActiveForm.Lv1.SelectedItem.Text
End Sub

Private Sub MnuCopyTo_Click()
FilePath = Me.ActiveForm.Caption
LastPathSelect = BrowseForFolderDlg(FilePath, "Select a folder", Me.hwnd)
LastPathSelect = LastPathSelect + IIf(Right(LastPathSelect, 1) <> "\", "\", "")
Oper copyto
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "LastPathSelect", LastPathSelect, REG_SZ
End Sub

Private Sub MnuCut_Click()
    ShowPreview
End Sub

Private Sub mnuDelete_Click()
    ShowDelPic = True
    OperDelete
End Sub

Private Sub MnuDeskShortCut_Click()
Me.MousePointer = 11
fCreateShellLink "..\..\Desktop", Me.ActiveForm.Lv1.SelectedItem.Text, Me.ActiveForm.Lv1.SelectedItem.Key, ""
Me.MousePointer = 0
End Sub

Private Sub MnuExifDate_Click()
Oper Exif_Date
End Sub

Private Sub MnuFlipHorizontal_Click()
Oper Flip_Horizontal
End Sub

Private Sub MnuFlipVertical_Click()
Oper Flip_Vertical
End Sub

Private Sub MnuFRefresh_Click()
GoPaths FilePath
End Sub

Private Sub MnuKill_Click()
ShowDelPic = False
OperDelete
End Sub

Private Sub MnuMoveTo_Click()
FilePath = Me.ActiveForm.Caption
LastPathSelect = BrowseForFolderDlg(FilePath, "Select a folder", Me.hwnd)
LastPathSelect = LastPathSelect + IIf(Right(LastPathSelect, 1) <> "\", "\", "")
Oper moveto
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "LastPathSelect", LastPathSelect, REG_SZ
End Sub

Private Sub MnuPaste_Click()
Me.MousePointer = 11
DiskOps CFile, Me.ActiveForm.Caption & cFileName, F_CopySmart, 1
MnuRefresh_Click
Me.MousePointer = 0
End Sub

Private Sub MnuPlugIn1_Click(Index As Integer)
ShellExecute Me.hwnd, "open", PlugInSoftware(Index + 1), Chr$(34) & Me.ActiveForm.Lv1.SelectedItem.Key & Chr$(34), vbNullString, SW_SHOWNORMAL
End Sub

Private Sub MnuProperties_Click()
FileProperties Me.ActiveForm.Lv1.SelectedItem.Key
End Sub

Private Sub MnuRefresh_Click()
Me.ActiveForm.RefreshLv Me.ActiveForm.Caption
End Sub

Private Sub MnuResize_Click()
Me.MousePointer = 11
FilePath = Me.ActiveForm.Caption
ResizePic
GotoForm FilePath
Me.MousePointer = 0
End Sub

Private Sub MnuRotate180_Click()
Oper Rotate_180
End Sub

Private Sub MnuRotate270_Click()
Oper Rotate_270
End Sub

Private Sub MnuRotate90_Click()
Oper Rotate_90
End Sub

Private Sub MnuSendtoFolder_Click()
Me.MousePointer = 11
SaveResizeCD CDBurnPath
Me.MousePointer = 0
End Sub

Private Sub MnuSmartMove_Click()
Oper SmartMove
End Sub

Private Sub MnuWallCenter_Click()
If Me.ActiveForm.Lv1.SelectedItem.Key = "" Then Exit Sub
Me.MousePointer = 11
SetWallPaper Me.ActiveForm.Lv1.SelectedItem.Key, Center
Me.MousePointer = 0
End Sub

Private Sub MnuWallStretch_Click()
If Me.ActiveForm.Lv1.SelectedItem.Key = "" Then Exit Sub
Me.MousePointer = 11
SetWallPaper Me.ActiveForm.Lv1.SelectedItem.Key, Stretch
Me.MousePointer = 0
End Sub

Private Sub MnuWallTile_Click()
If Me.ActiveForm.Lv1.SelectedItem.Key = "" Then Exit Sub
Me.MousePointer = 11
SetWallPaper Me.ActiveForm.Lv1.SelectedItem.Key, Tile
Me.MousePointer = 0
End Sub

Private Sub TMcmdAbout_Click()
frmAbout.Show 1
End Sub

Private Sub TMcmdbutton1_Click()
GoPaths TMcmdbutton1.Tag
End Sub

Private Sub TMcmdRefresh_Click()
Drive1.Refresh
GoPaths FilePath
End Sub

Private Sub TMcmdbutton2_Click()
GoPaths TMcmdbutton2.Tag
End Sub

Private Sub TMcmdbutton3_Click()
GoPaths TMcmdbutton3.Tag
End Sub

Private Sub TMcmdbutton4_Click()
Me.Arrange vbTileVertical
End Sub

Private Sub TMcmdbutton5_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub TMcmdbutton6_Click()
 Me.Arrange vbCascade
End Sub

Private Sub TMcmdbutton7_Click()
GoPaths FavoritePath '   FolderLocation(CSIDL_PERSONAL) ' My Document
End Sub

Private Sub TMcmdbutton8_Click()
 On Error Resume Next
If Me.ActiveForm.Lv1.View = 3 Then
    Me.ActiveForm.Lv1.View = 0 'icon
Else
    Me.ActiveForm.Lv1.View = 3
End If
End Sub

Private Sub TMcmdbutton9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then FormatDrive "a:"  ' format a:
If Button = 1 Then GoPaths "a:"
End Sub

Private Sub TMcmdDesktop_Click()
GoPaths FolderLocation(CSIDL_Desktop)
End Sub

Private Sub TMcmdNewFolder_Click()
On Error Resume Next
Dim ret As Boolean
If Len(TxtNewFolder.Text) <> 0 Then
    If Dir(FilePath & TxtNewFolder.Text, vbDirectory) = "" Then
        ret = ShowMsg("Create Folder :" & TxtNewFolder.Text, vbYesNo, "Create Folder")
        If ret Then
            MkDir FilePath & TxtNewFolder.Text
            GoPaths FilePath
        End If
    Else
        ShowMsg "Folder :" & TxtNewFolder.Text & vbCrLf & "Already Exist!", vbOKOnly, "Create Folder"
    End If
End If
End Sub

Private Sub TMcmdOption_Click()
Dim fOption As frmOptions
Set fOption = New frmOptions
With fOption
    .TxtFolder(0).Text = FavoritePath
    .TxtFolder(1).Text = CDBurnPath
    .TxtFolder(2).Text = StartPath
    .TxtFolder(3).Text = SmartPath
    .TxtThumbnailSize.Text = ThumbnailSize
    .TxtSlideTimer.Text = SlideTimer
    .Show 1
End With
Set fOption = Nothing
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Sub LoadNewWin(Path As String)
Dim iForm As Form
Dim fLv As FrmLv
For Each iForm In Forms
    If iForm.Caption = Path Then
        iForm.PicExif.Forecolor = DFontColor
        iForm.PicExif.Font = DFontName
        iForm.PicExif.FontSize = DFontSize
        iForm.SetFocus
        Exit Sub
    End If
Next
Set fLv = New FrmLv
fLv.Caption = Path
fLv.PicTop.Refresh
Screen.MousePointer = 11
sbStatusBar.Panels(1) = "Loading ..."
fLv.RefreshLv Path
Screen.MousePointer = 0
'On Error Resume Next
If fLv.Lv1.ListItems.Count > 0 Then
    If bFileExists(Path & "desktop.ini") And bFileExists(Path & "bg") Then
      fLv.Lv1.Picture = LoadPicture(Path & "bg")
    End If
    fLv.Show
    AddPath2Cb Path
Else
    Unload fLv
End If
Set fLv = Nothing
sbStatusBar.Panels(1) = ""
End Sub

Sub ClosedAll()
Dim iForm As Form
For Each iForm In Forms
    Unload iForm
Next
End Sub

Public Sub LoadPreview(FileName As String)
Dim ximg As cIMAGE
Dim PicRatio, xRatio As Single
On Error GoTo PreviewErr
Set ximg = New cIMAGE
Screen.MousePointer = 11
sbStatusBar.Panels(1) = "Loading ..."
With PicPreview
    PicRatio = .Width / .Height
    ximg.Load FileName
    xRatio = ximg.ImageWidth / ximg.ImageHeight
    If ximg.ImageHeight < ximg.ImageWidth Then
            ximg.ReSize .Width * 15 / 16, 0, False
    Else
            ximg.ReSize 0, .Height * 15 / 16, False
    End If
    ImgPreview.Visible = False
    ImgPreview.Left = (.ScaleWidth - ximg.ImageWidth) / 2
    ImgPreview.Top = (.ScaleHeight - ximg.ImageHeight) / 2
    ImgPreview.Width = ximg.ImageWidth
    ImgPreview.Height = ximg.ImageHeight
    ImgPreview.Picture = ximg.Picture
    ImgPreview.Visible = True
    .Tag = FileName
End With
Set ximg = Nothing
sbStatusBar.Panels(1) = Me.ActiveForm.Lv1.SelectedItem.Text
Screen.MousePointer = 0
Exit Sub
PreviewErr:
 Set ximg = Nothing
 sbStatusBar.Panels(1) = ""
Screen.MousePointer = 0
Resume Next
End Sub

Sub SaveResize()
Dim ximg As cIMAGE
Dim TempFile As String
Dim retval As Long
On Error GoTo SaveResizeErr
Set ximg = New cIMAGE
If DriveAReady Then
    ShowMsg "Drive A: not ready!", vbOKOnly, "Drive Error"
    Exit Sub
End If
If Val(Me.ActiveForm.Lv1.SelectedItem.Tag) > 1280000 Then
    ShowMsg "File :" & Me.ActiveForm.Lv1.SelectedItem.Text & vbCrLf & "File Size too big.", vbOKOnly, "File Size"
    Exit Sub
End If
FileSelect = Me.ActiveForm.Lv1.SelectedItem.Key
  TempFile = GetATemporaryFileName
  sbStatusBar.Panels(3).Text = "Compute background "
    If ximg.Load(Me.ActiveForm.Lv1.SelectedItem.Key) = True Then
        If ximg.ImageHeight < ximg.ImageWidth Then
            ximg.ReSize 800, 0, False
        Else
            ximg.ReSize 0, 600, False
        End If
        sbStatusBar.Panels(3).Text = "Compute desktop.ini "
        SaveDesktop "a:", "bg"
        sbStatusBar.Panels(3).Text = "Save desktop.ini"
        DrawCaption ximg, TempFile, "..."
        sbStatusBar.Panels(3).Text = "Save main file"
        DiskOps FileSelect, "a:\" & Me.ActiveForm.Lv1.SelectedItem.Text, F_CopySmart, 1
        sbStatusBar.Panels(3).Text = "Save background "
        DiskOps TempFile, "a:\bg", F_CopySmart, 1
        sbStatusBar.Panels(3).Text = "Copy background "
            SetAttr "a:\bg", vbHidden
            SetAttr "a:\desktop.ini", vbHidden
            sbStatusBar.Panels(3).Text = "Set attribute "
            DiskOps TempFile, TempFile, F_DelUndo, 1
            sbStatusBar.Panels(3).Text = "delete tempfile "
    End If
Set ximg = Nothing
sbStatusBar.Panels(3).Text = ""
Exit Sub
SaveResizeErr:
Set ximg = Nothing
Resume Next
End Sub

Sub SaveResizeCD(Dest As String)
Dim ximg As cIMAGE
Dim xexif As cEXIF
Dim Exifdate As String
Dim TempFile As String
Dim retval As Long
On Error GoTo SaveResizeCDErr
Dest = Dest + IIf(Right(Dest, 1) <> "\", "\", "")
If Dir(Dest, vbDirectory) = "" Then
    retval = ShowMsg("Folder : " & Dest & vbCrLf & "not found." & vbCrLf & "Create Folder?", vbYesNo, "SaveCD")
    If retval Then
        MkDir Dest
    Else
        Exit Sub
    End If
End If
Set ximg = New cIMAGE
Set xexif = New cEXIF
FileSelect = Me.ActiveForm.Lv1.SelectedItem.Key
If xexif.Load(FileSelect) Then
    Exifdate = Format$(xexif.EXIFmodified, "dd-mm-yyyy hh:nn:ss")
    Else
    Exifdate = ""
End If
Set xexif = Nothing
  TempFile = GetATemporaryFileName
  sbStatusBar.Panels(3).Text = "Compute background "
    If ximg.Load(Me.ActiveForm.Lv1.SelectedItem.Key) = True Then
        If ximg.ImageHeight < ximg.ImageWidth Then
            ximg.ReSize 800, 0, False
        Else
            ximg.ReSize 0, 600, False
        End If
        sbStatusBar.Panels(3).Text = "Compute desktop.ini "
        SaveDesktop Dest$, "bg"
        sbStatusBar.Panels(3).Text = "Save desktop.ini"
        DrawCaption ximg, TempFile, Exifdate
        sbStatusBar.Panels(3).Text = "Save main file"
        DiskOps FileSelect, Dest & Me.ActiveForm.Lv1.SelectedItem.Text, F_CopySmart, 1
        sbStatusBar.Panels(3).Text = "Save background "
        DiskOps TempFile, Dest & "bg", F_CopySmart, 1
        sbStatusBar.Panels(3).Text = "Copy background "
        SetAttr Dest & "bg", vbHidden
        SetAttr Dest & "desktop.ini", vbHidden
        sbStatusBar.Panels(3).Text = "Set attribute "
        DiskOps TempFile, TempFile, F_DelUndo, 1
        sbStatusBar.Panels(3).Text = "delete tempfile "
    End If
Set ximg = Nothing
sbStatusBar.Panels(3).Text = ""
Exit Sub
SaveResizeCDErr:
Set ximg = Nothing
Resume Next
End Sub

Public Sub SaveDesktop(Dest$, BgFile$)
Dim DeskStr$
Dim TempFile$
TempFile = GetATemporaryFileName
DeskStr = "[ExtShellFolderViews]" & vbCrLf
DeskStr = DeskStr & "{BE098140-A513-11D0-A3A4-00C04FD706EC}={BE098140-A513-11D0-A3A4-00C04FD706EC}" & vbCrLf
DeskStr = DeskStr & "[{BE098140-A513-11D0-A3A4-00C04FD706EC}]" & vbCrLf
DeskStr = DeskStr & "IconArea_Image =" & BgFile & vbCrLf
DeskStr = DeskStr & "IconArea_Text=0x00000000" & vbCrLf
DeskStr = DeskStr & "Attributes = 1" & vbCrLf
DeskStr = DeskStr & "[.ShellClassInfo]" & vbCrLf
DeskStr = DeskStr & "ConfirmFileOp = 0"
Dim OutStream As TextStream
Set OutStream = fsys.CreateTextFile(TempFile, True, False)
OutStream.WriteLine DeskStr
Set OutStream = Nothing
Dest$ = Dest$ + IIf(Right$(Dest$, 1) <> "\", "\", "")
DiskOps TempFile, Dest$ & "desktop.ini", F_CopySmart, 1
DiskOps TempFile, TempFile, F_DelUndo, 1
End Sub

Sub DrawCaption(ximg As cIMAGE, FileName As String, Caption$)
Dim fpreview As FrmPreview
Set fpreview = New FrmPreview
 Dim mRect As RECT
lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
With fpreview
        .PicPreview.Height = ximg.ImageHeight
        .PicPreview.Width = ximg.ImageWidth
        .PicPreview.Picture = ximg.Picture
        .PicPreview.FontName = "Courier New"
        .PicPreview.FontSize = 18
        .PicPreview.FontBold = True
        .PicPreview.DrawMode = 7
        .PicPreview.BackColor = vbBlue
        .PicPreview.Forecolor = vbBlack
        SetRect mRect, 1, 11, ximg.ImageWidth, 44
        DrawText .PicPreview.hDC, CompanyName, -1, mRect, lwFontAlign
        SetRect mRect, 1, ximg.ImageHeight - 41, ximg.ImageWidth, ximg.ImageHeight - 11
        DrawText .PicPreview.hDC, Caption$, -1, mRect, lwFontAlign
        .PicPreview.Forecolor = &HCCFF&       ' vbYellow
        SetRect mRect, 0, 10, ximg.ImageWidth, 44
        DrawText .PicPreview.hDC, CompanyName, -1, mRect, lwFontAlign
        SetRect mRect, 0, ximg.ImageHeight - 40, ximg.ImageWidth, ximg.ImageHeight - 10
        DrawText .PicPreview.hDC, Caption$, -1, mRect, lwFontAlign
        .PicPreview.Picture = .PicPreview.Image
        SaveJPG .PicPreview.Picture, FileName 'FilePath & "bg.jpg"
        End With
End Sub

Sub GoPaths(Path)
Dim chkdisk As Boolean
chkdisk = True
If Path = "a:" Then chkdisk = CheckDiskette
If chkdisk = False Then Exit Sub
Drive1.Drive = Left(Path, 1) & ":"
Dir1.Path = Path
End Sub

Function FileDate(FileName As String) As String
Dim fs, F
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.GetFile(FileName)
FileDate = F.DateLastModified
Set F = Nothing
Set fs = Nothing
End Function

Function GotoForm(Path As String)
Dim iForm As Form
For Each iForm In Forms
    If iForm.Caption = Path Then
        iForm.SetFocus
        sbStatusBar.Panels(1) = "Refreshing..."
        iForm.RefreshLv Path
        sbStatusBar.Panels(1) = ""
        Exit Function
    End If
Next
End Function

 Function DriveAReady() As Boolean
 Dim sPath   As String
  Dim uWFD    As WIN32_FIND_DATA
  Dim hSearch As Long
        sPath = "a:\*.*" & vbNullChar
        hSearch = FindFirstFile(sPath, uWFD)
        FindClose (hSearch)
        DriveAReady = hSearch
       ' Debug.Print hSearch
        DriveAReady = True
       'If (Not m_bEnding And Not ucFolderView.PathIsValid("a:\")) Then
      '  If (Not m_bEnding) Then
      '  DriveAReady = False
      '  End If
 End Function

Sub ResizePic()
Dim ximg As cIMAGE
Dim xexif As cEXIF
Dim TempFile As String
Dim retval As Long
On Error GoTo ResizeErr
Set ximg = New cIMAGE
FilePath = Me.ActiveForm.Caption
FileSelect = Me.ActiveForm.Lv1.SelectedItem.Key
Set xexif = New cEXIF
  TempFile = FilePath & "Resize_" & Me.ActiveForm.Lv1.SelectedItem.Text
    If ximg.Load(Me.ActiveForm.Lv1.SelectedItem.Key) = True Then
        sbStatusBar.Panels(3).Text = "Resize..."
        SaveJPG ximg.Picture, TempFile, 75
    End If
    Set ximg = Nothing
    xexif.Load TempFile
    xexif.EXIFmodified = Now & Chr$(13)
    xexif.EXIFsoftware = "XP Viewer ver 1.0 " & Chr$(13)
    xexif.Save
    Set xexif = Nothing
    DiskOps FileSelect, FileSelect, F_DelUndo, 1
    DiskOps TempFile, FileSelect, F_Rename, 1
    sbStatusBar.Panels(3).Text = ""
Exit Sub
ResizeErr:
Set ximg = Nothing
Resume Next
End Sub

Sub ShowPreview()
Dim fpreview As FrmPreview
Set fpreview = New FrmPreview
Dim ximg As cIMAGE
Dim retval As Long
Dim mRect As RECT
Set ximg = New cIMAGE
lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
FileSelect = Me.ActiveForm.Lv1.SelectedItem.Key
    If ximg.Load(Me.ActiveForm.Lv1.SelectedItem.Key) = True Then
        If ximg.ImageHeight < ximg.ImageWidth Then
            ximg.ReSize 800, 0, False
        Else
            ximg.ReSize 0, 600, False
        End If
          SetRect mRect, 1, 11, ximg.ImageWidth + 1, 45
         fpreview.PicPreview.Picture = ximg.Picture
         fpreview.PicPreview.Width = ximg.ImageWidth
         fpreview.PicPreview.Height = ximg.ImageHeight
         fpreview.PicPreview.Left = 10
         fpreview.PicPreview.Top = fpreview.Picture1.Height + 10
         DrawText fpreview.PicPreview.hDC, "    Summer Studio && Colour Photo    ", -1, mRect, lwFontAlign
         fpreview.PicPreview.Visible = True
        fpreview.Show 1
    End If
    Set ximg = Nothing
End Sub

Sub Oper(op As ImgOperation)
Dim xexif As cEXIF
Dim Loading As Boolean
Dim SelectFile As String
Dim SmartFolder As String
Dim NewFile As String
Dim itmx As ListItem
Dim ret As Boolean

On Error GoTo OperErr

FilePath = Me.ActiveForm.Caption
Set itmx = Me.ActiveForm.Lv1.SelectedItem
Me.MousePointer = 11
For Each itmx In Me.ActiveForm.Lv1.ListItems
    If itmx.Selected Then
        SetAttr itmx.Key, vbArchive
        Select Case op
            Case ImgOperation.Exif_Date:  'exit date
                sbStatusBar.Panels(1) = itmx.Text & "==> Exif Date..."
                Set xexif = New cEXIF
                Loading = xexif.Load(itmx.Key)
                If AddSoftwareName Then
                    xexif.EXIFsoftware = "XP Viewer"
                    xexif.Save
                End If
                If Loading Then
                    NewFile = FilePath & Format(xexif.EXIFmodified, "yyyy-mm-dd hh-nn-ss") & ".JPG"
                    If xexif.EXIFmodified = "12:00:00 AM" Then
                        NewFile = FilePath & Format$(FileDate(itmx.Key), "yyyy-mm-dd hh-nn-ss") & ".JPG"
                    End If
                End If
                Set xexif = Nothing
                If Loading Then
                    If bFileExists(NewFile) = False Then
                        SetAttr itmx.Key, vbArchive
                        Name itmx.Key As NewFile
                    Else
                        If ShowMsg("File : " & GetFName(NewFile) & " exist!" & vbCrLf & "Replace file?", vbYesNo, "File Exits") Then
                            SetAttr NewFile, vbArchive
                            Kill NewFile
                            SetAttr itmx.Key, vbArchive
                            Name itmx.Key As NewFile
                        End If
                    End If
                End If
                
            Case ImgOperation.Rotate_90:   'rotate 90
                sbStatusBar.Panels(1) = itmx.Text & "==> Rotate 90..."
                Set xexif = New cEXIF
                If xexif.Load(itmx.Key) = True Then
                    xexif.SaveAs EncoderValueTransformRotate90
                End If
                Set xexif = Nothing
                
            Case ImgOperation.Rotate_180:   'rotate 180
                sbStatusBar.Panels(1) = itmx.Text & "==> Rotate 180..."
                Set xexif = New cEXIF
                If xexif.Load(itmx.Key) = True Then
                    xexif.SaveAs EncoderValueTransformRotate180
                End If
                Set xexif = Nothing
                
            Case ImgOperation.Rotate_270:   'rotate 270
                sbStatusBar.Panels(1) = itmx.Text & "==> Rotate 270..."
                Set xexif = New cEXIF
                If xexif.Load(itmx.Key) = True Then
                    xexif.SaveAs EncoderValueTransformRotate270
                End If
                Set xexif = Nothing
                
             Case ImgOperation.Flip_Vertical:   'Flip Vertical
                sbStatusBar.Panels(1) = itmx.Text & "==> Flip Vertical..."
                Set xexif = New cEXIF
                If xexif.Load(itmx.Key) = True Then
                    xexif.SaveAs EncoderValueTransformFlipVertical
                End If
                Set xexif = Nothing
                
            Case ImgOperation.Flip_Horizontal:   'Flip Horizontal
                sbStatusBar.Panels(1) = itmx.Text & "==> Flip Horizontal..."
                Set xexif = New cEXIF
                If xexif.Load(itmx.Key) = True Then
                    xexif.SaveAs EncoderValueTransformFlipHorizontal
                End If
                Set xexif = Nothing
                
            Case ImgOperation.moveto:
                sbStatusBar.Panels(1) = itmx.Text & "==> Move to..."
                DiskOps itmx.Key, LastPathSelect & itmx.Text, F_Move, 1
            ' Smart Folder - Must set to exifDate first
            Case ImgOperation.SmartMove:
                sbStatusBar.Panels(1) = itmx.Text & "==> Smart Move..."
                SmartFolder = SmartPath & Mid(itmx.Text, 6, 2) & Mid(itmx.Text, 3, 2) & "\"
                If Dir$(SmartFolder, vbDirectory) <> "" Then
                    DiskOps itmx.Key, SmartFolder & itmx.Text, F_Move, 1
                Else
                    ret = ShowMsg("Folder " & SmartFolder & vbCrLf & "Not Found." & vbCrLf & "Create Folder ?", vbYesNo, "Smart Move")
                    If ret Then
                        MkDir SmartFolder
                        DiskOps itmx.Key, SmartFolder & itmx.Text, F_Move, 1
                    End If
                End If
                
            Case ImgOperation.copyto
                sbStatusBar.Panels(1) = itmx.Text & "==> copy to..."
                DiskOps itmx.Key, LastPathSelect & itmx.Text, F_CopySmart, 1
                
            Case ImgOperation.AddExifDate
                sbStatusBar.Panels(1) = itmx.Text & "==> Add Exif Date..."
                PrintDate itmx.Key
        End Select
        sbStatusBar.Panels(1) = ""
    End If
    
Next
Me.MousePointer = 0
GotoForm FilePath
Exit Sub
OperErr:
    ShowMsg Err.Description, vbOKOnly, "Operation Error"
    Resume Next
End Sub

Public Sub ShowExif(FileName As String)
Dim xexif As cEXIF
Set xexif = New cEXIF
If xexif.Load(FileName) = True Then
    LblExifdata = xexif.EXIFmake & vbCrLf
    LblExifdata = LblExifdata & xexif.EXIFmodel & vbCrLf
     LblExifdata = LblExifdata & Mid(xexif.EXIFsoftware, 1, 28) & vbCrLf
     LblExifdata = LblExifdata & xexif.EXIFmodified & vbCrLf
    ' LblExifdata = LblExifdata & xexif.EXIFFNumber & vbCrLf
    ' LblExifdata = LblExifdata & xexif.EXIFiso & vbCrLf
    ' LblExifdata = LblExifdata & xexif.ExifShutterSpeed & vbCrLf
    ' LblExifdata = LblExifdata & xexif.ExifExposureProg & vbCrLf
     LblExifdata = LblExifdata & xexif.Width & vbCrLf
     LblExifdata = LblExifdata & xexif.Height
End If
Set xexif = Nothing
End Sub

Sub LoadTitle()
LblExifTitle = "Make" & vbCrLf
LblExifTitle = LblExifTitle & "Model" & vbCrLf
LblExifTitle = LblExifTitle & "Software" & vbCrLf
LblExifTitle = LblExifTitle & "Date/Time" & vbCrLf
'LblExifTitle = LblExifTitle & "F Number" & vbCrLf
'LblExifTitle = LblExifTitle & "Iso" & vbCrLf
'LblExifTitle = LblExifTitle & "Speed" & vbCrLf
'LblExifTitle = LblExifTitle & "Mode" & vbCrLf
LblExifTitle = LblExifTitle & "Width" & vbCrLf
LblExifTitle = LblExifTitle & "Height"
TMcmdbutton1.Caption = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button1", "")
TMcmdbutton1.Tag = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button1", "Tag")
TMcmdbutton2.Caption = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button2", "")
TMcmdbutton2.Tag = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button2", "Tag")
TMcmdbutton3.Caption = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button3", "")
TMcmdbutton3.Tag = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button\Button3", "Tag")
End Sub

Public Sub OperDelete()
FilePath = Me.ActiveForm.Caption
Dim itmx As ListItem
Dim ximg As cIMAGE
Dim action As Boolean
Dim ret&
Set itmx = Me.ActiveForm.Lv1.SelectedItem
action = False
For Each itmx In Me.ActiveForm.Lv1.ListItems
    If itmx.Selected And ShowDelPic Then
        Set ximg = New cIMAGE
         ximg.Load itmx.Key
        If ximg.ImageHeight < ximg.ImageWidth Then
            ximg.ReSize 120, 0, False
        Else
            ximg.ReSize 0, 120, False
        End If
        With FrmMessage
        .Image1.Visible = False
        .Image1.Width = ximg.ImageWidth
        .Image1.Height = ximg.ImageHeight
        .Image1.Picture = ximg.Picture
        .Image1.Visible = True
        Set ximg = Nothing
        End With
        If ShowMsg("Delete File : " & itmx.Text & vbCrLf & "Are You Sure ?", vbYesNo, "Delete File") = True Then
            action = True
            If GetAttr(itmx.Key) <> vbArchive Then
            SetAttr itmx.Key, vbArchive
            End If
           ret = DiskOps(itmx.Key, itmx.Key, F_DelUndo, 1)
        'cannot undo
          ' ret = DeleteFile(itmx.Key)
          '  Debug.Print "delete", ret, GetAttr(itmx.Key)
'            Kill itmx.Key
        End If
    End If
    If itmx.Selected And Not ShowDelPic Then
        action = True
        If GetAttr(itmx.Key) <> vbArchive Then
            SetAttr itmx.Key, vbArchive
        End If
        sbStatusBar.Panels(1) = itmx.Text & "==> Delete ..."
        ret = DiskOps(itmx.Key, itmx.Key, F_DelUndo, 1)
         'DeleteFile itmx.Key
    End If
Next
sbStatusBar.Panels(1) = ""
If action Then GotoForm FilePath
End Sub


Sub PrintDate(sFileName$)
Dim xexif As cEXIF
Dim sexif As String
Dim rc As RECT
    Set xexif = New cEXIF
    xexif.Load sFileName
    If ChkDateAttach = "0" Then
        sexif = Format(xexif.EXIFmodified, "dd-mm-yyyy hh:nn AM/PM")
    Else
        sexif = Format(DateAttach, "dd-mm-yyyy")
    End If
    With Me.ActiveForm
        .scaleMode = 3
        .PicExif.Picture = LoadPicture
        .PicExif.Cls
        .PicExif.Refresh
        .PicExif.Forecolor = DFontColor
        .PicExif.Font = DFontName
        .PicExif.FontSize = .PicExif.Height \ 12  'DFontSize
        .PicExif.AutoRedraw = True
        .PicExif.Height = xexif.Height
        .PicExif.Width = xexif.Width
        SetRect rc, 1, 1, .PicExif.Width - Val(OffsetX), .PicExif.Height - Val(OffsetY)
        xexif.PaintDC .PicExif.hDC, 0, 0
        DrawText .PicExif.hDC, sexif, -1, rc, FontAlign
        .PicExif.AutoRedraw = False
        .PicExif.Picture = .PicExif.Image
        .PicExif.Refresh
        SaveJPG .PicExif.Picture, GetFPath(sFileName) & "x_" & GetFName(sFileName), 92
        .scaleMode = 1
    End With
    InvalidateRect hwnd, rc, False
    SaveExif GetFPath(sFileName) & "x_" & GetFName(sFileName), xexif
    Set xexif = Nothing
End Sub

Sub LoadWallPaper()
WallFileName = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "Filename")
WallTiles = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "Tiles")
WallBackColor = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackColor")
WallBackPicture = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackPicture")
End Sub

Sub LoadPrintDate()
FontAlign = DT_BOTTOM Or DT_RIGHT Or DT_SINGLELINE
DFontName = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Name")
DFontSize = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Size")
DFontColor = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Color")
OffsetX = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "x")
OffsetY = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "y")
DFormat = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach", "Format")
DateAttach = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach", "Date")
ChkDateAttach = QueryValue(HKEY_CURRENT_USER, "XPViewer\Attach", "Check")
If ChkDateAttach = "1" Then
    MnuAddExifDate.Caption = "Add date :" & Format(DateAttach, "dd-mm-yyyy")
Else
    MnuAddExifDate.Caption = "Add Exif Date"
End If
End Sub

Sub LoadPlugIns()
On Error Resume Next
Dim i%
PluginCount = Val(QueryValue(HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugInCount"))
If PluginCount > 0 Then
For i% = 1 To PluginCount
    PlugInSoftware(i%) = QueryValue(HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugIn " & Trim$(i%))
     Load MnuPlugIn1(i% - 1)
    MnuPlugIn1(i% - 1).Caption = GetFName(PlugInSoftware(i%))
    MnuPlugIn1(i% - 1).Tag = PlugInSoftware(i%)
    MnuPlugIn1(i% - 1).Visible = True
Next i%
End If
End Sub

Sub LoadFolderReg()
LastPathSelect = QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder", "LastPathSelect")
CDBurnPath = QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder", "CdBurnPath")
FavoritePath = QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder", "FavoritePath")
StartPath = QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder", "StartPath")
SmartPath = QueryValue(HKEY_CURRENT_USER, "XPViewer\Folder", "SmartPath")
End Sub

Sub LoadMiscReg()
BurnSoftware = QueryValue(HKEY_CURRENT_USER, "XPViewer\Application", "BurnSoftware")
CompanyName = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Picture", "CompanyName")
'TMcmdBurnCD.Visible = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Button", "BurnCD")
ThumbnailSize = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailSize")
SlideTimer = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Picture", "SlideTimer")
Shadow = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailShadow")
End Sub

Sub GetExif(exifFname$)
Dim xexif As cEXIF
    Set xexif = New cEXIF
    If xexif.Load(exifFname) Then
        Software = xexif.EXIFsoftware
        ModiDate = xexif.EXIFmodified
    End If
    Set xexif = Nothing
   
End Sub

Sub SaveExif(ImgFName$, xexif2 As cEXIF)
Dim xexif As cEXIF
       Set xexif = New cEXIF
    If xexif.Load(ImgFName) Then
        xexif.EXIFsoftware = xexif2.EXIFsoftware
        xexif.EXIFmodified = xexif2.EXIFmodified
        xexif.Save
    End If
        Set xexif = Nothing
End Sub

Sub AddPath2Cb(Path$)
Dim i%
For i% = cbPath.ListCount - 1 To 0 Step -1
    If UCase$(cbPath.List(i%)) = UCase$(Path) Then Exit Sub
Next
cbPath.AddItem Path$
End Sub


