VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   5115
   ClientLeft      =   2520
   ClientTop       =   1170
   ClientWidth     =   10830
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOption 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   5145
      Left            =   -120
      ScaleHeight     =   5145
      ScaleWidth      =   10965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10965
      Begin VB.TextBox TxtSlideTimer 
         Appearance      =   0  'Flat
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
         Left            =   4305
         TabIndex        =   44
         Top             =   4590
         Width           =   435
      End
      Begin VB.CheckBox ChkBorder 
         Caption         =   "Check1"
         Height          =   165
         Left            =   2880
         TabIndex        =   42
         Top             =   4635
         Width           =   165
      End
      Begin VB.TextBox TxtThumbnailSize 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   41
         Top             =   4560
         Width           =   420
      End
      Begin VB.TextBox TxtTile 
         Alignment       =   2  'Center
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
         Left            =   8520
         TabIndex        =   38
         Text            =   "2"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox TxtScale 
         Alignment       =   2  'Center
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
         Left            =   9600
         TabIndex        =   37
         Text            =   "100"
         Top             =   1680
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   10080
         TabIndex        =   36
         Top             =   1680
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "TxtScale"
         BuddyDispid     =   196614
         OrigLeft        =   6960
         OrigTop         =   2880
         OrigRight       =   7200
         OrigBottom      =   3255
         Max             =   100
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5880
         TabIndex        =   30
         Top             =   960
         Width           =   4095
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   4080
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtOffsetY 
         Appearance      =   0  'Flat
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
         Left            =   9885
         TabIndex        =   26
         Text            =   "50"
         Top             =   3630
         Width           =   450
      End
      Begin VB.TextBox TxtOffsetX 
         Appearance      =   0  'Flat
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
         Left            =   9015
         TabIndex        =   25
         Text            =   "50"
         Top             =   3630
         Width           =   405
      End
      Begin VB.TextBox TxtFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox TxtFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox TxtFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   4575
      End
      Begin VB.ListBox LstPlugIn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         ItemData        =   "frmOptions.frx":000C
         Left            =   375
         List            =   "frmOptions.frx":000E
         TabIndex        =   4
         Top             =   3360
         Width           =   4575
      End
      Begin VB.TextBox TxtFolder 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CheckBox ChkDate 
         Caption         =   "Check1"
         Height          =   200
         Left            =   6960
         TabIndex        =   1
         Top             =   3000
         Width           =   200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8640
         TabIndex        =   2
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20054017
         CurrentDate     =   38964
      End
      Begin ProXPViewer.TMcmdbutton TMcmdDelete 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   3720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "-"
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
      Begin ProXPViewer.TMcmdbutton TMcmdAdd 
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   3360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "+"
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
      Begin ProXPViewer.TMcmdbutton TMcmdApply 
         Height          =   375
         Left            =   9720
         TabIndex        =   7
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "Apply"
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
      Begin ProXPViewer.TMcmdbutton TMcmdCancel 
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Left            =   6960
         TabIndex        =   9
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "OK"
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
      Begin ProXPViewer.TMcmdbutton TMcmdOpen 
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   11
         Tag             =   "Favorite Path"
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "..."
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
      Begin ProXPViewer.TMcmdbutton TMcmdOpen 
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   14
         Tag             =   "CDBurn Path"
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "..."
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
      Begin ProXPViewer.TMcmdbutton TMcmdOpen 
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   15
         Tag             =   "Startup Path"
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "..."
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
      Begin ProXPViewer.TMcmdbutton TMcmdOpen 
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   16
         Tag             =   "Smart Path"
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "..."
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
      Begin ProXPViewer.TMcmdbutton TMcmdOpen 
         Height          =   375
         Index           =   4
         Left            =   9960
         TabIndex        =   29
         Tag             =   "Favorite Path"
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         XpButton        =   1
         Caption         =   "..."
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slide Timer"
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
         Left            =   3300
         TabIndex        =   45
         Top             =   4575
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow"
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
         Left            =   2130
         TabIndex        =   43
         Top             =   4575
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thumbnail Size"
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
         Left            =   360
         TabIndex        =   40
         Top             =   4560
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Attach"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5760
         TabIndex        =   39
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scale"
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
         Left            =   9120
         TabIndex        =   35
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wall Tiles"
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
         Left            =   7680
         TabIndex        =   34
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background Image"
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
         Left            =   5880
         TabIndex        =   33
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5880
         TabIndex        =   32
         Top             =   1560
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WallPaper"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5760
         TabIndex        =   31
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label LblBackGround 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6960
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Height          =   1695
         Left            =   5760
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset    X             Y"
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
         Left            =   7980
         TabIndex        =   27
         Top             =   3675
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   2
         Height          =   1335
         Left            =   5760
         Top             =   2880
         Width           =   4815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font Color"
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
         Left            =   5880
         TabIndex        =   24
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label LblFontColor 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Favorite Folder"
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
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDBurn Folder"
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
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Folder"
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
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PlugIn Software"
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
         Left            =   360
         TabIndex        =   19
         Top             =   3120
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Smart Folder"
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
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Date"
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
         Left            =   5880
         TabIndex        =   17
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   5160
         Left            =   120
         Picture         =   "frmOptions.frx":0010
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10905
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Apply As Boolean

Private Sub ChkDate_Click()
ChkDateAttach = ChkDate.Value
End Sub

Private Sub DTPicker1_Change()
DateAttach = DTPicker1.Value
End Sub

Private Sub LblBackGround_DblClick()
dlg1.ShowColor
If dlg1.CancelError <> 32755 Then
    WallBackColor = dlg1.Color
    LblBackGround.BackColor = dlg1.Color
End If
End Sub

Private Sub LblFontColor_DblClick()
dlg1.ShowColor
If dlg1.CancelError <> 32755 Then
    DFontColor = dlg1.Color
    LblFontColor.BackColor = dlg1.Color
End If
End Sub

Private Sub TMcmdAdd_Click()
Dim Pluginstr As String
Pluginstr = BrowseForFolderDlg("", "Select A PlugIn", Me.hwnd, True)
LstPlugIn.AddItem Pluginstr
End Sub

Private Sub TMcmdApply_Click()
OffsetX = TxtOffsetX.Text
OffsetY = TxtOffsetY.Text
If ChkBorder.Value = 1 Then
    Shadow = "True"
Else
    Shadow = "False"
End If
ThumbnailSize = TxtThumbnailSize.Text
SlideTimer = TxtSlideTimer.Text
SmartPath = TxtFolder(3)
StartPath = TxtFolder(2)
CDBurnPath = TxtFolder(1)
FavoritePath = TxtFolder(0)
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailSize", ThumbnailSize, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "ThumbnailShadow", Shadow, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\Picture", "SlideTimer", SlideTimer, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "FavoritePath", FavoritePath, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "CDBurnPath", CDBurnPath, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "StartPath", StartPath, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Folder", "SmartPath", SmartPath, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach", "Date", DateAttach, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach", "Check", ChkDateAttach, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Name", DFontName, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Size", DFontSize, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Font", "Color", DFontColor, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "x", OffsetX, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach\Offset", "y", OffsetY, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Attach", "Format", "ff", REG_SZ

SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "Filename", WallFileName, REG_SZ
'WallTiles = QueryValue(HKEY_CURRENT_USER, "XPViewer\Control\WallPaper\Tiles", "Tiles")
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackColor", WallBackColor, REG_SZ
SetKeyValue HKEY_CURRENT_USER, "XPViewer\Control\WallPaper", "BackPicture", WallBackPicture, REG_SZ
SavePlugInsLst
If ChkDateAttach = "1" Then
    FrmMdi.MnuAddExifDate.Caption = "Add date :" & Format(DateAttach, "dd-mm-yyyy")
Else
    FrmMdi.MnuAddExifDate.Caption = "Add Exif Date"
End If
Apply = True
End Sub

Private Sub TMcmdCancel_Click()
    Unload Me
End Sub

Private Sub TMcmdDelete_Click()
On Error Resume Next
If LstPlugIn.ListCount > 0 Then
    LstPlugIn.RemoveItem LstPlugIn.ListIndex
End If
End Sub

Private Sub TMcmdOK_Click()
If Not Apply Then
    TMcmdApply_Click
End If
Unload Me
End Sub

Private Sub Form_Load()
LoadPlugInsLst
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
ChkDate.Value = Val(IIf(ChkDateAttach = "1", 1, 0))
LblFontColor.BackColor = Val(DFontColor)
LblBackGround.BackColor = Val(WallBackColor)
TxtFolder(4).Text = WallFileName
TxtOffsetX.Text = OffsetX
TxtOffsetY.Text = OffsetY
ChkBorder.Value = 0
If Shadow = "True" Then ChkBorder.Value = 1
'        DTPicker1.Value = DateAttach
Apply = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub TMcmdOpen_Click(Index As Integer)
Dim Folder As String
If Index = 4 Then
    Folder = BrowseForFolderDlg(TxtFolder(Index).Text, "Select " & TMcmdOpen(Index).Tag, Me.hwnd, True)
Else
    Folder = BrowseForFolderDlg(TxtFolder(Index).Text, "Select " & TMcmdOpen(Index).Tag, Me.hwnd)
End If
If Folder <> "" Then TxtFolder(Index).Text = Folder
Select Case Index
    Case 0
        FavoritePath = TxtFolder(0).Text
    Case 1
        CDBurnPath = TxtFolder(1).Text
    Case 2
        StartPath = TxtFolder(2).Text
End Select
End Sub

Sub LoadPlugInsLst()
On Error Resume Next
Dim i%
PluginCount = Val(QueryValue(HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugInCount"))
If PluginCount > 0 Then
For i% = 1 To PluginCount
    LstPlugIn.AddItem QueryValue(HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugIn " & Trim$(i%))
   
Next i%
End If
End Sub
Sub SavePlugInsLst()
On Error Resume Next
Dim i%, PluginCount%
PluginCount = LstPlugIn.ListCount
SetKeyValue HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugInCount", PluginCount, REG_SZ
If PluginCount > 0 Then
For i% = 1 To PluginCount
   SetKeyValue HKEY_CURRENT_USER, "XPViewer\PlugIn", "PlugIn " & Trim$(i%), LstPlugIn.List(i% - 1), REG_SZ
Next i%
End If
End Sub
