VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   480
   End
   Begin ProXPViewer.TMcmdbutton TMOk 
      Height          =   465
      Left            =   7650
      TabIndex        =   3
      Top             =   5520
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   794
      Enabled3D       =   0   'False
      XpButton        =   3
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
   Begin ProXPViewer.TMcmdbutton TMXPThumb 
      Height          =   465
      Left            =   2520
      TabIndex        =   4
      Top             =   5520
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   847
      Enabled3D       =   0   'False
      XpButton        =   3
      Caption         =   "XpThumbnail"
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
   Begin ProXPViewer.TMcmdbutton TMXPCalendar 
      Height          =   465
      Left            =   1125
      TabIndex        =   5
      Top             =   5520
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   847
      Enabled3D       =   0   'False
      XpButton        =   3
      Caption         =   "XpCalendar"
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
   Begin ProXPViewer.TMcmdbutton TMVote 
      Height          =   465
      Left            =   195
      TabIndex        =   6
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   820
      Enabled3D       =   0   'False
      XpButton        =   3
      Caption         =   "Vote"
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
   Begin VB.Label LblCreator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mailto:tmax_visiber@yahoo.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4305
      MouseIcon       =   "frmAbout.frx":208F5
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "mailto:tmax_net@yahoo.com"
      Top             =   5625
      Width           =   3195
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Width           =   3045
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5520
      TabIndex        =   1
      Top             =   3960
      Width           =   3045
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim start As Boolean
Dim Apppath$
Const Addr1 = "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId="
Const Addr2 = "&lngWId=1"
Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
Dim hRgn1 As Long
    Me.scaleMode = 3
    hRgn1 = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 12, 12)
    SetWindowRgn Me.hwnd, hRgn1, True
    Me.scaleMode = 1
    Apppath = App.Path + IIf(Right(App.Path, 1) <> "\", "\", "")
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.FileDescription
    Me.Left = -Me.Width
    Me.Top = (Screen.Height - Me.Height) / 2
    start = True
    Timer1.Enabled = True
    If Dir(Apppath + "Vote.Adr") = "" Then
        CheckAddr
    Else
        OpenAddr
    End If
  ''  DropShadow Me.hwnd
End Sub

Private Sub LblCreator_DblClick()
Dim ret&
ret& = ShellExecute(Me.hwnd, "open", LblCreator.Tag, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub Timer1_Timer()
If start Then
    If (Me.Left < (Screen.Width - Me.ScaleWidth) / 2) Then
        Me.Left = Me.Left + 1400
    Else
        Timer1.Enabled = False
    End If
Else
    If (Me.Left < Screen.Width) Then
        Me.Left = Me.Left + 700
    Else
        Timer1.Enabled = False
        Unload Me
    End If
End If
End Sub

Sub CheckAddr()
Dim FileName$, txtfile$, Addr$
Dim f1
    FileName = Dir(Apppath + "@PSC*.txt")
    txtfile = Mid(FileName, InStr(1, FileName, "Me_") + 3, 5)
    Addr$ = Addr1 + txtfile + Addr2
    TMVote.Tag = Addr$
    f1 = FreeFile
    Open Apppath + "Vote.Adr" For Output As #f1
        Print #1, Addr$
    Close #1
End Sub

Sub OpenAddr()
Dim f1, ReadAdr
    f1 = FreeFile
    Open Apppath + "Vote.Adr" For Input As #f1
        Line Input #f1, ReadAdr
    Close #f1
    TMVote.Tag = ReadAdr
End Sub

Private Sub TMOK_Click()
 start = False
 Timer1.Enabled = True
End Sub

Private Sub TMVote_Click()
Dim ret&
    ret& = ShellExecute(Me.hwnd, "open", TMVote.Tag, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub TMXPCalendar_Click()
Dim ret&, Addr$
    Addr$ = Addr1$ + "58401" + Addr2$
    ret& = ShellExecute(Me.hwnd, "open", Addr$, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub TMXPThumb_Click()
Dim ret&, Addr$
    Addr$ = Addr1$ + "58401" + Addr2$
    ret& = ShellExecute(Me.hwnd, "open", Addr$, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub
