VERSION 5.00
Begin VB.Form FrmMessage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Title"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMessage.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProXPViewer.TMcmdbutton CmdNo 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      XpButton        =   1
      Caption         =   "No"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
   End
   Begin ProXPViewer.TMcmdbutton CmdYes 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      XpButton        =   1
      Caption         =   "Yes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
   End
   Begin VB.Label LblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   765
      TabIndex        =   0
      Top             =   1050
      Width           =   4305
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   480
   End
   Begin VB.Label LblTitle2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2530
      TabIndex        =   1
      Top             =   370
      Width           =   480
   End
   Begin VB.Label LblMsg2 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   780
      TabIndex        =   5
      Top             =   1065
      Width           =   4305
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   360
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   4095
      Left            =   0
      Picture         =   "FrmMessage.frx":53F9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdNo_Click()
ReturnMsg = "No"
Unload Me
End Sub

Private Sub CmdYes_Click()
ReturnMsg = "Yes"
Unload Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub


