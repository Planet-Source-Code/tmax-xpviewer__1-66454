VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl TMcmdbutton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   ToolboxBitmap   =   "TmCmdbutton.ctx":0000
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   3600
      Picture         =   "TmCmdbutton.ctx":0312
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   5
      Top             =   1080
      Width           =   1350
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   3480
      Picture         =   "TmCmdbutton.ctx":2DF8
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   4
      Top             =   840
      Width           =   1350
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3360
      Picture         =   "TmCmdbutton.ctx":5959
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   480
      Width           =   1350
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   2760
      Picture         =   "TmCmdbutton.ctx":6527
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   2.45745e5
      Width           =   1350
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   3465
      Picture         =   "TmCmdbutton.ctx":6F59
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   90
      Width           =   1350
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   2
      Left            =   2880
      Picture         =   "TmCmdbutton.ctx":A656
      ScaleHeight     =   600
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   2.45745e5
      Width           =   6000
   End
   Begin PicClip.PictureClip pc5 
      Left            =   3120
      Top             =   2.45745e5
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "TmCmdbutton.ctx":FE18
   End
   Begin PicClip.PictureClip pc 
      Left            =   0
      Top             =   -1.50000e5
      _ExtentX        =   10583
      _ExtentY        =   1058
      _Version        =   393216
      Cols            =   5
      Picture         =   "TmCmdbutton.ctx":109F6
   End
   Begin PicClip.PictureClip pc2 
      Left            =   1680
      Top             =   2.45745e5
      _ExtentX        =   10583
      _ExtentY        =   1058
      _Version        =   393216
      Cols            =   5
      Picture         =   "TmCmdbutton.ctx":1C5C8
   End
End
Attribute VB_Name = "TMcmdbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINT_API
    x As Long
    y As Long
End Type
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const DT_CENTER = &H1
Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4
Public Enum ButtonStyle
    Xp_Normal = 0
    Xp_Metallic = 1
    CrystalRed = 2
    Xp_Met_Green = 3
    Gold = 4
    Apple = 5
    Custom = 6
End Enum
Dim m_Font As Font
Dim m_ForeColor As OLE_COLOR
Dim m_XpButton As ButtonStyle
Private m_txtRect  As RECT
Dim m_Picture As Picture
Dim m_PictureOriginal As Picture
Private m_sCaption  As String
Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event MouseLeave()

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Enabled = True
    XpButton = m_XpButton
    Set m_PictureOriginal = LoadPicture("")
    m_ForeColor = vbBlack
    Set Font = UserControl.Ambient.Font
    m_sCaption = "Tmax"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    make_xpbutton 1
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If UserControl.Enabled = False Then Exit Sub
        If x >= 0 And x <= UserControl.ScaleWidth And _
           y >= 0 And y <= UserControl.ScaleHeight Then
            ' Make all messages get sent to the UserControl for a while
            SetCapture UserControl.hwnd
            make_xpbutton 3
            RaiseEvent MouseMove(Button, Shift, x, y)
        Else
            ' Cursor went outside of the control. Release messages to be sent
            '  to wherever. Repaint the control with a "Lost focus" state
            make_xpbutton 0
            ReleaseCapture
            RaiseEvent MouseLeave
        End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    make_xpbutton 0
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    If UserControl.Enabled = True Then
        make_xpbutton 0
    Else
        make_xpbutton 2
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    m_XpButton = PropBag.ReadProperty("XpButton", ButtonStyle.Xp_Normal)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_PictureOriginal = PropBag.ReadProperty("Picture", Nothing)
    m_sCaption = PropBag.ReadProperty("Caption", "Tmax")
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.Ambient.Forecolor)
    XpButton = m_XpButton
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

Private Sub UserControl_Resize()
    Dim hRgn2 As Long
    hRgn2 = CreateRoundRectRgn(1, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 3, 3)
    SetWindowRgn UserControl.hwnd, hRgn2, True
    DeleteObject hRgn2
    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("XpButton", m_XpButton, ButtonStyle.Xp_Normal)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Caption", m_sCaption, "")
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, UserControl.Ambient.Forecolor)
End Sub

Public Property Let Caption(ByVal NewCaption As String)
Attribute Caption.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    m_sCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Paint
    Refresh
End Property

Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = m_ForeColor
End Property

Public Property Let Forecolor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property


Public Property Get XpButton() As ButtonStyle
    XpButton = m_XpButton
End Property

Public Property Let XpButton(ByVal New_XpButton As ButtonStyle)
    m_XpButton = New_XpButton
    PropertyChanged "XpButton"
    UserControl_Paint
End Property

Private Sub make_xpbutton(z As Integer)
    Dim brx, bry, bw, bh As Integer
    Dim Py1, Py2, Px1, Px2, Pw, Ph As Integer
    Pw = 3
    Ph = 3
    Px1 = 3
    Py1 = 3
    brx = UserControl.ScaleWidth - Pw
    bry = UserControl.ScaleHeight - Ph
    bw = UserControl.ScaleWidth - (Pw * 2)
    bh = UserControl.ScaleHeight - (Ph * 2)
    SetRect m_txtRect, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
    If XpButton = ButtonStyle.Custom Then
            pc.Picture = m_Picture
        Else
            pc.Picture = PicButton(XpButton).Picture
        End If
            Py2 = pc.Height - Py1
            Px2 = (pc.Width / 5) - Px1
            UserControl.PaintPicture pc.GraphicCell(z), 0, 0, Pw, Ph, 0, 0, Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), brx, 0, Pw, Ph, Px2, 0, Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), brx, bry, Pw, Ph, Px2, Py2, Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), 0, bry, Pw, Ph, 0, Py2, Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), Px1, 0, bw, Ph, Px1, 0, Px2 - Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), brx, Py1, Pw, bh, Px2, Py1, Pw, Py2 - Ph
            UserControl.PaintPicture pc.GraphicCell(z), 0, Py1, Pw, bh, 0, Py1, Pw, Py2 - Ph
            UserControl.PaintPicture pc.GraphicCell(z), Px1, bry, bw, Ph, Px1, Py2, Px2 - Pw, Ph
            UserControl.PaintPicture pc.GraphicCell(z), Px1, Py1, bw, bh, Px1, Py1, Px2 - Pw, Py2 - Ph
            Select Case z
            Case 0: DrawEdge UserControl.hDC, m_txtRect, BDR_RAISEDINNER, BF_RECT
            Case 1: DrawEdge UserControl.hDC, m_txtRect, BDR_INNER, BF_RECT
            Case 2: DrawEdge UserControl.hDC, m_txtRect, BDR_SUNKENINNER, BF_RECT
            Case 3: DrawEdge UserControl.hDC, m_txtRect, BDR_SUNKENOUTER, BF_RECT
            Case 4: DrawEdge UserControl.hDC, m_txtRect, BDR_SUNKENINNER, BF_RECT
            DrawEdge UserControl.hDC, m_txtRect, BDR_SUNKENINNER, BF_RECT
            End Select
   DrawCaption
End Sub

Sub DrawCaption()
    SetRect m_txtRect, 4, 4, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    DrawText UserControl.hDC, m_sCaption, -1, m_txtRect, lwFontAlign
End Sub

Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set m_PictureOriginal = New_Picture
    PropertyChanged "Picture"
    UserControl_Paint
End Property
