VERSION 5.00
Begin VB.Form fLeftTray 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   1200
      ItemData        =   "fLeftTray.frx":0000
      Left            =   495
      List            =   "fLeftTray.frx":001C
      TabIndex        =   1
      Top             =   1110
      Width           =   2925
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   495
      TabIndex        =   0
      Top             =   570
      Width           =   2925
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   345
      Top             =   2265
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   405
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   555
      Width           =   2970
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1215
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1110
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H00808080&
      Height          =   2475
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   255
      Width           =   3750
   End
End
Attribute VB_Name = "fLeftTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long

Private Sub Form_Load()
    Dim hrgn As Long
    Dim grad As New clsGradient
    
    hrgn = CreateRoundRectRgn(0, 0, ScaleX(Text1.Width, vbTwips, vbPixels), ScaleY(Text1.Height, vbTwips, vbPixels), 5, 5)
    SetWindowRgn Text1.hWnd, hrgn, True
    DeleteObject hrgn
    hrgn = CreateRoundRectRgn(0, 0, ScaleX(List1.Width, vbTwips, vbPixels), ScaleY(List1.Height, vbTwips, vbPixels), 20, 20)
    SetWindowRgn List1.hWnd, hrgn, True
    DeleteObject hrgn
    hrgn = LoadRegionDataFromFile(PATHSTR & "lh.pts")
    SetWindowRgn hWnd, hrgn, True
    DeleteObject hrgn
    ' makes a funky gradient on the form
    grad.MetalCylander RGB(25, 25, 25), _
                       RGB(215, 215, 215), _
                       RGB(50, 50, 50), _
                       RGB(200, 200, 200), _
                       Width \ 6, _
                       Width - 150, Width, Me
End Sub

Private Sub List1_Click()
    Text1 = List1.List(List1.ListIndex)
End Sub

' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================

