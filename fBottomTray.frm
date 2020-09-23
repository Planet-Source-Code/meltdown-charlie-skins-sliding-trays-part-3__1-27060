VERSION 5.00
Begin VB.Form fBottomTray 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   450
      TabIndex        =   1
      Top             =   690
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Test"
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   1335
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   1980
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   -30
      Width           =   3960
   End
End
Attribute VB_Name = "fBottomTray"
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

Private Sub Command1_Click()
    MsgBox Text1
End Sub

Private Sub Form_Load()
    Dim hrgn As Long
    Dim grad As New clsGradient
    
    hrgn = CreateRoundRectRgn(0, 0, ScaleX(Text1.Width, vbTwips, vbPixels), ScaleY(Text1.Height, vbTwips, vbPixels), 15, 15)
    SetWindowRgn Text1.hWnd, hrgn, True
    DeleteObject hrgn
    hrgn = LoadRegionDataFromFile(PATHSTR & "bot.pts")
    SetWindowRgn hWnd, hrgn, True
    DeleteObject hrgn
    ' makes a funky gradient on the form
    grad.MetalCylander RGB(25, 25, 25), _
                       RGB(215, 215, 215), _
                       RGB(50, 50, 50), _
                       RGB(200, 200, 200), _
                       Width \ 2, _
                       (Width \ 2) + (Width \ 3), _
                       Width, Me
End Sub

' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================

