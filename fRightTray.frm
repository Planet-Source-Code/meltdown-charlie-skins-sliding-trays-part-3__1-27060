VERSION 5.00
Begin VB.Form fRightTray 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   2520
      Left            =   -150
      Shape           =   4  'Rounded Rectangle
      Top             =   225
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Put all your cool stuff and doodads, widgets and whatever on these cool sliding panels "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1290
      Left            =   375
      TabIndex        =   0
      Top             =   855
      Width           =   2910
   End
End
Attribute VB_Name = "fRightTray"
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

Private Sub Form_Load()
    Dim hrgn As Long
    Dim grad As New clsGradient
    
    hrgn = LoadRegionDataFromFile(PATHSTR & "rh.pts")
    SetWindowRgn hWnd, hrgn, True
    DeleteObject hrgn
    ' makes a funky gradient on the form
    grad.MetalCylander RGB(200, 200, 200), _
                       RGB(50, 50, 50), _
                       RGB(215, 215, 215), _
                       RGB(25, 25, 25), _
                       Width \ 6, _
                       Width - 150, Width, Me
End Sub

' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================

