VERSION 5.00
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Slider Plus"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnBottomUp 
      Height          =   165
      Index           =   2
      Left            =   1410
      Picture         =   "fMain.frx":0000
      Top             =   2535
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image btnBottomUp 
      Height          =   165
      Index           =   1
      Left            =   1380
      Picture         =   "fMain.frx":035F
      Top             =   2745
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   2
      Left            =   1185
      Picture         =   "fMain.frx":06BE
      Top             =   2550
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   1
      Left            =   1185
      Picture         =   "fMain.frx":0A1C
      Top             =   2730
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit us on the web by clicking below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   7
      Top             =   2025
      Width           =   3210
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WWW.INTACTINTERACTIVE.COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   570
      TabIndex        =   6
      Top             =   2355
      Width           =   3600
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "With the techniques illustrated in this project you can create cool and different looking interfaces."
      Height          =   975
      Left            =   630
      TabIndex        =   5
      Top             =   1005
      Width           =   3495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stuff"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   2685
      Width           =   525
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   4425
      Shape           =   2  'Oval
      Top             =   2940
      Width           =   180
   End
   Begin VB.Image btnRightIn 
      Height          =   300
      Index           =   2
      Left            =   4155
      Picture         =   "fMain.frx":0D7A
      Top             =   1395
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnRightIn 
      Height          =   300
      Index           =   1
      Left            =   3960
      Picture         =   "fMain.frx":10E4
      Top             =   1425
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnRightOut 
      Height          =   300
      Index           =   2
      Left            =   4140
      Picture         =   "fMain.frx":144E
      Top             =   1200
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnRightOut 
      Height          =   300
      Index           =   1
      Left            =   3960
      Picture         =   "fMain.frx":17B8
      Top             =   1215
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnToolsIn 
      Height          =   300
      Index           =   2
      Left            =   555
      Picture         =   "fMain.frx":1B22
      Top             =   1200
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnToolsIn 
      Height          =   300
      Index           =   1
      Left            =   405
      Picture         =   "fMain.frx":1E8C
      Top             =   1170
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnToolsOut 
      Height          =   300
      Index           =   2
      Left            =   540
      Picture         =   "fMain.frx":21F6
      Top             =   1020
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnToolsOut 
      Height          =   300
      Index           =   1
      Left            =   390
      Picture         =   "fMain.frx":2560
      Top             =   1005
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3420
      TabIndex        =   3
      Top             =   555
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Cool Tools "
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   225
      TabIndex        =   2
      Top             =   555
      Width           =   1335
   End
   Begin VB.Image btnBottomUp 
      Height          =   165
      Index           =   0
      Left            =   375
      Picture         =   "fMain.frx":28CA
      Top             =   3000
      Width           =   300
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   0
      Left            =   90
      Picture         =   "fMain.frx":2C29
      Top             =   3000
      Width           =   300
   End
   Begin VB.Image btnRightOut 
      Height          =   300
      Index           =   0
      Left            =   4470
      Picture         =   "fMain.frx":2F87
      Top             =   570
      Width           =   165
   End
   Begin VB.Image btnRightIn 
      Height          =   300
      Index           =   0
      Left            =   4455
      Picture         =   "fMain.frx":32F1
      Top             =   855
      Width           =   165
   End
   Begin VB.Image btnToolsIn 
      Height          =   300
      Index           =   0
      Left            =   75
      Picture         =   "fMain.frx":365B
      Top             =   825
      Width           =   165
   End
   Begin VB.Image btnToolsOut 
      Height          =   300
      Index           =   0
      Left            =   60
      Picture         =   "fMain.frx":39C5
      Top             =   555
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   165
      Top             =   120
      Width           =   4020
   End
   Begin VB.Image btnClose 
      Height          =   300
      Left            =   4200
      Top             =   150
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4290
      TabIndex        =   1
      Top             =   195
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Slider Plus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1635
      TabIndex        =   0
      Top             =   180
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   4185
      X2              =   225
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   4140
      X2              =   180
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   4110
      X2              =   150
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   4095
      X2              =   135
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   4110
      X2              =   195
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   4155
      X2              =   345
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   315
      Left            =   4185
      Shape           =   3  'Circle
      Top             =   150
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   210
      Top             =   555
      Width           =   4260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2325
      Left            =   75
      Top             =   690
      Width           =   180
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   240
      Top             =   3000
      Width           =   4140
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2160
      Left            =   4455
      Top             =   735
      Width           =   180
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2670
      Width           =   930
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================
'
' Developers please note ...
'
' This project was developed by us as a test bed for techniques that we use in our
' own software, so forgive any problems in it, if you do find any let us know and
' we will endevaour to resolve them.
'
' Also forgive any sparseness of comments or oddities of sytle, this is after all
' only a test bed app and not the final commercial code...
'
' Also don't forget to visit us at www.intactinteractive.com and see what we are up
' to .
'
' ===================================================================================
' ===================================================================================

'For dragging the form
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const SW_SHOWMAXIMIZED = 3

Private bLeftOut As Boolean
Private bRightOut As Boolean
Private bBottomOut As Boolean

Private iRelLeftTrayOffset As Integer
Private iRelRightTrayOffset As Integer
Private iRelBottomTrayOffset As Integer

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' down comes the bottom slider
Private Sub btnBottomDown_Click(Index As Integer)
    ' don't do anything if it's already slid down
    If bBottomOut Then Exit Sub
    ' postition the tray ready to slide down
    fBottomTray.Left = fMain.Left + 150
    fBottomTray.Top = fMain.Top
    fBottomTray.Show
    iRelBottomTrayOffset = fMain.Top + fMain.Height - 100
    Do While fBottomTray.Top < iRelBottomTrayOffset
        'down 1 more pixel
        fBottomTray.Top = fBottomTray.Top + 15
        ' make sure the main form stays on top
        fMain.ZOrder
        DoEvents
    Loop
    bBottomOut = True
End Sub

Private Sub btnBottomDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hilite the little green down arrow and unhilite the up arrow by swapping pics
    btnBottomDown(0) = btnBottomDown(2)
    btnBottomUp(0) = btnBottomUp(1)
End Sub

' up it comes again
Private Sub btnBottomUp_Click(Index As Integer)
    ' only do this if we are slid out
    If bBottomOut Then
        ' reset the target offest ...
        iRelBottomTrayOffset = fMain.Top
        ' and reel her in
        Do While fBottomTray.Top > iRelBottomTrayOffset
            ' in two pixels
            fBottomTray.Top = fBottomTray.Top - 30
            ' and make sure the main form stays on top
            fMain.ZOrder
            DoEvents
        Loop
        ' now hide the tray and set the bBottomOut flag false to allow sliding down
        fBottomTray.Hide
        bBottomOut = False
    End If
End Sub

Private Sub btnBottomUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hilite the little green up arrow and unhilite the down arrow by swapping pics
    btnBottomDown(0) = btnBottomDown(1)
    btnBottomUp(0) = btnBottomUp(2)
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' we pass on the event to the form so it can take care of unhiliting any pics that
    ' are hilighted ...
    Form_MouseMove Button, Shift, X, Y
End Sub


Private Sub SlideLeftIn()
    ' slide in the tools tray
    Do While fLeftTray.Left < fMain.Left
        fLeftTray.Left = fLeftTray.Left + 30
        fMain.ZOrder
        DoEvents
    Loop
    iRelLeftTrayOffset = fMain.Left
    fLeftTray.Hide
    bLeftOut = False
End Sub

' slide out the right tray ...
Private Sub SlideRightIn()
    iRelRightTrayOffset = fMain.Left
    Do While fRightTray.Left > iRelRightTrayOffset
        fRightTray.Left = fRightTray.Left - 30
        fMain.ZOrder
        DoEvents
    Loop
    fRightTray.Hide
    bRightOut = False
End Sub

Private Sub btnRightIn_Click(Index As Integer)
        If bRightOut Then SlideRightIn
End Sub

Private Sub btnRightIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnRightOut(0) = btnRightOut(1)
    btnRightIn(0) = btnRightIn(2)
End Sub

Private Sub btnRightOut_Click(Index As Integer)
    ' slide out the right tray
    If bRightOut Then Exit Sub
    fRightTray.Left = fMain.Left
    fRightTray.Top = fMain.Top + 100
    fRightTray.Refresh
    fRightTray.Show
    iRelRightTrayOffset = fMain.Left + fMain.Width - 180
    Do While fRightTray.Left < iRelRightTrayOffset
        fRightTray.Left = fRightTray.Left + 30
        fMain.ZOrder
        DoEvents
    Loop
    bRightOut = True
End Sub

Private Sub btnRightOut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnRightOut(0) = btnRightOut(2)
    btnRightIn(0) = btnRightIn(1)
End Sub

' slide in the tools tray ...
Private Sub btnToolsIn_Click(Index As Integer)
    If bLeftOut Then SlideLeftIn
End Sub

Private Sub btnToolsIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnToolsOut(0) = btnToolsOut(1)
    btnToolsIn(0) = btnToolsIn(2)
End Sub

' slide out the tools tray ...
Private Sub btnToolsOut_Click(Index As Integer)
    If bLeftOut Then Exit Sub
    fLeftTray.Left = iRelLeftTrayOffset
    fLeftTray.Top = fMain.Top + 100
    
    fLeftTray.Show
    Do While fLeftTray.Left > fMain.Left - fLeftTray.Width + 100
        fLeftTray.Left = fLeftTray.Left - 30
        fMain.ZOrder
        DoEvents
    Loop
    iRelLeftTrayOffset = fLeftTray.Width - 100
    bLeftOut = True
End Sub

Private Sub btnToolsOut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnToolsOut(0) = btnToolsOut(2)
    btnToolsIn(0) = btnToolsIn(1)
End Sub

Private Sub Form_Load()
    Dim hrgn As Long
    Dim grad As New clsGradient
    
    PATHSTR = App.Path & IIf(Mid(App.Path, Len(App.Path), 1) = "\", "", "\")
    hrgn = LoadRegionDataFromFile(PATHSTR & "main.pts")
    SetWindowRgn hWnd, hrgn, True
    DeleteObject hrgn
    ' initialize our tray state flags
    bLeftOut = False
    bRightOut = False
    bBottomOut = False
    ' here we are
    Show
    ' put ourselves at the top of the local window stack
    fMain.ZOrder
    ' initialize the offset postions for the trays
    iRelLeftTrayOffset = fMain.Left
    iRelRightTrayOffset = fMain.Left
    fLeftTray.Left = fMain.Left
    fLeftTray.Top = fMain.Top + 100
    fRightTray.Left = fMain.Left
    fRightTray.Top = fMain.Top + 100
    fBottomTray.Left = fMain.Left + 150
    fBottomTray.Top = fMain.Top
    ' makes a funky gradient on the form
    grad.MetalCylander RGB(25, 25, 25), _
                       RGB(215, 215, 215), _
                       RGB(50, 50, 50), _
                       RGB(200, 200, 200), _
                       Width \ 2, _
                       (Width \ 2) + (Width \ 3), _
                       Width, Me
End Sub

Private Sub btnClose_Click()
    ' call appexit to close the application
    AppExit
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' we unhilite all of those arrows here so that we don't get sticking
    ' light effects you will notice that other mousemove subs delegate to
    ' this code to avoid repetition ...
    btnToolsOut(0) = btnToolsOut(1)
    btnToolsIn(0) = btnToolsIn(1)
    btnRightOut(0) = btnRightOut(1)
    btnRightIn(0) = btnRightIn(1)
    btnBottomDown(0) = btnBottomDown(1)
    btnBottomUp(0) = btnBottomUp(1)
End Sub

Private Sub Form_Paint()
    ' we make sure that all of the slid out trays are redrawn - ie. when we have just
    ' been moved the paint event usually fires
    If bBottomOut Then fBottomTray.Show
    If bRightOut Then fRightTray.Show
    If bLeftOut Then fLeftTray.Show
End Sub

' move da window
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' first release the current mouse capture scope so that dragging is possible
    ReleaseCapture
    ' then simply tell the system that the form is being dragged
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' we need to keep updating the relative positions for each tray window so this
    ' is the obvious place to do it I think ...
    If bLeftOut Then
        fLeftTray.Left = fMain.Left - iRelLeftTrayOffset
        fLeftTray.Top = fMain.Top + 100
    End If
    If bRightOut Then
        fRightTray.Left = fMain.Left + fMain.Width - 180
        fRightTray.Top = fMain.Top + 100
    End If
    If bBottomOut Then
        fBottomTray.Left = fMain.Left + 150
        fBottomTray.Top = fMain.Top + fMain.Height - 100
    End If
    ' also don't forget to pass on the event to the form so that hilighted arrows can
    ' be unhilited ...
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' delegate - delegate - delegate - let the form take care of it ...
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' delegate - delegate - delegate - let the form take care of it ...
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' delegate - delegate - delegate - let the form take care of it ...
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label7_Click()
    ' run the Internet Explorer or associated program and open it at our website's
    ' address - the window is maximized ...
    ShellExecute Me.hWnd, "open", "http://www.intactinteractive.com", ByVal 0&, "", SW_SHOWMAXIMIZED
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

' ===================================================================================
' Written by M Ferris  - Intact Interactive Software - 2001
' ===================================================================================

