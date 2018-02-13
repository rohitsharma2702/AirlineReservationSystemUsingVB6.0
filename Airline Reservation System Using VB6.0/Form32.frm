VERSION 5.00
Begin VB.Form Form32 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Welcome Form"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   2055
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   600
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   10935
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   1920
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   720
         Top             =   480
      End
      Begin VB.Line Line33 
         BorderWidth     =   2
         X1              =   9360
         X2              =   9480
         Y1              =   4920
         Y2              =   5040
      End
      Begin VB.Line Line32 
         BorderWidth     =   2
         X1              =   9360
         X2              =   9480
         Y1              =   4920
         Y2              =   4800
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   9720
         X2              =   9360
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   10320
         X2              =   10200
         Y1              =   4920
         Y2              =   5040
      End
      Begin VB.Line Line29 
         BorderWidth     =   2
         X1              =   10320
         X2              =   10200
         Y1              =   4920
         Y2              =   4800
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   9960
         X2              =   10320
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9720
         Y1              =   5400
         Y2              =   5280
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9960
         Y1              =   5400
         Y2              =   5280
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9840
         Y1              =   5040
         Y2              =   5400
      End
      Begin VB.Line Line24 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9960
         Y1              =   4560
         Y2              =   4680
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9720
         Y1              =   4560
         Y2              =   4680
      End
      Begin VB.Line Line22 
         BorderWidth     =   2
         X1              =   9840
         X2              =   9840
         Y1              =   4560
         Y2              =   4920
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   840
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   255
      End
      Begin VB.Line Line21 
         X1              =   1320
         X2              =   1320
         Y1              =   1920
         Y2              =   2160
      End
      Begin VB.Line Line20 
         X1              =   1320
         X2              =   1080
         Y1              =   1920
         Y2              =   2040
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   5280
         X2              =   6360
         Y1              =   3000
         Y2              =   4440
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   4560
         X2              =   6120
         Y1              =   3000
         Y2              =   4440
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   5760
         X2              =   7200
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   4680
         X2              =   6840
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You Can Navigate This Aeroplane Using Arrow Keys"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   4920
         Width           =   6495
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         Height          =   255
         Left            =   9480
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         Height          =   255
         Left            =   9480
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   8640
         Shape           =   3  'Circle
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   6840
         Shape           =   3  'Circle
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   3000
         Shape           =   3  'Circle
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   1440
         X2              =   600
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1440
         Y1              =   1560
         Y2              =   2160
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   495
         Left            =   7680
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   495
         Left            =   5760
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   495
         Left            =   3840
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   495
         Left            =   1920
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   6600
         X2              =   5760
         Y1              =   4440
         Y2              =   3000
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   5640
         X2              =   6600
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   4080
         X2              =   5640
         Y1              =   3000
         Y2              =   4440
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   7320
         X2              =   6240
         Y1              =   360
         Y2              =   1560
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   6360
         X2              =   7320
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   4080
         X2              =   6360
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   9480
         X2              =   9480
         Y1              =   1560
         Y2              =   3000
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   360
         X2              =   480
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   600
         X2              =   360
         Y1              =   2160
         Y2              =   2760
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   960
         X2              =   600
         Y1              =   1920
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   1200
         X2              =   960
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   480
         X2              =   9480
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   1200
         X2              =   9480
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      FillColor       =   &H00C00000&
      Height          =   975
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   12855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Aviation Management System     "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   12135
   End
   Begin VB.Menu mnuBookFlight 
      Caption         =   "           Book &Flight             "
   End
   Begin VB.Menu mnuCheckStatus 
      Caption         =   "          Check &Status             "
   End
   Begin VB.Menu mnuUpdateMyProfile 
      Caption         =   "          Update &My &Profile      "
   End
   Begin VB.Menu mnuAboutDeveloper 
      Caption         =   "          About &Developer           "
   End
   Begin VB.Menu mnuFeedback 
      Caption         =   "          Feedback          "
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "          Help               "
   End
   Begin VB.Menu mnuLogOut 
      Caption         =   "          Log &Out          "
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        Frame1.Top = Frame1.Top - 200
    ElseIf KeyCode = 40 Then
        Frame1.Top = Frame1.Top + 200
    ElseIf KeyCode = 37 Then
        Frame1.Left = Frame1.Left - 200
    ElseIf KeyCode = 39 Then
        Frame1.Left = Frame1.Left + 200
    End If
End Sub

Private Sub mnuAboutDeveloper_Click()
frmSplash6.Show
End Sub


Private Sub mnuBookFlight_Click()
Form2.Show
Unload Me
End Sub

Private Sub mnuCheckStatus_Click()
frmSplash5.Show
End Sub

Private Sub mnuFeedback_Click()
Form5.Show
End Sub

Private Sub mnuHelp_Click()
Form62.Show
End Sub
Private Sub mnuUser_Click()
Form1.Show
End Sub

Private Sub mnuLogOut_Click()
    
        MsgBox "Please Login To Continue", vbInformation, "Airline Reservation System"
        Form1.Show
        Unload Me
    
End Sub

Private Sub mnuUpdateMyProfile_Click()
Form7.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Static a As Integer
a = a + 1
    If a = 1 Then
        Shape5.Visible = True
        Shape6.Visible = False
        Shape7.Visible = False
        Shape8.Visible = False
    ElseIf a = 2 Then
        Shape5.Visible = False
        Shape6.Visible = True
        Shape7.Visible = False
        Shape8.Visible = False
    ElseIf a = 3 Then
        Shape5.Visible = False
        Shape6.Visible = False
        Shape7.Visible = True
        Shape8.Visible = False
    ElseIf a = 4 Then
        Shape5.Visible = False
        Shape6.Visible = False
        Shape7.Visible = False
        Shape8.Visible = True
    Else: a = 0
    End If
End Sub

Private Sub Timer2_Timer()
Static X As Integer
X = X + 1
    If X = 1 Then
        Frame1.BackColor = &HFFC0FF
        Form32.BackColor = &HFFC0FF
    ElseIf X = 2 Then
        Frame1.BackColor = &H80FF80
        Form32.BackColor = &H80FF80
    ElseIf X = 3 Then
        Frame1.BackColor = &HFF8080
        Form32.BackColor = &HFF8080
    ElseIf X = 4 Then
        Frame1.BackColor = &H8080FF
        Form32.BackColor = &H8080FF
    ElseIf X = 5 Then
        Frame1.BackColor = &HC0FFFF
        Form32.BackColor = &HC0FFFF
    Else: X = 0
    End If
End Sub

Private Sub Timer3_Timer()
Label2.Caption = Mid(Label2.Caption, 2) & Left(Label2.Caption, 1)
End Sub
