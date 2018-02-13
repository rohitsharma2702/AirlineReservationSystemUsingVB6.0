VERSION 5.00
Begin VB.Form Form92 
   BackColor       =   &H00C0C000&
   Caption         =   "Book Your Flight Here"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Etihad Airways"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delta Airlines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "British Airways"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Emirates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Etihad Airways"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   11775
      Begin VB.PictureBox Picture4 
         Height          =   1095
         Left            =   8400
         Picture         =   "Form92.frx":0000
         ScaleHeight     =   1035
         ScaleWidth      =   3195
         TabIndex        =   30
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FFFF80&
         Caption         =   "Etihad Airways 3        [ Saturday( 10 p.m.)  ]              ( Rs.4000 + Rs.4000/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   7935
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FFFF80&
         Caption         =   "Etihad Airways 2        [   Friday ( 10 p.m. )   ]             ( Rs.4000 + Rs.4000/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   7935
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FFFF80&
         Caption         =   "Etihad Airways 1        [ Thursday(10 p.m.)   ]             ( Rs.4000 + Rs.4000/- )"
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
         Height          =   300
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Delta Airlines Flights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   11775
      Begin VB.PictureBox Picture3 
         Height          =   1095
         Left            =   8400
         Picture         =   "Form92.frx":1470
         ScaleHeight     =   1035
         ScaleWidth      =   3195
         TabIndex        =   29
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delta Airlines 3          [  Saturday( 6 p.m. )  ]              ( Rs.3500 + Rs.3500/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   7935
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delta Airlines 2          [     Friday ( 6 p.m. )  ]              ( Rs.3500 + Rs.3500/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   7935
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delta Airlines 1          [ Thursday( 6 p.m.)  ]               ( Rs.3500 + Rs.3500/- )"
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
         Height          =   300
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "British Airways Flights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   11775
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   8400
         Picture         =   "Form92.frx":2361
         ScaleHeight     =   1035
         ScaleWidth      =   3195
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFF80&
         Caption         =   "British Airways 3        [ Wednesday( 3 p.m. )  ]          ( Rs.3000 + Rs.3000/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   7935
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
         Caption         =   "British Airways 2        [ Tuesday    ( 3 p.m. )  ]           ( Rs.3000 + Rs.3000/- )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   7935
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "British Airways 1        [   Monday   ( 3 p.m.)   ]           ( Rs.3000 + Rs.3000/- )"
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
         Height          =   300
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "List of Available International Round-Trip Flights"
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
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14895
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Main Form"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Emirates Flights"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3000
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   11775
         Begin VB.PictureBox Picture1 
            Height          =   975
            Left            =   8400
            Picture         =   "Form92.frx":35F0
            ScaleHeight     =   915
            ScaleWidth      =   3195
            TabIndex        =   27
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFF80&
            Caption         =   "Emirates 3                 [ Wednesday( 9 a.m. ) ]           ( Rs.2500 + Rs.2500/- )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   300
            Left            =   360
            TabIndex        =   11
            Top             =   960
            Width           =   7935
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF80&
            Caption         =   "Emirates 2                 [     Tuesday( 9 a.m. )  ]           ( Rs.2500 + Rs.2500/- )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   300
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   7935
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF80&
            Caption         =   "Emirates 1                 [      Monday( 9 a.m.)   ]            ( Rs.2500 + Rs.2500/- )"
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
            Height          =   300
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   7935
         End
      End
   End
End
Attribute VB_Name = "Form92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As String, flag As Integer
Private Sub Command1_Click()
    If MsgBox("Are You Sure You Want To Log Out ? ", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
        MsgBox "Please Login To Continue", vbInformation, "Airline Reservation System"

        If flag = 0 Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from intround", con, adOpenDynamic, adLockOptimistic, adCmdText
            con.Execute ("delete from intround where fno = " & rst(0))
            rst.Close
            con.Close
        End If

        Form1.Show
        Unload Me
    End If
End Sub

Private Sub Command11_Click()
Form32.Show
Unload Me
End Sub

Private Sub Command2_Click()
Frame2.Visible = True
Command6.Visible = True
End Sub

Private Sub Command3_Click()
Frame3.Visible = True
Command7.Visible = True
End Sub

Private Sub Command4_Click()
Frame4.Visible = True
Command8.Visible = True
End Sub

Private Sub Command5_Click()
Frame5.Visible = True
Command9.Visible = True
End Sub

Private Sub Command6_Click()
    If Option1.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Emirates 1 ',' " & Option1.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option2.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Emirates 2 ',' " & Option2.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option3.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Emirates 3 ',' " & Option3.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    Else:  MsgBox "You Have Not Selected Any Flight.", vbCritical, "Airline Reservation System"
    End If
End Sub

Private Sub Command7_Click()
    If Option4.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' British Airways 1 ',' " & Option4.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option5.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' British Airways 2 ',' " & Option5.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option6.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' British Airways 3 ',' " & Option6.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    Else: MsgBox "You Have Not Selected Any Flight.", vbCritical, "Airline Reservation System"
    End If
End Sub

Private Sub Command8_Click()
    If Option7.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Delta Airlines 1 ',' " & Option7.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option8.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Delta Airlines 2 ',' " & Option8.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option9.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Delta Airlines 3 ',' " & Option9.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    Else: MsgBox "You Have Not Selected Any Flight.", vbCritical, "Airline Reservation System"
    End If
End Sub

Private Sub Command9_Click()
    If Option10.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Etihad Airways 1 ',' " & Option10.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option11.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Etihad Airways 2 ',' " & Option11.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    ElseIf Option12.Value = True Then
        If (MsgBox("Are You Sure You Want To Confirm Your Flight ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
            con.Open
            rst.Open "select max(fno) from confirmintround", con, adOpenDynamic, adLockOptimistic, adCmdText
            n = 1 + Val(rst(0))
            rst.Close
            con.Execute ("insert into confirmintround values(' " & n & " ',' ARSIRT' & '" & n & " ',' Etihad Airways 3 ',' " & Option12.Caption & " ')")
            MsgBox "Your Flight Has Been Confirmed", vbInformation, "Airline Reservation System"
            rst5.Open "select regno from confirmintround where regno = ' ARSIRT' & '" & n & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            MsgBox "Your Registration Number is :  " & rst5(0), vbInformation, "Airline Reservation System"
            rst5.Close
            con.Close
            Command6.Enabled = False
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
            Command8.Enabled = False
            Command9.Enabled = False
            Command11.Enabled = True
            flag = 1
            Form4.Show
            Unload Me
        End If
    Else: MsgBox "You Have Not Selected Any Flight.", vbCritical, "Airline Reservation System"
    End If
End Sub


Private Sub Picture1_Click()
frmSplash9.Show
End Sub

Private Sub Picture2_Click()
frmSplash10.Show
End Sub

Private Sub Picture3_Click()
frmSplash11.Show
End Sub

Private Sub Picture4_Click()
frmSplash12.Show
End Sub
