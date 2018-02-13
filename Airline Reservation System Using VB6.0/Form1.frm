VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Login Page"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Forgot Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      Height          =   9015
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8955
      ScaleWidth      =   15795
      TabIndex        =   1
      Top             =   -600
      Width           =   15855
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6600
         Top             =   2040
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   12240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   420
         Left            =   12240
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   8160
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   4560
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   12240
         TabIndex        =   8
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   12240
         TabIndex        =   0
         Top             =   1080
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String
Private Sub Command1_Click()
con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
con.Open
rst.Open "select max(userial) from ulog", con, adOpenDynamic, adLockOptimistic, adCmdText
X = 1 + rst(0)
rst.Close
rst2.Open "select count(*) from aircust where uname=' " & Text1.Text & " ' and upass=' " & Text2.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst3.Open "select uname,upass from aircust where uname=' " & Text1.Text & " ' and upass=' " & Text2.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst4.Open "select count(*) from aircust where uname=' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
    If Text1.Text = "" And Text2.Text = "" Then
        MsgBox "Please Type Username & Password", vbExclamation, "Airline Reservation System"
        Text1.SetFocus
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
    ElseIf Text1.Text = "" Then
        MsgBox "Please Type Username", vbExclamation, "Airline Reservation System"
        Text1.SetFocus
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
    ElseIf Text2.Text = "" Then
        MsgBox "Please Type Password", vbExclamation, "Airline Reservation System"
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
        Text2.SetFocus
    ElseIf Text1.Text = "admin" And Text2.Text = "rohit" Then
        con.Execute ("insert into ulog values(' " & X & " ',' " & Text1.Text & " ',' Rohit Sharma ',' " & Now & " ',' " & Format(Now, "dddd") & " ',' Administrator ')")
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
        Form6.Show
        Unload Me
    ElseIf Val(rst4(0)) > 0 Then
        If Val(rst2(0)) = 0 Then
            MsgBox "Invalid Password", vbExclamation, "Airline Reservation System"
            rst2.Close
            rst3.Close
            rst4.Close
            con.Close
            Text2.Text = Clear
            Text2.SetFocus
        Else: rst5.Open "select ufirst,ulast from aircust where uname = ' " & Text1.Text & " ' and upass = ' " & Text2.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            con.Execute ("insert into ulog values(' " & X & " ',' " & Text1.Text & " ','" & rst5(0) & LTrim(rst5(1)) & "',' " & Now & " ',' " & Format(Now, "dddd") & " ',' Local User ')")
            rst2.Close
            rst3.Close
            rst4.Close
            rst5.Close
            con.Close
            Form32.Show
            Unload Me
        End If
    ElseIf Val(rst4(0)) = 0 Then
        MsgBox "Invalid Username", vbExclamation, "Airline Reservation System"
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
        Text1.Text = Clear
        Text1.SetFocus
    End If
End Sub

Private Sub Command1_GotFocus()
Command1.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command4.BackColor = &H8000000F
End Sub

Private Sub Command1_LostFocus()
Command1.BackColor = &H8000000F
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command4.BackColor = &H8000000F

End Sub

Private Sub Command2_Click()
If MsgBox("Are You Sure You Want To Exit The Project ? ", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
End
End If
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = &HFFFFC0
Command1.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command4.BackColor = &H8000000F
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HFFFFC0
Command1.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command4.BackColor = &H8000000F

End Sub

Private Sub Command3_Click()
frmSplash1.Show
End Sub

Private Sub Command3_GotFocus()
Command3.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command1.BackColor = &H8000000F
Command4.BackColor = &H8000000F
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command1.BackColor = &H8000000F
Command4.BackColor = &H8000000F

End Sub

Private Sub Command4_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command4_GotFocus()
Command4.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command1.BackColor = &H8000000F
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BackColor = &HFFFFC0
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command1.BackColor = &H8000000F

End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Picture1_GotFocus()
Command2.BackColor = &H8000000F
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format(Now, "dddd")
Label4.Caption = Format(Now, "dd-mmm-yyyy")
Label5.Caption = Format(Now, "hh:mm:ss am/pm")
End Sub
