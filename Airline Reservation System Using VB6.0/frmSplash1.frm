VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   8085
   ClipControls    =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Get Your Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Type Your Username :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Answer Your Security Question and Get Your Password Within Seconds"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Type Your Security Answer :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Your Security Question :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
con.Open
    If Len(Trim(Text1.Text)) = 0 And Len(Trim(Text2.Text)) = 0 Then
        MsgBox "Please Fill Your Details First", vbCritical, "Airline Reservation System"
        Unload Me
        con.Close
    ElseIf Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Please Enter Your Username", vbCritical, "Airline Reservation System"
        Unload Me
        con.Close
    ElseIf Len(Trim(Text2.Text)) = 0 Then
        MsgBox "Please Enter Your Security Answer", vbCritical, "Airline Reservation System"
        Unload Me
        con.Close
    Else:   rst2.Open "select count(*) from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            rst3.Open "select upass,usec,uans from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                If rst2(0) = 0 Then
                    MsgBox "Invalid Username", vbCritical, "Airline Reservation System"
                    Unload Me
                    rst2.Close
                    rst3.Close
                    con.Close
                ElseIf rst2(0) > 0 Then
                    If StrComp(Text2.Text, Trim(rst3!uans)) <> 0 Then
                        MsgBox "Invalid Security Answer", vbCritical, "Airline Reservation System"
                        Unload Me
                        rst2.Close
                        rst3.Close
                        con.Close
                    Else: MsgBox "Your Password is : " & rst3!upass, vbInformation, "Airline Reservation System"
                        Unload Me
                        rst2.Close
                        rst3.Close
                        con.Close
                    End If
                End If
    End If
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
End Sub

Private Sub Text2_GotFocus()
Label1.Visible = True
Label5.Visible = True
    If Len(Trim(Text1.Text)) = 0 Then
        Label5.Caption = ""
    Else:   con.Open
            rst2.Open "select count(*) from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            rst3.Open "select usec from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                If rst2(0) > 0 Then
                    Label5.Caption = rst3(0)
                    rst2.Close
                    rst3.Close
                    con.Close
                Else: rst2.Close
                      rst3.Close
                      con.Close
                End If
    End If
End Sub
