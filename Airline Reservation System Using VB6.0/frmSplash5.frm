VERSION 5.00
Begin VB.Form frmSplash5 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Your Registration Status"
   ClientHeight    =   6780
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10425
   ClipControls    =   0   'False
   Icon            =   "frmSplash5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1815
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
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   150
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   10080
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4200
         Width           =   9855
      End
      Begin VB.TextBox Text8 
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
         Height          =   615
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Text7 
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
         Height          =   615
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox Text6 
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
         Height          =   615
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text5 
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
         Height          =   615
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text4 
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
         Height          =   615
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text3 
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
         Height          =   735
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
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
         Height          =   735
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         Caption         =   "Flight Details :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   4200
         TabIndex        =   19
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF80&
         Caption         =   "Flight Name :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2160
         TabIndex        =   18
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Return :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   5760
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Departure :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   6360
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF80&
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Trip Type :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   735
         Left            =   5520
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Registration Number :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Enter Your Registration  Number Here :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
con.Open
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Please Enter Your Flight Registration Number", vbCritical, "Airline Reservation System"
        Text1.SetFocus
        con.Close
    ElseIf Text1.Text Like "ARSDOW*" Then
        rst5.Open "select count(*) from domone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            If rst5(0) > 0 Then
                MsgBox "Your Flight was Successfully booked", vbInformation, "Airline Reservation System"
                rst5.Close
                    If MsgBox("Do You Want To Check The Details", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
                        Frame1.Visible = True
                        Command1.Enabled = False
                        Text1.Locked = True
                    End If
                rst.Open "select * from domone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text2.Text = rst!regno
                Text3.Text = rst!triptype
                Text4.Text = rst!From
                Text5.Text = rst!to
                Text6.Text = rst!departure
                rst.Close
                rst1.Open "select * from confirmdomone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text8.Text = rst1!fname
                Text9.Text = rst1!fdetails
                rst1.Close
                con.Close
            Else:   MsgBox "Invalid Registration Number", vbCritical, "Airline Reservation System"
                    Text1.Text = Clear
                    Text1.SetFocus
                    rst5.Close
                    con.Close
            End If
    ElseIf Text1.Text Like "ARSDRT*" Then
        rst5.Open "select count(*) from domround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            If rst5(0) > 0 Then
                MsgBox "Your Flight was Successfully booked", vbInformation, "Airline Reservation System"
                rst5.Close
                    If MsgBox("Do You Want To Check The Details", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
                        Frame1.Visible = True
                        Command1.Enabled = False
                        Text1.Locked = True
                    End If
                rst.Open "select * from domround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text2.Text = rst!regno
                Text3.Text = rst!triptype
                Text4.Text = rst!From
                Text5.Text = rst!to
                Text6.Text = rst!departure
                Text7.Text = rst!return
                rst.Close
                rst1.Open "select * from confirmdomround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text8.Text = rst1!fname
                Text9.Text = rst1!fdetails
                rst1.Close
                con.Close
                Label7.Visible = True
                Text7.Visible = True
            Else: MsgBox "Invalid Registration Number", vbCritical, "Airline Reservation System"
                Text1.Text = Clear
                Text1.SetFocus
                rst5.Close
                con.Close
            End If
    ElseIf Text1.Text Like "ARSIOW*" Then
        rst5.Open "select count(*) from intone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            If rst5(0) > 0 Then
                MsgBox "Your Flight was Successfully booked", vbInformation, "Airline Reservation System"
                rst5.Close
                    If MsgBox("Do You Want To Check The Details", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
                        Frame1.Visible = True
                        Command1.Enabled = False
                        Text1.Locked = True
                    End If
                rst.Open "select * from intone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text2.Text = rst!regno
                Text3.Text = rst!triptype
                Text4.Text = rst!From
                Text5.Text = rst!to
                Text6.Text = rst!departure
                rst.Close
                rst1.Open "select * from confirmintone where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text8.Text = rst1!fname
                Text9.Text = rst1!fdetails
                rst1.Close
                con.Close
            Else: MsgBox "Invalid Registration Number", vbCritical, "Airline Reservation System"
                Text1.Text = Clear
                Text1.SetFocus
                rst5.Close
                con.Close
            End If
    ElseIf Text1.Text Like "ARSIRT*" Then
        rst5.Open "select count(*) from intround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            If rst5(0) > 0 Then
                MsgBox "Your Flight was Successfully booked", vbInformation, "Airline Reservation System"
                rst5.Close
                    If MsgBox("Do You Want To Check The Details", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
                        Frame1.Visible = True
                        Command1.Enabled = False
                        Text1.Locked = True
                    End If
                rst.Open "select * from intround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text2.Text = rst!regno
                Text3.Text = rst!triptype
                Text4.Text = rst!From
                Text5.Text = rst!to
                Text6.Text = rst!departure
                Text7.Text = rst!return
                rst.Close
                rst1.Open "select * from confirmintround where regno = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
                Text8.Text = rst1!fname
                Text9.Text = rst1!fdetails
                rst1.Close
                con.Close
                Label7.Visible = True
                Text7.Visible = True
            Else: MsgBox "Invalid Registration Number", vbCritical, "Airline Reservation System"
                Text1.Text = Clear
                Text1.SetFocus
                rst5.Close
                con.Close
            End If
    Else:   MsgBox "Invalid Registration Number", vbCritical, "Airline Reservation System"
            Text1.Text = Clear
            Text1.SetFocus
            con.Close
    End If
End Sub

Private Sub Form_Load()
con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
End Sub
