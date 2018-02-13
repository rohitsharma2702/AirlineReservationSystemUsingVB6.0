VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C000&
   Caption         =   "Feedback Form"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   19695
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   3840
         TabIndex        =   7
         Top             =   1680
         Width           =   11055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Don't Submit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7440
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Submit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7440
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3360
         Width           =   14655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Your E-Mail ID :          ( Required)"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Start From Below :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Please Give Your Valuable Feedback . You Feedback will be helpful in upgradation of this Software ."
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   12855
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
    If Len(Trim(Text1.Text)) > 0 And Len(Trim(Text2.Text)) > 0 Then
        con.Execute "insert into feedback values(' " & Trim(Text2.Text) & " ' , ' " & Trim(Text1.Text) & " ' ) "
        con.Close
        MsgBox "Your Feedback is Stored Successfully.You Will Be Reverted Back As Soon As Possible.Thank You.", vbInformation, "Thank You"
        Text1.Text = Clear
        Text2.Text = Clear
    Else: MsgBox "Both Are Mandatory", vbCritical, "Sorry"
        con.Close
    End If
End Sub

Private Sub Command2_Click()
MsgBox "Your Feedback Was Not Submitted", vbCritical, "Failure"
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Text1_Change()
Command1.Enabled = Len(Trim(Text1.Text)) > 0
Command2.Enabled = Len(Trim(Text1.Text)) > 0
End Sub

Private Sub Text2_Change()
Command1.Enabled = Len(Trim(Text2.Text)) > 0
Command2.Enabled = Len(Trim(Text2.Text)) > 0
End Sub
