VERSION 5.00
Begin VB.Form frmSplash2 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6075
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3960
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Select Your Security Question :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Type Your Username :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Type Your Password To Save The Changes :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "New Answer :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Change Your New Security Answer By Entering Your Password Below"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
If Len(Trim(Text1.Text)) = 0 And Len(Trim(Combo1.Text)) = 0 And Len(Trim(Text2.Text)) = 0 And Len(Trim(Text3.Text)) = 0 Then
MsgBox "Please Fill Your Details First", vbCritical, "Airline Reservation System"
Text1.SetFocus
con.Close
ElseIf Len(Trim(Text1.Text)) = 0 Then
MsgBox "Please Enter Your Username", vbCritical, "Airline Reservation System"
Text1.SetFocus
con.Close
ElseIf Len(Trim(Combo1.Text)) = 0 Then
MsgBox "Please Select Your Security Question", vbCritical, "Airline Reservation System"
Combo1.SetFocus
con.Close
ElseIf Len(Trim(Text2.Text)) = 0 Then
MsgBox "Please Enter Your New Answer", vbCritical, "Airline Reservation System"
Text2.SetFocus
con.Close
ElseIf Len(Trim(Text3.Text)) = 0 Then
MsgBox "Please Enter Your Password", vbCritical, "Airline Reservation System"
Text2.SetFocus
con.Close
Else: rst2.Open "select count(*) from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst3.Open "select upass,usec,uans from aircust where uname = ' " & Text1.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
If rst2(0) = 0 Then
MsgBox "Invalid Username", vbCritical, "Airline Reservation System"
Text1.SetFocus
rst2.Close
rst3.Close
con.Close
ElseIf rst2(0) > 0 Then
If StrComp(Combo1.Text, Trim(rst3!usec)) <> 0 Then
MsgBox "Invalid Security Question", vbCritical, "Airline Reservation System"
Combo1.SetFocus
rst2.Close
rst3.Close
con.Close
ElseIf StrComp(Text3.Text, Trim(rst3!upass)) <> 0 Then
MsgBox "Invalid Password", vbCritical, "Airline Reservation System"
Text3.SetFocus
rst2.Close
rst3.Close
con.Close
Else: con.Execute "update aircust set uans=' " & Text2.Text & " ' where uname=' " & Text1.Text & " ' and upass=' " & Text3.Text & " ' and usec = ' " & Combo1.Text & " '"
MsgBox "Security Answer Changed Successfully", vbInformation, "Airline Reservation System"
con.Close
Me.Hide
Unload Me
End If
End If
End If
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
rst.Open "select sec from q", con, adOpenDynamic, adLockOptimistic, adCmdText
rst.MoveFirst
While rst.EOF <> True
Combo1.AddItem rst(0)
rst.MoveNext
Wend
rst.Close
con.Close
End Sub

