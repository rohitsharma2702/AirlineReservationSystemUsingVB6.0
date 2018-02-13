VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   Caption         =   "Sign Up Form"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "<-  Back"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   15375
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Use Original Name"
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
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Use Original Name"
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
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   9360
         TabIndex        =   9
         Top             =   4080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85983233
         CurrentDate     =   42024
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   12360
         TabIndex        =   6
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   540
         Left            =   5880
         TabIndex        =   5
         Top             =   2880
         Width           =   6015
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   495
         Left            =   5880
         TabIndex        =   2
         Top             =   720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   32768
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??????????????????????????????"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   495
         Left            =   5880
         TabIndex        =   1
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   32768
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??????????????????????????????"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   495
         Left            =   6600
         TabIndex        =   7
         Top             =   3480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   255
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   540
         Left            =   5880
         TabIndex        =   11
         Top             =   5280
         Width           =   6015
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4680
         Width           =   6015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Create My Account"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6720
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "I Hereby Agree To The "
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
         Height          =   615
         Left            =   1080
         TabIndex        =   12
         Top             =   6000
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         IMEMode         =   3  'DISABLE
         Left            =   5880
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox Text3 
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   5880
         TabIndex        =   3
         Top             =   1440
         Width           =   6015
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080FFFF&
         Height          =   135
         Left            =   4200
         TabIndex        =   36
         Top             =   6000
         Width           =   2775
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080FFFF&
         Height          =   135
         Left            =   6960
         TabIndex        =   35
         Top             =   6000
         Width           =   5055
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080FFFF&
         Caption         =   "of The Airlines Department."
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6960
         TabIndex        =   34
         Top             =   6120
         Width           =   5055
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FFFF&
         Caption         =   "Terms and Conditions"
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
         Height          =   495
         Left            =   4200
         TabIndex        =   33
         Top             =   6120
         Width           =   2775
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   " +91  Will Automatically Be Prefixed and Saved With Your Contact Number"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   12000
         TabIndex        =   30
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Date of Birth :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   6840
         TabIndex        =   29
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   11880
         TabIndex        =   28
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* E-mail ID :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   12000
         TabIndex        =   26
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   12000
         TabIndex        =   25
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "+91"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   5880
         TabIndex        =   24
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "*Choose Your Security Answer :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   1200
         TabIndex        =   23
         Top             =   5280
         Width           =   4575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "*Choose Your Security Question :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   4680
         Width           =   4815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Contact Number :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Gender : "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   2400
         TabIndex        =   19
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Choose Your Password(Avoid Leading and Trailing Space) :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   735
         Left            =   1080
         TabIndex        =   18
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   " * Choose Your Username(Can't Be Changed Later and Must Be of 6 Characters) :"
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
         Height          =   615
         Left            =   600
         TabIndex        =   17
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Last Name(Maximum Length Upto 30 Characters) :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   960
         TabIndex        =   16
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "* First Name(Maximum Length Upto 30 Characters) :"
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
         Height          =   615
         Left            =   960
         TabIndex        =   15
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please Fill Your Necessary Details ( All Fields Are Mandatory ) :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Command1.Enabled = (Check1.Value = Checked)
End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex = 4 Then
Combo3.Text = Clear
End If
End Sub

Private Sub Command1_Click()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
    If Len(Trim(Combo1.Text)) > 0 And Len(Trim(Combo2.Text)) > 0 And Len(Trim(Combo3.Text)) > 0 And Len(Trim(MaskEdBox2.Text)) > 0 And Len(Trim(MaskEdBox3.Text)) > 0 And Len(Trim(Text3.Text)) > 0 And Len(Trim(Text4.Text)) > 0 And Len(Trim(MaskEdBox1.Text)) > 0 And Len(Trim(Text6.Text)) > 0 And Len(Trim(Text1.Text)) > 0 Then
        rst1.Open "select count(*) from aircust where uname=' " & Text3.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
            If rst1(0) > 0 Then
                MsgBox "Username is Not Available", vbCritical, "Airline Reservation System"
                Text3.SetFocus
                rst1.Close
                con.Close
            ElseIf Len(Trim(Text3.Text)) < 6 Then
                MsgBox "Username Must Be Atleast of 6 Characters", vbCritical, "Airline Reservation System"
                Text3.SetFocus
                rst1.Close
                con.Close
            ElseIf Len(Trim(Text4.Text)) < 8 Then
                MsgBox "Your Password is Very Weak", vbCritical, "Airline Reservation System"
                Text4.SetFocus
                rst1.Close
                con.Close
            ElseIf StrComp(Right(Trim(Combo3.Text), 4), ".com") <> 0 And StrComp(Right(Trim(Combo3.Text), 6), ".co.in") <> 0 Then
                MsgBox "E-mail ID should end with .com or .co.in", vbCritical, "Airline Reservation System"
                Combo3.SetFocus
                rst1.Close
                con.Close
            ElseIf Len(Trim(MaskEdBox1.Text)) <> 10 Then
                MsgBox "Contact Number Must Be of 10 Digits", vbCritical, "Airline Reservation System"
                MaskEdBox1.SetFocus
                rst1.Close
                con.Close
            ElseIf StrComp(DTPicker1.Value, Format(Now, "dd-mm-yyyy")) > 0 Then
                MsgBox "Invalid Date of Birth", vbCritical, "Airline Reservation System"
                DTPicker1.SetFocus
                rst1.Close
                con.Close
            Else: con.Execute ("insert into aircust values(' " & Trim(MaskEdBox2.Text) & " ',' " & Trim(MaskEdBox3.Text) & " ',' " & Trim(Text3.Text) & " ',' " & Text4.Text & " ',' " & Combo1.Text & " ',' " & DTPicker1.Value & " ',' " & Trim(Text1.Text) & Trim(Label14.Caption) & Trim(Combo3.Text) & " ',' " & Label8.Caption & Trim(MaskEdBox1.Text) & " ',' " & Combo2.Text & " ',' " & Trim(Text6.Text) & " ')")
                con.Close
                MsgBox "Sign Up Complete.", vbInformation, "Airline Reservation System"
                MsgBox "Username : " & Text3.Text & "            Password : " & Text4.Text, vbInformation, "Airline Reservation System"
                DTPicker1.Enabled = False
                Command1.Enabled = False
                Check1.Enabled = False
                MaskEdBox2.Enabled = False
                MaskEdBox3.Enabled = False
                MaskEdBox1.Enabled = False
                Text1.Text = Clear
                Text3.Text = Clear
                Text4.Text = Clear
                Text6.Text = Clear
                Combo1.Clear
                Combo2.Clear
                Combo3.Clear
                Label8.Caption = Clear
                Label14.Caption = Clear
                Label11.Visible = False
                Label12.Visible = False
            End If
    Else: MsgBox "You Can't Leave Any Mandatory Field Blank.", vbCritical, "Airline Reservation System"
        con.Close
    End If
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command3_Click()
MaskEdBox2.Text = LCase(Left(MaskEdBox2.Text, 1)) & Mid(MaskEdBox2.Text, 2)
End Sub

Private Sub Command4_Click()
MaskEdBox3.Text = LCase(Left(MaskEdBox3.Text, 1)) & Mid(MaskEdBox3.Text, 2)
End Sub

Private Sub Form_Activate()
DTPicker1.Value = Format(Now, "dd-mm-yyyy")
End Sub

Private Sub Form_Load()
Combo1.AddItem "Male"
Combo1.AddItem "Female"
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
rst.Open "select sec from q", con, adOpenDynamic, adLockOptimistic, adCmdText
rst.MoveFirst
While rst.EOF <> True
Combo2.AddItem rst(0)
rst.MoveNext
Wend
rst.Close
con.Close
Combo3.AddItem "gmail.com"
Combo3.AddItem "yahoo.com"
Combo3.AddItem "hotmail.com"
Combo3.AddItem "yandex.com"
Combo3.AddItem "Let Me Choose..."
End Sub

Private Sub Label17_Click()
frmSplash3.Show
End Sub

Private Sub MaskEdBox1_Change()
    If Len(Trim(MaskEdBox1.Text)) < 10 Then
        MaskEdBox1.ForeColor = &HFF&
    ElseIf Len(Trim(MaskEdBox1.Text)) = 10 Then
        MaskEdBox1.ForeColor = &H8000&
    End If
End Sub

Private Sub MaskEdBox2_LostFocus()
MaskEdBox2.Text = UCase(Left(MaskEdBox2.Text, 1)) & Mid(MaskEdBox2.Text, 2)
Command3.Visible = True
End Sub

Private Sub MaskEdBox3_LostFocus()
MaskEdBox3.Text = UCase(Left(MaskEdBox3.Text, 1)) & Mid(MaskEdBox3.Text, 2)
Command4.Visible = True
End Sub

Private Sub Text3_Change()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
rst.Open "select count(*) from aircust where uname = ' " & Text3.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
    If rst(0) > 0 Then
        Label12.Caption = "Username Not Available"
        Label12.ForeColor = &HFF&
        Text3.ForeColor = &HFF&
    Else: Label12.Caption = "Username Available"
        If Len(Trim(Text3.Text)) < 6 Then
            Label12.Caption = "Username Must Be Atleast of 6 Characters"
            Label12.ForeColor = &HFF&
            Text3.ForeColor = &HFF&
        Else: Label12.ForeColor = &H8000&
              Text3.ForeColor = &H8000&
        End If
    End If
rst.Close
con.Close
End Sub

Private Sub Text3_GotFocus()
    If Len(Trim(Text3.Text)) = 0 Then
        Label12.Caption = "*Make Sure That Your Username is Unique"
    End If
End Sub

Private Sub Text3_LostFocus()
    If Len(Trim(Text3.Text)) = 0 Then
        Label12.Caption = Clear
    End If
End Sub

Private Sub Text4_Change()
    If Len(Trim(Text4.Text)) > 10 Then
        Label11.Caption = "Strong Password"
        Label11.ForeColor = &H8000&
        Text4.ForeColor = &H8000&
    ElseIf Len(Trim(Text4.Text)) > 7 Then
        Label11.Caption = "Fair Password"
        Label11.ForeColor = &HFF0000
        Text4.ForeColor = &HFF0000
    Else: Label11.Caption = "Weak Password"
        Label11.ForeColor = &HFF&
        Text4.ForeColor = &HFF&
    End If
End Sub

Private Sub Text4_GotFocus()
    If Len(Trim(Text4.Text)) = 0 Then
        Label11.Caption = "Password Should Contain More Than 7 Characters"
    End If
End Sub

Private Sub Text4_LostFocus()
    If Len(Trim(Text4.Text)) = 0 Then
        Label11.Caption = Clear
    End If
End Sub
