VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00808080&
   Caption         =   "Payment Gateway"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Methods"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   14175
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   1560
         TabIndex        =   11
         Top             =   1920
         Width           =   10815
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   4920
            MaxLength       =   3
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   6
            Top             =   2160
            Width           =   1335
         End
         Begin VB.PictureBox Picture1 
            Height          =   1935
            Left            =   6720
            Picture         =   "Form4.frx":0000
            ScaleHeight     =   1875
            ScaleWidth      =   3555
            TabIndex        =   16
            Top             =   1320
            Width           =   3615
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pay Now"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   4920
            MaxLength       =   4
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1440
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   495
            Left            =   3600
            TabIndex        =   4
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
            _Version        =   393216
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "################"
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "CVV*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   15
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Card Expiry Date*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PIN*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   13
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Card Number*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   12
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Card Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   8055
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Debit Card"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4920
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Credit Card"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Please Choose Your Payment Option :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Len(Trim(MaskEdBox1.Text)) > 0 And Len(Trim(Text1.Text)) > 0 And Len(Trim(Text2.Text)) > 0 And Len(Trim(Text3.Text)) > 0 Then
frmSplash13.Show
Else
MsgBox "All Fields Are Mandatory", vbCritical, "Something is Missing"
End If

End Sub


Private Sub Command3_Click()

Frame3.Enabled = True
Frame2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Command3.Enabled = False

    If Option1.Value = True Then
        Frame3.Caption = "Enter Credit Card Details"
    ElseIf Option2.Value = True Then
        Frame3.Caption = "Enter Debit Card Details"
    End If

End Sub

