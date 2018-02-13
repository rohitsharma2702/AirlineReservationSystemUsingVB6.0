VERSION 5.00
Begin VB.Form Form62 
   Caption         =   "Help Form"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "Form62.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "You Can Exit The Project anytime by this menu."
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   7320
         Width           =   14895
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exit : "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   6960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Here you will get every guideline regarding the use of this project i.e. the working of the project and its benefits."
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   6360
         Width           =   14895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help : "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"Form62.frx":AAC85
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Width           =   14895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Feedback :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"Form62.frx":AAD99
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   14895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "About Developer : "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"Form62.frx":AAE46
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   14895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sign Up :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"Form62.frx":AAF2B
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   14895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "User Login :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Admin Login :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"Form62.frx":AB01C
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   14895
      End
   End
End
Attribute VB_Name = "Form62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

