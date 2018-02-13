VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C0C000&
   Caption         =   "Update Your Details"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Main Form"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFF00&
      Caption         =   "Log Out"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
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
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text11 
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
      IMEMode         =   3  'DISABLE
      Left            =   7200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text10 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox Text12 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3120
         Width           =   4575
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   9000
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87490561
         CurrentDate     =   42024
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   495
         Left            =   7680
         TabIndex        =   25
         Top             =   5520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   32768
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   495
         Left            =   7080
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   32768
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
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
         Left            =   7080
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   873
         _Version        =   393216
         ForeColor       =   32768
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??????????????????????????????"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00FFFF00&
         Caption         =   "Update My Profile"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6240
         Width           =   2295
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
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   4455
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
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text17 
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
         Left            =   7080
         TabIndex        =   23
         Top             =   4320
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text15 
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
         Left            =   7080
         TabIndex        =   24
         Top             =   4920
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text13 
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
         Left            =   7080
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Modify"
         Enabled         =   0   'False
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1455
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5520
         Width           =   4575
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   4575
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Edit"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text9 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4320
         Width           =   4575
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3720
         Width           =   4575
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4920
         Width           =   4575
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   4575
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   4575
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
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date of Birth :"
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
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Username Can't Be Changed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7560
         TabIndex        =   56
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+91"
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
         Left            =   7080
         TabIndex        =   55
         Top             =   5520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Security Answer :"
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
         Left            =   120
         TabIndex        =   52
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Security Question :"
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
         Left            =   120
         TabIndex        =   51
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Contact Number :"
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
         Left            =   120
         TabIndex        =   50
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "E-Mail ID :"
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
         Left            =   120
         TabIndex        =   49
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gender :"
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
         Left            =   120
         TabIndex        =   48
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Last Name :"
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
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "First Name :"
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
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Type Your Password :"
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
      Left            =   3600
      TabIndex        =   54
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Type Your Username :"
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
      Left            =   3600
      TabIndex        =   53
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
con.Open
rst1.Open "select * from aircust where uname = ' " & Text10.Text & " ' and upass = ' " & Text11.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst2.Open "select count(*) from aircust where uname=' " & Text10.Text & " ' and upass=' " & Text11.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst3.Open "select uname,upass from aircust where uname=' " & Text10.Text & " ' and upass=' " & Text11.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
rst4.Open "select count(*) from aircust where uname=' " & Text10.Text & " '", con, adOpenDynamic, adLockOptimistic, adCmdText
    If Len(Trim(Text10.Text)) = 0 And Len(Trim(Text11.Text)) = 0 Then
        MsgBox "Please Enter Username and Password", vbCritical, "Airline Reservation System"
        Text10.SetFocus
        rst1.Close
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
    ElseIf Len(Trim(Text10.Text)) = 0 Then
        MsgBox "Please Enter Username", vbCritical, "Airline Reservation System"
        Text10.Text = Clear
        Text10.SetFocus
        rst1.Close
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
    ElseIf Len(Trim(Text11.Text)) = 0 Then
        MsgBox "Please Enter Password", vbCritical, "Airline Reservation System"
        Text11.Text = Clear
        Text11.SetFocus
        rst1.Close
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
    ElseIf Val(rst4(0)) > 0 Then
        If Val(rst2(0)) = 0 Then
            MsgBox "Invalid Password", vbExclamation, "Airline Reservation System"
            rst1.Close
            rst2.Close
            rst3.Close
            rst4.Close
            con.Close
            Text11.Text = Clear
            Text11.SetFocus
        Else: Frame1.Visible = True
            Text1.SetFocus
            Text1.Text = Trim(rst1!ufirst)
            Text2.Text = Trim(rst1!ulast)
            Text3.Text = Trim(rst1!uname)
            Text4.Text = Trim(rst1!upass)
            Text5.Text = Trim(rst1!ugender)
            Text6.Text = Trim(rst1!umail)
            Text7.Text = Trim(rst1!uphone)
            Text8.Text = Trim(rst1!usec)
            Text9.Text = Trim(rst1!uans)
            Text12.Text = Trim(rst1!udob)
            rst1.Close
            rst2.Close
            rst3.Close
            rst4.Close
            con.Close
            Command1.Enabled = False
            Text10.Locked = True
            Text11.Locked = True
        End If
    ElseIf Val(rst4(0)) = 0 Then
        MsgBox "Invalid Username", vbExclamation, "Airline Reservation System"
        rst1.Close
        rst2.Close
        rst3.Close
        rst4.Close
        con.Close
        Text10.Text = Clear
        Text10.SetFocus
    End If
End Sub

Private Sub Command10_Click()
Text17.Visible = True
Text17.SetFocus
Command19.Enabled = True
Command10.Enabled = False
End Sub

Private Sub Command11_Click()
Text1.Text = Trim(MaskEdBox1.Text)
Command11.Enabled = False
Command2.Enabled = True
Text1.SetFocus
MaskEdBox1.Visible = False
End Sub

Private Sub Command12_Click()
Text2.Text = Trim(MaskEdBox2.Text)
Command3.Enabled = True
Command12.Enabled = False
Text2.SetFocus
MaskEdBox2.Visible = False
End Sub

Private Sub Command13_Click()
DTPicker1.Visible = True
DTPicker1.SetFocus
Command22.Enabled = True
Command13.Enabled = False
DTPicker1.Value = Trim(Text12.Text)
End Sub

Private Sub Command14_Click()
Text4.Text = Trim(Text13.Text)
Command5.Enabled = True
Command14.Enabled = False
Text4.SetFocus
Text13.Visible = False
End Sub

Private Sub Command15_Click()
Text5.Text = Trim(Combo1.Text)
Command6.Enabled = True
Command15.Enabled = False
Text5.SetFocus
Combo1.Visible = False
End Sub

Private Sub Command16_Click()
Text6.Text = Trim(Text15.Text)
Command16.Enabled = False
Command7.Enabled = True
Text6.SetFocus
Text15.Visible = False
End Sub

Private Sub Command17_Click()
Text7.Text = Label12.Caption & MaskEdBox3.Text
Command8.Enabled = True
Command17.Enabled = False
Text7.SetFocus
MaskEdBox3.Visible = False
Label12.Visible = False
End Sub

Private Sub Command18_Click()
Text8.Text = Trim(Combo2.Text)
Command9.Enabled = True
Command18.Enabled = False
Text8.SetFocus
Combo2.Visible = False
End Sub

Private Sub Command19_Click()
Text9.Text = Trim(Text17.Text)
Command10.Enabled = True
Command19.Enabled = False
Text9.SetFocus
Text17.Visible = False
End Sub

Private Sub Command2_Click()
Command11.Enabled = True
Command2.Enabled = False
MaskEdBox1.Visible = True
MaskEdBox1.SetFocus
End Sub

Private Sub Command20_Click()
    If MsgBox("Are You Sure You Want To Log Out ? ", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
        MsgBox "Please Login To Continue", vbInformation, "Airline Reservation System"
        Form1.Show
        Unload Me
    End If
End Sub

Private Sub Command21_Click()
con.Open
    If Len(Trim(Text1.Text)) > 0 And Len(Trim(Text2.Text)) > 0 And Len(Trim(Text3.Text)) > 0 And Len(Text4.Text) > 0 And Len(Trim(Text5.Text)) > 0 And Len(Trim(Text6.Text)) > 0 And Len(Trim(Text7.Text)) > 0 And Len(Trim(Text8.Text)) > 0 And Len(Trim(Text9.Text)) > 0 Then
        If StrComp(DTPicker1.Value, Format(Now, "dd-mm-yyyy")) > 0 Then
            MsgBox "Invalid Date of Birth", vbCritical, "Airline Reservation System"
            con.Close
        ElseIf StrComp(Trim(Text7.Text), "+91") = 0 Then
            MsgBox "Please Input the Contact Number !!!", vbCritical, "Airline Reservation System"
            con.Close
        ElseIf MsgBox("Are You Sure You Want To Save The Changes", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes Then
            con.Execute ("update aircust set ufirst = ' " & Text1.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set ulast = ' " & Text2.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set upass = ' " & Text4.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set ugender = ' " & Text5.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set udob = ' " & Text12.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set usec = ' " & Text8.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set uans = ' " & Text9.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set umail = ' " & Text6.Text & " ' where uname = ' " & Text3.Text & " '")
            con.Execute ("update aircust set uphone = ' " & Text7.Text & " ' where uname = ' " & Text3.Text & " '")
            MsgBox "Profile Successfully Updated", vbInformation, "Airline Reservation System"
            Frame1.Visible = False
            Command1.Enabled = True
            Text10.Locked = False
            Text11.Locked = False
            con.Close
        Else: con.Close
        End If
    Else: MsgBox "You Can't Leave Any Mandatory Field Blank", vbCritical, "Airline Reservation System"
          con.Close
    End If
End Sub

Private Sub Command22_Click()
Text12.Text = DTPicker1.Value
Command13.Enabled = True
Command22.Enabled = False
Text12.SetFocus
DTPicker1.Visible = False
End Sub

Private Sub Command3_Click()
MaskEdBox2.Visible = True
MaskEdBox2.SetFocus
Command12.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Form32.Show
Unload Me
End Sub

Private Sub Command5_Click()
Text13.Visible = True
Text13.SetFocus
Command14.Enabled = True
Command5.Enabled = False
End Sub

Private Sub Command6_Click()
Combo1.Visible = True
Combo1.SetFocus
Command15.Enabled = True
Command6.Enabled = False
End Sub

Private Sub Command7_Click()
Text15.Visible = True
Text15.SetFocus
Command16.Enabled = True
Command7.Enabled = False
End Sub

Private Sub Command8_Click()
Label12.Visible = True
MaskEdBox3.Visible = True
MaskEdBox3.SetFocus
Command17.Enabled = True
Command8.Enabled = False
End Sub

Private Sub Command9_Click()
Combo2.Visible = True
Combo2.SetFocus
Command18.Enabled = True
Command9.Enabled = False
End Sub

Private Sub Form_Activate()
DTPicker1.Value = Format(Now, "dd-mm-yyyy")
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\airticket.mdb;"
con.Open
rst1.Open "select * from q", con, adOpenDynamic, adLockOptimistic, adCmdText
rst1.MoveFirst
While rst1.EOF <> True
Combo2.AddItem rst1(0)
rst1.MoveNext
Wend
rst1.Close
con.Close
Combo1.AddItem "Male"
Combo1.AddItem "Female"
End Sub

