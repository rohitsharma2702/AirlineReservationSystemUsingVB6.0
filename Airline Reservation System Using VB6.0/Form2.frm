VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C000&
   Caption         =   "Select Your Journey Type"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "International"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   7680
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Trip Type"
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
         Height          =   1215
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFF80&
            Caption         =   "One-Way Trip"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFF80&
            Caption         =   "Round Trip"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   2880
            TabIndex        =   33
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFF00&
            Caption         =   "Confirm Trip Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Trip Details"
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
         Height          =   4575
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   6855
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   5520
            TabIndex        =   39
            Top             =   2880
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox Combo12 
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
            Height          =   360
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   840
            Width           =   2415
         End
         Begin VB.ComboBox Combo11 
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
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFF00&
            Caption         =   "Search For Flights"
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
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3720
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   85983233
            CurrentDate     =   41947
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   4200
            TabIndex        =   24
            Top             =   2040
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   85983233
            CurrentDate     =   41947
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   4800
            TabIndex        =   42
            Top             =   2880
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   3840
            TabIndex        =   38
            Top             =   2880
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Line Line14 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   3600
            X2              =   3480
            Y1              =   3120
            Y2              =   3240
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   3600
            X2              =   3480
            Y1              =   3120
            Y2              =   3000
         End
         Begin VB.Line Line12 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   3000
            X2              =   3615
            Y1              =   3120
            Y2              =   3135
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFF80&
            Caption         =   "Type This Number in The Box Provided To Continue"
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
            Height          =   615
            Left            =   120
            TabIndex        =   37
            Top             =   2880
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3600
            X2              =   3360
            Y1              =   1800
            Y2              =   2040
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3000
            X2              =   3240
            Y1              =   1560
            Y2              =   1320
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3000
            X2              =   3600
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3000
            X2              =   3600
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Height          =   975
            Left            =   2760
            Shape           =   2  'Oval
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFF80&
            Caption         =   "From"
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
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFF80&
            Caption         =   "To"
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
            Height          =   375
            Left            =   4200
            TabIndex        =   29
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFF80&
            Caption         =   "Departure"
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
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFF80&
            Caption         =   "Return"
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
            Height          =   375
            Left            =   4200
            TabIndex        =   27
            Top             =   1680
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20415
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Domestic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   7335
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Trip Details"
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
            Height          =   4575
            Left            =   240
            TabIndex        =   10
            Top             =   1680
            Visible         =   0   'False
            Width           =   6855
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   5520
               TabIndex        =   17
               Top             =   2880
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00FFFF00&
               Caption         =   "Search For Flights"
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
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   3720
               Width           =   2895
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   4320
               TabIndex        =   16
               Top             =   2040
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   85983233
               CurrentDate     =   41947
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   120
               TabIndex        =   15
               Top             =   2040
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   85983233
               CurrentDate     =   41947
            End
            Begin VB.ComboBox Combo2 
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
               Height          =   360
               Left            =   4320
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   840
               Width           =   2295
            End
            Begin VB.ComboBox Combo1 
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
               Height          =   360
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   4680
               TabIndex        =   41
               Top             =   2880
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               Visible         =   0   'False
               X1              =   3600
               X2              =   3480
               Y1              =   3120
               Y2              =   3240
            End
            Begin VB.Line Line10 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               Visible         =   0   'False
               X1              =   3600
               X2              =   3480
               Y1              =   3120
               Y2              =   3000
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               Visible         =   0   'False
               X1              =   3000
               X2              =   3600
               Y1              =   3120
               Y2              =   3120
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFFF80&
               Caption         =   "Type This Number in The Box Provided To Continue"
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
               Height          =   615
               Left            =   120
               TabIndex        =   36
               Top             =   2880
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   3720
               TabIndex        =   35
               Top             =   2880
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   3600
               X2              =   3360
               Y1              =   1800
               Y2              =   2040
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   3000
               X2              =   3360
               Y1              =   1560
               Y2              =   1320
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   3000
               X2              =   3600
               Y1              =   1800
               Y2              =   1800
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   3000
               X2              =   3600
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               Height          =   975
               Left            =   2760
               Shape           =   2  'Oval
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFFF80&
               Caption         =   "Return"
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
               Height          =   375
               Left            =   4320
               TabIndex        =   20
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFFF80&
               Caption         =   "Departure"
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
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFFF80&
               Caption         =   "To"
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
               Height          =   375
               Left            =   4320
               TabIndex        =   14
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFF80&
               Caption         =   "From"
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
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Trip Type"
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
            Height          =   1215
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   6855
            Begin VB.CommandButton Command3 
               BackColor       =   &H00FFFF00&
               Caption         =   "Confirm Trip Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   360
               Width           =   2175
            End
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFFF80&
               Caption         =   "Round Trip"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   2640
               TabIndex        =   8
               Top             =   480
               Width           =   975
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFFF80&
               Caption         =   "One-Way Trip"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   360
               TabIndex        =   7
               Top             =   480
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   4320
         TabIndex        =   1
         Top             =   120
         Width           =   8655
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFF00&
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
            Height          =   615
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF80&
            Caption         =   "International"
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
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF80&
            Caption         =   "Domestic"
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
            Height          =   375
            Left            =   480
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin VB.Menu mnuMainForm 
      Caption         =   "          Main &Form        "
   End
   Begin VB.Menu mnuLogOut 
      Caption         =   "          Log &Out          "
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As String
Private Sub Command1_Click()
    If Option1.Value = True Then
        Frame4.Enabled = True
        Frame6.Enabled = True
        Command1.Enabled = False
        Option2.Enabled = False
    ElseIf Option2.Value = True Then
        Frame3.Enabled = True
        Frame7.Enabled = True
        Command1.Enabled = False
        Option1.Enabled = False
    Else: MsgBox "Please Select The Type of Journey To Continue.", vbQuestion, "Airline Reservation System"
    End If
End Sub

Private Sub Command3_Click()
    If Option5.Value = True Then
        Frame8.Enabled = True
        DTPicker2.Enabled = False
        Command3.Enabled = False
        Option6.Enabled = False
        Frame2.Enabled = False
        Frame6.Enabled = False
        Label10.Visible = True
        Line9.Visible = True
        Line10.Visible = True
        Line11.Visible = True
        Label9.Visible = True
        Text1.Visible = True
    ElseIf Option6.Value = True Then
        Frame8.Enabled = True
        DTPicker2.Enabled = True
        Command3.Enabled = False
        Option5.Enabled = False
        Frame2.Enabled = False
        Frame6.Enabled = False
        Label10.Visible = True
        Line9.Visible = True
        Line10.Visible = True
        Line11.Visible = True
        Label13.Visible = True
        Text1.Visible = True
    End If
End Sub

Private Sub Command4_Click()
b = Text2.Text
    If Option7.Value = True Then
        If (MsgBox("Are You Sure You Want To Save Your Details and Continue ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            If StrComp(Combo11.Text, Combo12.Text) = 0 Or Len(Combo11.Text) = 0 Or Len(Combo12.Text) = 0 Then
                MsgBox "Both The Places Can't Be Same or Left Blanked", vbCritical, "Airline Reservation System"
            ElseIf StrComp(DTPicker3.Value, Format(Now, "dd-mm-yyyy")) < 0 Then
                MsgBox "Historical Date in Departure Field is Not Allowed", vbCritical, "Airline Reservation System"
            ElseIf Len(Trim(Text2.Text)) = 0 Then
              MsgBox "Please Type The Number", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Label12.Caption, Text2.Text) <> 0 Then
                MsgBox "Incorrect Number Typed", vbCritical, "Airline Reservation System"
                Text2.Text = Clear
                Text2.SetFocus
            Else: con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
                  con.Open
                  con.Execute ("insert into intone values(' " & Text2.Text & " ',' ARSIOW' & '" & b & " ',' " & Option7.Caption & " ' , ' " & Combo11.Text & " ' , ' " & Combo12.Text & " ' , ' " & DTPicker3.Value & " ')")
                  con.Close
                  Form9.Show
                  Unload Me
            End If
        Else: MsgBox "Your Details Were Not Submitted.", vbCritical, "Airline Reservation System"
        End If
    ElseIf Option3.Value = True Then
        If (MsgBox("Are You Sure You Want To Save Your Details and Continue ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            If StrComp(Combo11.Text, Combo12.Text) = 0 Or Len(Combo11.Text) = 0 Or Len(Combo12.Text) = 0 Then
                MsgBox "Both The Places Can't Be Same or Left Blanked", vbCritical, "Airline Reservation System"
            ElseIf StrComp(DTPicker3.Value, Format(Now, "dd-mm-yyyy")) < 0 Then
                MsgBox "Historical Date in Departure Field is Not Allowed", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Format(DTPicker4.Value, "dd-mm-yyyy"), Format(DTPicker3.Value, "dd-mm-yyyy")) < 0 Or StrComp(Format(DTPicker4.Value, "yyyy"), Format(DTPicker3.Value, "yyyy")) < 0 Then
                MsgBox "Return Date Can't Be Earlier Than Departure Date", vbCritical, "Airline Reservation System"
            ElseIf Len(Trim(Text2.Text)) = 0 Then
                MsgBox "Please Type The Number", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Label14.Caption, Text2.Text) <> 0 Then
                MsgBox "Incorrect Number Typed", vbCritical, "Airline Reservation System"
                Text2.Text = Clear
                Text2.SetFocus
            Else: con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
                  con.Open
                  con.Execute ("insert into intround values(' " & Text2.Text & " ',' ARSIRT' & '" & b & " ',' " & Option3.Caption & " ' , ' " & Combo11.Text & " ' , ' " & Combo12.Text & " ' , ' " & DTPicker3.Value & " ' , ' " & DTPicker4.Value & " ')")
                  con.Close
                  Form92.Show
                  Unload Me
            End If
        Else: MsgBox "Your Details Were Not Submitted.", vbCritical, "Airline Reservation System"
        End If
    End If
End Sub

Private Sub Command5_Click()
a = Text1.Text
    If Option5.Value = True Then
        If (MsgBox("Are You Sure You Want To Save Your Details and Continue ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            If StrComp(Combo1.Text, Combo2.Text) = 0 Or Len(Combo1.Text) = 0 Or Len(Combo2.Text) = 0 Then
               MsgBox "Both The Places Can't Be Same or Left Blanked", vbCritical, "Airline Reservation System"
            ElseIf StrComp(DTPicker1.Value, Format(Now, "dd-mm-yyyy")) < 0 Then
                MsgBox "Historical Date in Departure Field is Not Allowed", vbCritical, "Airline Reservation System"
            ElseIf Len(Trim(Text1.Text)) = 0 Then
                MsgBox "Please Type The Number", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Label9.Caption, Text1.Text) <> 0 Then
                MsgBox "Incorrect Number Typed", vbCritical, "Airline Reservation System"
                Text1.Text = Clear
                Text1.SetFocus
            Else: con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
                  con.Open
                  con.Execute ("insert into domone values(' " & Text1.Text & " ',' ARSDOW' & '" & a & " ',' " & Option5.Caption & " ' , ' " & Combo1.Text & " ' , ' " & Combo2.Text & " ' , ' " & DTPicker1.Value & " ') ")
                  con.Close
                  Form8.Show
                  Unload Me
            End If
        End If
    ElseIf Option6.Value = True Then
        If (MsgBox("Are You Sure You Want To Save Your Details and Continue ?", vbYesNo + vbQuestion, "Airline Reservation System") = vbYes) Then
            If StrComp(Combo1.Text, Combo2.Text) = 0 Or Len(Combo1.Text) = 0 Or Len(Combo2.Text) = 0 Then
               MsgBox "Both The Places Can't Be Same or Left Blanked", vbCritical, "Airline Reservation System"
            ElseIf StrComp(DTPicker1.Value, Format(Now, "dd-mm-yyyy")) < 0 Then
                MsgBox "Historical Date in Departure Field is Not Allowed", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Format(DTPicker2.Value, "dd-mm-yyyy"), Format(DTPicker1.Value, "dd-mm-yyyy")) < 0 Or StrComp(Format(DTPicker2.Value, "yyyy"), Format(DTPicker1.Value, "yyyy")) < 0 Then
                MsgBox "Return Date Can't Be Earlier Than Departure Date", vbCritical, "Airline Reservation System"
            ElseIf Len(Trim(Text1.Text)) = 0 Then
                MsgBox "Please Type The Number", vbCritical, "Airline Reservation System"
            ElseIf StrComp(Label13.Caption, Text1.Text) <> 0 Then
                MsgBox "Incorrect Number Typed", vbCritical, "Airline Reservation System"
                Text1.Text = Clear
                Text1.SetFocus
            Else: con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
                  con.Open
                  con.Execute ("insert into domround values(' " & Text1.Text & " ',' ARSDRT' & '" & a & " ',' " & Option6.Caption & " ' , ' " & Combo1.Text & " ' , ' " & Combo2.Text & " ' , ' " & DTPicker1.Value & " ' , ' " & DTPicker2.Value & " '  )")
                  con.Close
                  Form82.Show
                  Unload Me
            End If
        End If
    End If
End Sub

Private Sub Command7_Click()
If Option7.Value = True Then
Frame5.Visible = True
Frame5.Enabled = True
DTPicker4.Enabled = False
Command7.Enabled = False
Option3.Enabled = False
Label11.Visible = True
Line12.Visible = True
Line13.Visible = True
Line14.Visible = True
Label12.Visible = True
Text2.Visible = True
ElseIf Option3.Value = True Then
Frame5.Visible = True
Frame5.Enabled = True
DTPicker4.Enabled = True
Command7.Enabled = False
Option7.Enabled = False
Label11.Visible = True
Line12.Visible = True
Line13.Visible = True
Line14.Visible = True
Label14.Visible = True
Text2.Visible = True
End If
End Sub

Private Sub Form_Activate()
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame8.Visible = False
DTPicker1.Value = Format(Now, "dd-mm-yyyy")
DTPicker2.Value = Format(Now, "dd-mm-yyyy")
DTPicker3.Value = Format(Now, "dd-mm-yyyy")
DTPicker4.Value = Format(Now, "dd-mm-yyyy")
End Sub

Private Sub Form_Load()
con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\airticket.mdb"
con.Open
rst1.Open "select * from dom order by domcity", con, adOpenDynamic, adLockOptimistic, adCmdText
rst1.MoveFirst
While rst1.EOF <> True
Combo1.AddItem rst1(0)
Combo2.AddItem rst1(0)
rst1.MoveNext
Wend
rst1.Close
rst2.Open "select * from intt order by intcity", con, adOpenDynamic, adLockOptimistic, adCmdText
rst2.MoveFirst
While rst2.EOF <> True
Combo12.AddItem rst2(0)
rst2.MoveNext
Wend
rst2.Close
rst2.Open "select * from dom order by domcity", con, adOpenDynamic, adLockOptimistic, adCmdText
rst2.MoveFirst
While rst2.EOF <> True
Combo11.AddItem rst2(0)
rst2.MoveNext
Wend
rst2.Close
rst3.Open "select max(fno) from domone", con, adOpenDynamic, adLockOptimistic, adCmdText
Label9.Caption = 1 + rst3(0)
rst3.Close
rst3.Open "select max(fno) from domround", con, adOpenDynamic, adLockOptimistic, adCmdText
Label13.Caption = 1 + rst3(0)
rst3.Close
rst3.Open "select max(fno) from intone", con, adOpenDynamic, adLockOptimistic, adCmdText
Label12.Caption = 1 + rst3(0)
rst3.Close
rst3.Open "select max(fno) from intround", con, adOpenDynamic, adLockOptimistic, adCmdText
Label14.Caption = 1 + rst3(0)
rst3.Close
con.Close
End Sub

Private Sub mnuLogOut_Click()
    If MsgBox("Are You Sure You Want To Log Out ? ", vbQuestion + vbYesNo, "Airline Reservation System") = vbYes Then
        MsgBox "Please Login To Continue", vbInformation, "Airline Reservation System"
        Form1.Show
        Unload Me
    End If
End Sub

Private Sub mnuMainForm_Click()
Form32.Show
Unload Me
End Sub

Private Sub Option1_Click()
Frame4.Visible = True
Frame6.Visible = True
Frame8.Visible = True
Frame3.Visible = False
Frame4.Enabled = False
Frame6.Enabled = False
Frame8.Enabled = False
End Sub

Private Sub Option2_Click()
Frame3.Visible = True
Frame5.Visible = True
Frame7.Visible = True
Frame3.Enabled = False
Frame5.Enabled = False
Frame7.Enabled = False
Frame4.Visible = False
End Sub
