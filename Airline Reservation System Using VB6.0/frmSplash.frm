VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   5730
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   9480
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6600
         Top             =   600
      End
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   5640
         Top             =   480
      End
      Begin VB.PictureBox Picture1 
         Height          =   4695
         Left            =   720
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   4635
         ScaleWidth      =   2355
         TabIndex        =   10
         Top             =   -4455
         Width           =   2415
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   5040
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   1320
         Top             =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   4680
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Rohit Sharma"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Back End : Microsoft  Access 2007"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Front End : Microsoft Visual Basic 6.0"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Image imgLogo 
         Height          =   825
         Left            =   2520
         Picture         =   "frmSplash.frx":33EB
         Stretch         =   -1  'True
         Top             =   3195
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Warning : This Project is not for sale"
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
         Height          =   315
         Left            =   990
         TabIndex        =   2
         Top             =   4140
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblPlatform 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Platforms Used : "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2205
         TabIndex        =   3
         Top             =   1860
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Project  :  Aviation  Management  System"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   345
         TabIndex        =   4
         Top             =   1020
         Visible         =   0   'False
         Width           =   5325
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Developer :"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    If ProgressBar1.Value < 100 Then
        ProgressBar1.Value = ProgressBar1.Value + 1
            If ProgressBar1.Value < 30 Then
                Label5.Caption = "Preparing The Application..."
            ElseIf ProgressBar1.Value < 60 Then
                Label5.Caption = "Integrating With Database..."
            Else: Label5.Caption = "Loading The Necessary Files..."
            End If
    Else:   Form1.Show
            Unload Me
    End If
End Sub

Private Sub Timer2_Timer()
    If Picture1.Top < 720 Then
        Picture1.Top = Picture1.Top + 200
    Else: Timer3.Enabled = True
    End If
End Sub

Private Sub Timer3_Timer()
    If Picture1.Left < 6720 Then
        Picture1.Left = Picture1.Left + 200
    Else:   Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label5.Visible = True
            lblLicenseTo.Visible = True
            lblProductName.Visible = True
            lblPlatform.Visible = True
            lblWarning.Visible = True
            ProgressBar1.Visible = True
            imgLogo.Visible = True
            Timer1.Enabled = True
    End If
End Sub
