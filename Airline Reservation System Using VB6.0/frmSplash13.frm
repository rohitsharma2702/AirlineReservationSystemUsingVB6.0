VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash13 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   990
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash13.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   120
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmSplash13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Frame1_Click()
    
End Sub

Private Sub Timer1_Timer()

If ProgressBar1.Value < 140 Then
    ProgressBar1.Value = ProgressBar1.Value + 1
    
        If ProgressBar1.Value < 40 Then
            Label1.Caption = "Please Wait..."
        ElseIf ProgressBar1.Value < 80 Then
            Label1.Caption = "Payment process is about to be completed..."
        ElseIf ProgressBar1.Value = 120 Then
            Label1.Caption = "Payment Successful..."
        End If
        
Else
     Unload Me
     Form32.Show
     Unload Form4
End If

End Sub
