VERSION 5.00
Begin VB.Form frmSplash12 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10455
   ClipControls    =   0   'False
   Icon            =   "frmSplash12.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Etihad Airways"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   2880
         Picture         =   "frmSplash12.frx":000C
         ScaleHeight     =   1035
         ScaleWidth      =   4035
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash12.frx":147C
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   9975
      End
   End
End
Attribute VB_Name = "frmSplash12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

