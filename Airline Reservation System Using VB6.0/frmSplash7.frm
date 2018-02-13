VERSION 5.00
Begin VB.Form frmSplash7 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6420
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10350
   ClipControls    =   0   'False
   Icon            =   "frmSplash7.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Air India"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   2880
         Picture         =   "frmSplash7.frx":000C
         ScaleHeight     =   915
         ScaleWidth      =   3555
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash7.frx":0EAA
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   9375
      End
   End
End
Attribute VB_Name = "frmSplash7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

