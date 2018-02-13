VERSION 5.00
Begin VB.Form frmSplash9 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6390
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10515
   ClipControls    =   0   'False
   Icon            =   "frmSplash9.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Emirates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   3120
         Picture         =   "frmSplash9.frx":000C
         ScaleHeight     =   915
         ScaleWidth      =   3915
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash9.frx":122F
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmSplash9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

