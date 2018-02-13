VERSION 5.00
Begin VB.Form frmSplash10 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10665
   ClipControls    =   0   'False
   Icon            =   "frmSplash10.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About British Airways"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   3240
         Picture         =   "frmSplash10.frx":000C
         ScaleHeight     =   1035
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash10.frx":129B
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   10215
      End
   End
End
Attribute VB_Name = "frmSplash10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

