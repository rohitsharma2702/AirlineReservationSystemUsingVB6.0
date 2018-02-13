VERSION 5.00
Begin VB.Form frmSplash11 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10455
   ClipControls    =   0   'False
   Icon            =   "frmSplash11.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Delta Airlines"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   3240
         Picture         =   "frmSplash11.frx":000C
         ScaleHeight     =   1035
         ScaleWidth      =   3195
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash11.frx":0EFD
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmSplash11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

