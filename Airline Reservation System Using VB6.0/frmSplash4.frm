VERSION 5.00
Begin VB.Form frmSplash4 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   10290
   ClipControls    =   0   'False
   Icon            =   "frmSplash4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Indigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   3240
         Picture         =   "frmSplash4.frx":000C
         ScaleHeight     =   1155
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   $"frmSplash4.frx":2054
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmSplash4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

