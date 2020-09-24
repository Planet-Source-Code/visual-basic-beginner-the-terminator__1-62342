VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About The Terminator"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2370
      TabIndex        =   1
      Top             =   0
      Width           =   2430
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "The Terminator can be useful for terminating adware/spyware that tries to prevent itself from being closed."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
