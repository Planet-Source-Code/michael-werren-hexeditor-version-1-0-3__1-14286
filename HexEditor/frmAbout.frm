VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4335
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2992.093
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   5535
      Begin VB.Label Label4 
         Caption         =   "Condition to download the source on www.planet-source-code.com, vote for the program."
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   $"frmAbout.frx":08CA
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Condition to work with this program. Send me a mail if you  like it, found any bug or you have an idee for next releases. "
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0962
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2235
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "The program is written with Visual Basic 6.0 (SP4)"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   960
      TabIndex        =   10
      Top             =   1200
      Width           =   4485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "mike@werren.com"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   4485
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Source by Michael Werren, January 2001"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   4485
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   4305
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   4245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About HexEditor"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = "HexEditor"
End Sub


