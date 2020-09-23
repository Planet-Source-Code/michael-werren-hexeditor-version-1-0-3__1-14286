VERSION 5.00
Begin VB.Form FrmFileInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   ClipControls    =   0   'False
   Icon            =   "FrmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFileSize 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtFileName 
      Height          =   495
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Byte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Open File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   435
   End
End
Attribute VB_Name = "FrmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  FrmFileInfo.Hide
End Sub

Private Sub Form_Activate()
  txtFileName.Text = gtFileInfo.Name
  txtFileSize.Text = gtFileInfo.Size
End Sub

