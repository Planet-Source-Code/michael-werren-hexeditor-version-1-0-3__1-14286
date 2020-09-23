VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editor ASCII options"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton opt00B7 
         Caption         =   "Do not ask, use allways a 00  for a (·)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin VB.OptionButton opt00B7 
         Caption         =   "Allways ask about the 00 or B7 identification for a (·)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton opt00B7 
         Caption         =   "Do not ask, use allways a B7  for a (·)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  If opt00B7(0).Value Then
    giOpt00B7 = 0
  ElseIf opt00B7(1).Value Then
    giOpt00B7 = 1
  ElseIf opt00B7(2).Value Then
    giOpt00B7 = 2
  End If
End Sub

Private Sub cmdCancel_Click()
  frmOptions.Hide
End Sub

Private Sub cmdOK_Click()
  ' Save the  Informations
  If opt00B7(0).Value Then
    giOpt00B7 = 0
  ElseIf opt00B7(1).Value Then
    giOpt00B7 = 1
  ElseIf opt00B7(2).Value Then
    giOpt00B7 = 2
  End If
  frmOptions.Hide
End Sub

Private Sub Form_Activate()
  If giOpt00B7 = 0 Then
    opt00B7(0).Value = True
  ElseIf giOpt00B7 = 1 Then
    opt00B7(1).Value = True
  ElseIf giOpt00B7 = 2 Then
    opt00B7(2).Value = True
  End If

End Sub


