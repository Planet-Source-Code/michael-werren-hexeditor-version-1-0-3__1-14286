VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find text"
   ClientHeight    =   3945
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ClipControls    =   0   'False
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   2775
      Begin VB.CheckBox optStartOnTop 
         Caption         =   "Start at first line"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox optFillB7 
         Caption         =   "Fill the String with B7 values"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox optMatchCase 
         Caption         =   "Match upper case"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search in"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   5775
      Begin VB.OptionButton optSearchIn 
         Caption         =   "ASCII"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
      Begin VB.OptionButton optSearchIn 
         Caption         =   "HEX Mode"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Mode"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   5775
      Begin VB.OptionButton optSearchTyp 
         Caption         =   "Expanded Search (Search the String  on more than one line."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
      Begin VB.OptionButton optSearchTyp 
         Caption         =   "Simple Search (Line Mode)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartButton 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   4695
      TabIndex        =   7
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Text you like to search:"
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
      TabIndex        =   10
      Top             =   480
      Width           =   2025
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancelButton_Click()
  frmFind.txtSearch.Text = ""
  frmFind.Hide
End Sub


Private Sub cmdStartButton_Click()
  Dim lsSearchText As String
  Dim liX As Integer
  Dim lbHexSearch As Boolean
  Dim lbMatchCase As Boolean
  
  Const cB7Value = 183
  Const cFindLine = 3
  
  If txtSearch.Text <> "" Then
    If optFillB7.Value = vbChecked Then
      lsSearchText = ""
      For liX = 1 To Len(txtSearch.Text)
        lsSearchText = lsSearchText + Mid(txtSearch.Text, liX, 1) + Chr(cB7Value)
      Next liX
      ' Cut the lase B7 Value
      lsSearchText = Left(lsSearchText, Len(lsSearchText) - 1)
    Else
      lsSearchText = txtSearch.Text
    End If
    
    ' Hex or ASCII Server
    If optSearchIn(0).Value = True Then
      lbHexSearch = False
    Else
      lbHexSearch = True
    End If
    
    If optMatchCase.Value = vbChecked Then
      lbMatchCase = True
    Else
      lbMatchCase = False
    End If
    
    ' Set the Start
    If optStartOnTop.Value = vbChecked Then
      glLastFound = 1
    Else
      glLastFound = frmMain.LstHexView.SelectedItem.Index
    End If
    
    ' Start the search
    If optSearchTyp(0).Value = True Then
      ' Simple Search
      glLastFound = SearchText(glLastFound, lsSearchText, lbHexSearch, lbMatchCase, 1)
      If glLastFound = -1 Then
        ' Notfound
        glLastFound = frmMain.LstHexView.SelectedItem.Index
      Else
        ' Found it
        With frmMain.LstHexView.ListItems.Item(glLastFound)
          .EnsureVisible
          .Selected = True
        End With
        glLastFound = glLastFound + 1
        frmFind.Hide
      End If
    Else
      ' Extended Search
      glLastFound = SearchText(glLastFound, lsSearchText, lbHexSearch, lbMatchCase, cFindLine)
      If glLastFound = -1 Then
        ' Notfound
        glLastFound = frmMain.LstHexView.SelectedItem.Index
      Else
        ' Found it
        With frmMain.LstHexView.ListItems.Item(glLastFound)
          .EnsureVisible
          .Selected = True
        End With
        glLastFound = glLastFound + 1
        frmFind.Hide
      End If
    End If
    SetMenu "Viewer"
  End If
End Sub

Private Sub Form_Activate()
  txtSearch.SetFocus
End Sub

Private Sub optSearchIn_GotFocus(Index As Integer)
  Select Case Index
    Case 0
      optFillB7.Enabled = True
      optMatchCase.Enabled = True
    Case 1
      optFillB7.Enabled = False
      optMatchCase.Enabled = False
  End Select
End Sub

Private Sub txtSearch_Change()
  glLastFound = 1
End Sub
