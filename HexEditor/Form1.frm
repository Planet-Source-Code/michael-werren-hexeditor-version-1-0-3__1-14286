VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HexEditor"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9525
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageListOff 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2672
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3826
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4100
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":49DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6468
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProBar 
      Height          =   255
      Left            =   4500
      TabIndex        =   1
      Top             =   6300
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame FrameEditor 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      TabIndex        =   348
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
      Begin VB.VScrollBar VScrollEditor 
         Height          =   5175
         Left            =   9240
         Max             =   100
         Min             =   1
         TabIndex        =   369
         Top             =   0
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   342
         Top             =   4920
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   325
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   308
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   291
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   274
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   257
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   240
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   223
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   206
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   189
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   172
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   155
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   138
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   121
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   104
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   87
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   70
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   53
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtAscii 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7200
         MaxLength       =   16
         TabIndex        =   19
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   159
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   171
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   158
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   170
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   157
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   169
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   156
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   168
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   155
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   167
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   154
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   166
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   153
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   165
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   152
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   164
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   151
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   163
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   150
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   162
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   149
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   161
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   148
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   160
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   147
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   159
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   146
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   158
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   145
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   157
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   144
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   156
         Text            =   "00"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   143
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   154
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   142
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   153
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   141
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   152
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   140
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   151
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   139
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   150
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   138
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   149
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   137
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   148
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   136
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   147
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   135
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   146
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   134
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   145
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   133
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   144
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   132
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   143
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   131
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   142
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   130
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   141
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   129
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   140
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   128
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   139
         Text            =   "00"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   127
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   137
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   126
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   136
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   125
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   135
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   124
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   134
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   123
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   133
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   122
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   132
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   121
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   131
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   120
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   130
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   119
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   129
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   118
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   128
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   117
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   127
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   116
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   126
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   115
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   125
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   114
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   124
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   113
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   123
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   112
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   122
         Text            =   "00"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   111
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   120
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   110
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   119
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   109
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   118
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   108
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   117
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   107
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   116
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   106
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   115
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   105
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   114
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   104
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   113
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   103
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   112
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   102
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   111
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   101
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   110
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   100
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   109
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   99
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   108
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   98
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   107
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   97
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   106
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   96
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   105
         Text            =   "00"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   95
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   103
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   94
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   102
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   93
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   101
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   92
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   100
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   91
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   99
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   90
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   98
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   89
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   97
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   88
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   96
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   87
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   95
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   86
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   94
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   85
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   93
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   84
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   92
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   83
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   91
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   82
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   90
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   81
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   89
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   80
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   88
         Text            =   "00"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   79
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   86
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   78
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   85
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   77
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   84
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   76
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   83
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   75
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   82
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   74
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   81
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   73
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   80
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   72
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   79
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   71
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   78
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   70
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   77
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   69
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   76
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   68
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   75
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   67
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   74
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   66
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   73
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   65
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   72
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   64
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "00"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   63
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   69
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   62
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   68
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   61
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   60
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   66
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   59
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   65
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   58
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   64
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   57
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   63
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   56
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   62
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   55
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   61
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   54
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   60
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   59
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   51
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   50
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   49
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   48
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   54
         Text            =   "00"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   47
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   52
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   46
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   45
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   50
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   43
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   48
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   42
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   47
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   41
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   46
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   40
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   45
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   44
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   42
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "00"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   32
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   354
         Text            =   "00000000"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   355
         Text            =   "00000000"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   356
         Text            =   "00000000"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   357
         Text            =   "00000000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   358
         Text            =   "00000000"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   349
         Text            =   "00000000"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   350
         Text            =   "00000000"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   0
         TabIndex        =   351
         Text            =   "00000000"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   352
         Text            =   "00000000"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   353
         Text            =   "00000000"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   0
         TabIndex        =   368
         Text            =   "00000000"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   0
         TabIndex        =   367
         Text            =   "00000000"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   366
         Text            =   "00000000"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   365
         Text            =   "00000000"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   364
         Text            =   "00000000"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   0
         TabIndex        =   363
         Text            =   "00000000"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   362
         Text            =   "00000000"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   361
         Text            =   "00000000"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   360
         Text            =   "00000000"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtHexIndex 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   359
         Text            =   "00000000"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   319
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   341
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   318
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   340
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   317
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   339
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   316
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   338
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   315
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   337
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   314
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   336
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   313
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   335
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   312
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   334
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   311
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   333
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   310
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   332
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   309
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   331
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   308
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   330
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   307
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   329
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   306
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   328
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   305
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   327
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   304
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   326
         Text            =   "00"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   303
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   324
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   302
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   323
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   301
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   322
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   300
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   321
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   299
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   320
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   298
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   319
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   297
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   318
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   296
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   317
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   295
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   316
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   294
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   315
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   293
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   314
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   292
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   313
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   291
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   312
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   290
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   311
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   289
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   310
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   288
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   309
         Text            =   "00"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   287
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   307
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   286
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   306
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   285
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   305
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   284
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   304
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   283
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   303
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   282
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   302
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   281
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   301
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   280
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   300
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   279
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   299
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   278
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   298
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   277
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   297
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   276
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   296
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   275
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   295
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   274
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   294
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   273
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   293
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   272
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   292
         Text            =   "00"
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   271
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   290
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   270
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   289
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   269
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   288
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   268
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   287
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   267
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   286
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   266
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   285
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   265
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   284
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   264
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   283
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   263
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   282
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   262
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   281
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   261
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   280
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   260
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   279
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   259
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   278
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   258
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   277
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   257
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   276
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   256
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   275
         Text            =   "00"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   255
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   273
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   254
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   272
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   253
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   271
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   252
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   270
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   251
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   269
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   250
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   268
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   249
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   267
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   248
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   266
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   247
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   265
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   246
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   264
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   245
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   263
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   244
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   262
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   243
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   261
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   242
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   260
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   241
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   259
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   240
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   258
         Text            =   "00"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   239
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   256
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   238
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   255
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   237
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   254
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   236
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   253
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   235
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   252
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   234
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   251
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   233
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   250
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   232
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   249
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   231
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   248
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   230
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   247
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   229
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   246
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   228
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   245
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   227
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   244
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   226
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   243
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   225
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   242
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   224
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   241
         Text            =   "00"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   223
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   239
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   222
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   238
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   221
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   237
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   220
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   236
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   219
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   235
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   218
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   234
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   217
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   233
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   216
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   232
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   215
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   231
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   214
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   230
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   213
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   229
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   212
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   228
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   211
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   227
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   210
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   226
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   209
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   225
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   208
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   224
         Text            =   "00"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   207
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   222
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   206
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   221
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   205
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   220
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   204
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   219
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   203
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   218
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   202
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   217
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   201
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   216
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   200
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   215
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   199
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   214
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   198
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   213
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   197
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   212
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   196
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   211
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   195
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   210
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   194
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   209
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   193
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   208
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   192
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   207
         Text            =   "00"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   191
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   205
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   190
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   204
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   189
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   203
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   188
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   202
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   187
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   201
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   186
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   200
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   185
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   199
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   184
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   198
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   183
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   197
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   182
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   196
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   181
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   195
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   180
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   194
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   179
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   193
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   178
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   192
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   177
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   191
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   176
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   190
         Text            =   "00"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   175
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   188
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   174
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   187
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   173
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   186
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   172
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   185
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   171
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   184
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   170
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   183
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   169
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   182
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   168
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   181
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   167
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   180
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   166
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   179
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   165
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   178
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   164
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   177
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   163
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   176
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   162
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   175
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   161
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   174
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   160
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   173
         Text            =   "00"
         Top             =   2640
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D42
            Key             =   "Load"
            Object.Tag             =   "Load"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B94
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":89E6
            Key             =   "Editor"
            Object.Tag             =   "Editor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9838
            Key             =   "ReloadFrame"
            Object.Tag             =   "ReloadFrame"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A68A
            Key             =   "EditorOptions"
            Object.Tag             =   "EditorOptions"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B4DC
            Key             =   "FindPrevious"
            Object.Tag             =   "FindPrevious"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C32E
            Key             =   "Find"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D180
            Key             =   "FindNext"
            Object.Tag             =   "FindNext"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DFD2
            Key             =   "About"
            Object.Tag             =   "About"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E8AC
            Key             =   "StatusOn"
            Object.Tag             =   "StatusOn"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F6FE
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10550
            Key             =   "StatusOff"
            Object.Tag             =   "StatusOff"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageListOff"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Load"
            Object.ToolTipText     =   "Load File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editor"
            Object.ToolTipText     =   "set Editor or Viewer Modus"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReloadFrame"
            Object.ToolTipText     =   "Reload Frame"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   4550
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog DlgLoad 
      Left            =   600
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Display file in HexViewer"
      Filter          =   "All (*.*)|*.*"
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   347
      Top             =   7095
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7763
            MinWidth        =   7763
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Object.ToolTipText     =   "Show the file state of modification"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Object.ToolTipText     =   "Show the frame state of modification"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstHexView 
      Height          =   5175
      Left            =   0
      TabIndex        =   345
      Top             =   960
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "00 01 02 03 04 05 06 07 08  09 0A 0B 0C 0D 0E 0F"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ASCII"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.Image ImgGreen 
      Height          =   480
      Left            =   1560
      Picture         =   "Form1.frx":113A2
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgRed 
      Height          =   480
      Left            =   1080
      Picture         =   "Form1.frx":121E4
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label txtInfo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   100
      TabIndex        =   346
      Top             =   5760
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Index"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   120
      TabIndex        =   344
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ASCII"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7200
      TabIndex        =   343
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   5760
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuViewer 
      Caption         =   "&Viewer"
      Begin VB.Menu mnuViewerFind 
         Caption         =   "&Find"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewerNext 
         Caption         =   "Find &Next"
         Shortcut        =   +{F3}
      End
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "&Editor"
      Begin VB.Menu mnuEditorEdit 
         Caption         =   "&Edit"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditorReload 
         Caption         =   "&Reload Frame"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  ' set the Windowsize
  frmMain.Height = 7250

  ' Set global Setting
  giOpt00B7 = 0
  gbChangedFile = False
  gbChangedField = False
  StatusBar.Panels(3).Picture = ImgGreen.Picture
  StatusBar.Panels(4).Picture = ImgGreen.Picture
  SetMenu "Viewer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If UCase(frmMain.StatusBar.Panels(1).Text) = "EDITOR MODE" Then
    Cancel = True
  Else
    If gbChangedFile Then
      If MsgBox("File has been changed but not saved. You realy like to close HexEditor Application ?", vbQuestion + vbYesNo, "Close Application") = vbYes Then
        End
      Else
        Cancel = True
      End If
    Else
      If MsgBox("Close HexEdior Application ?", vbQuestion + vbYesNo, "Close Application") = vbYes Then
        End
      Else
        Cancel = True
      End If
    End If
  End If
End Sub


Private Sub LstHexView_DblClick()
  Call mnuEditorEdit_Click
End Sub

Private Sub mnuEditorEdit_Click()
    
  Const cMaxRecord = 19 ' 0-19 = 20 Records
  
  If LstHexView.ListItems.Count <> 0 Then
    If mnuEditorEdit.Checked Then
      ' Set the Menu
      SetMenu "Viewer"
      ' check if data has change and must be saved
      If gbChangedField Then
        SaveBlock SearchIndex(txtHexIndex(0).Text)
        gbChangedFile = True
        gbChangedField = False
        StatusBar.Panels(4).Picture = ImgGreen.Picture
      End If
      SetIndex txtHexIndex(0).Text
      DispMsg "View mode"
      FrameEditor.Visible = False
    Else
      ' set the menu
      SetMenu "Editor"
      LoadHexBlock LstHexView.SelectedItem.Index, LstHexView.ListItems.Count
      
      FrameEditor.Left = 0
      FrameEditor.Top = 960
      FrameEditor.Visible = True
      frmMain.txtHex(0).SetFocus
    End If
    mnuEditorEdit.Checked = Not mnuEditorEdit.Checked
  Else
    Beep
  End If
End Sub

Private Sub mnuEditorReload_Click()
  If MsgBox("Changes will be lost, are you sure ?", vbQuestion + vbYesNo, "Reload Frame") = vbYes Then
    LoadHexBlock VScrollEditor.Value, LstHexView.ListItems.Count
  End If
End Sub

Private Sub mnuFileExit_Click()
  Call Form_Unload(True)
End Sub

Private Sub mnuFileInfo_Click()
  FrmFileInfo.Show vbModal, Me
End Sub

Private Sub mnuFileLoad_Click()
  Dim lsFileName As String
  Dim liX As Integer
  Dim lsTransText As String
  
  On Error GoTo ErrAboardLoad
  DlgLoad.DialogTitle = "Load the file..."
  DlgLoad.ShowOpen
  lsFileName = DlgLoad.FileName
  On Error GoTo 0
  
  
  ' Set the Main WindowName
  frmMain.Caption = "HexEditor [" + lsFileName + "]"
  LstHexView.Visible = False
  
  ' Load  the File
  DispMsg "Loading the file, please wait..."
  On Error GoTo ErrLoadFile
  liX = FreeFile
  Open lsFileName For Binary Access Read As #liX
    ' Store the Informations
    gtFileInfo.Name = lsFileName
    gtFileInfo.Size = LOF(liX)
    ' Load the file
    lsTransText = Space$(LOF(liX))
    Get #liX, , lsTransText
  Close #liX
  DoEvents
  On Error GoTo 0
  
  ' Clear the ListView
  frmMain.LstHexView.ListItems.Clear
  
  ' Translate the Text
  HexTranslate (lsTransText)
  lsTransText = ""
  
  ' Show the HexBox and set status, reset the find
  frmFind.txtSearch.Text = ""
  LstHexView.Visible = True
  StatusBar.Panels(3).Picture = ImgGreen.Picture
  StatusBar.Panels(4).Picture = ImgGreen.Picture
  
  GoTo ErrCont
ErrAboardLoad:
  GoTo ErrCont
ErrLoadFile:
  MsgBox "File is  to large..."
  GoTo ErrCont
ErrCont:
  On Error GoTo 0
End Sub

Private Sub mnuFileOptions_Click()
  frmOptions.Show vbModal, Me
End Sub

Private Sub mnuFileSave_Click()
  Dim lsFileName As String
  Dim liX As Integer
  Dim lsTransText As String
  
  On Error GoTo ErrAboardSave
  DlgLoad.Flags = cdlOFNOverwritePrompt
  DlgLoad.DialogTitle = "Save the file..."
  DlgLoad.ShowSave
  lsFileName = DlgLoad.FileName
  On Error GoTo 0
  
  
  ' Save the file
  On Error GoTo ErrSaveError
  liX = FreeFile
  Open lsFileName For Binary Access Write As #liX
    ' Build the output file (set B7 to  00)
    DispMsg "Build the file, please wait..."
    lsTransText = SaveTranslate
    DispMsg "Saving the file, please wait..."
    Put #liX, , lsTransText
  Close #liX
  On Error GoTo 0
 
 ' Set the Blockstate, set the message
  frmMain.StatusBar.Panels(3).Picture = frmMain.ImgGreen.Picture
  gbChangedFile = False
  DispMsg ""
  
  GoTo ErrSaveCont
  
ErrAboardSave:
  GoTo ErrSaveCont
ErrSaveError:
  MsgBox "Save error, file open error", vbCritical + vbOKOnly, "Save file"
  GoTo ErrSaveCont
ErrSaveCont:
  On Error GoTo 0
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewerFind_Click()
  If LstHexView.ListItems.Count <> 0 Then
    frmFind.Show vbModal, Me
  Else
    Beep
  End If
End Sub

Private Sub mnuViewerNext_Click()
  FindNext
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  If UCase(Button.Key) = "ABOUT" Then
    Call mnuHelpAbout_Click
  ElseIf UCase(Button.Key) = "EXIT" Then
    Call mnuFileExit_Click
  ElseIf UCase(Button.Key) = "EDITOR" Then
    Call mnuEditorEdit_Click
  ElseIf UCase(Button.Key) = "RELOADFRAME" Then
    Call mnuEditorReload_Click
  ElseIf UCase(Button.Key) = "LOAD" Then
    Call mnuFileLoad_Click
  ElseIf UCase(Button.Key) = "SAVE" Then
    Call mnuFileSave_Click
  ElseIf UCase(Button.Key) = "OPTIONS" Then
    Call mnuFileOptions_Click
  ElseIf UCase(Button.Key) = "FIND" Then
    Call mnuViewerFind_Click
  End If
End Sub
Private Sub txtAscii_Change(Index As Integer)
  If Len(txtAscii(Index).Text) = 16 Then
    txtAscii(Index).ForeColor = &H80000008
  Else
    txtAscii(Index).ForeColor = &HFF&
  End If
End Sub

Private Sub txtAscii_GotFocus(Index As Integer)
  gsOldValue = txtAscii(Index).Text
  If gbChangedField Then
    frmMain.StatusBar.Panels(4).Picture = frmMain.ImgRed.Picture
  Else
    frmMain.StatusBar.Panels(4).Picture = frmMain.ImgGreen.Picture
  End If
End Sub

Private Sub txtAscii_Validate(Index As Integer, Cancel As Boolean)
  Dim liX As Integer
  Dim liLine As Integer
  Dim liLinePos As Integer
  Dim lsChar As String
  Dim lsCharHex As String
  Dim lsMsgText As String
  
  If Len(txtAscii(Index).Text) = 16 Then
    If txtAscii(Index).Text = gsOldValue Then
      ' Set the color to black
      txtAscii(Index).ForeColor = &H80000008
      ' gbChangedField = False
    Else
      ' set the color to green
      txtAscii(Index).ForeColor = &H8000&
      gbChangedField = True
    End If
    ' Build the HEX String
    For liX = 1 To Len(txtAscii(Index).Text)
      lsChar = Mid(txtAscii(Index).Text, liX, 1)
      ' is the char a . or code 00
      liLinePos = Index * 16 + liX - 1
      If lsChar = "·" Then
        If txtHex(liLinePos).Text <> "00" And txtHex(liLinePos) <> "B7" Then
          If giOpt00B7 = 0 Then
            ' user like to be asked about 00 or B7
            lsMsgText = ""
            lsMsgText = lsMsgText + "On positon" + Str(liX) + " you have enter a '·'" + vbCrLf
            lsMsgText = lsMsgText + "What do you like to enter a '00' or 'B7'" + vbCrLf + vbCrLf
            lsMsgText = lsMsgText + "For a '00' click to Yes" + vbCrLf
            lsMsgText = lsMsgText + "For a 'B7' click to No"
            If MsgBox(lsMsgText, vbYesNo + vbQuestion, "What you like to enter...") = vbYes Then
              ' a code 0 selected by user
              lsCharHex = "00"
              txtHex(liLinePos).Text = lsCharHex
            Else
              ' a . selected by user
              lsCharHex = Hex(Asc(lsChar))
              If Len(lsCharHex) < 2 Then
                lsCharHex = "0" + lsCharHex
              End If
              txtHex(liLinePos).Text = lsCharHex
            End If
          ElseIf giOpt00B7 = 1 Then
            ' Allways use a 00 for a .
            lsCharHex = "00"
            txtHex(liLinePos).Text = lsCharHex
          ElseIf giOpt00B7 = 2 Then
            ' Allways use a B7 for a .
              lsCharHex = Hex(Asc(lsChar))
              If Len(lsCharHex) < 2 Then
                lsCharHex = "0" + lsCharHex
              End If
              txtHex(liLinePos).Text = lsCharHex
          End If
        End If
      Else
        If txtHex(liLinePos).Enabled Then
          lsCharHex = Hex(Asc(lsChar))
          If Len(lsCharHex) < 2 Then
            lsCharHex = "0" + lsCharHex
          End If
          txtHex(liLinePos).Text = lsCharHex
        End If
      End If
    Next liX
  Else
    ' The ASCII field is to short
    Beep
    txtAscii(Index).ForeColor = &HFF&
    Cancel = True
    ' txtAscii(Index).SetFocus
  End If

End Sub

Private Sub txtHex_GotFocus(Index As Integer)
  gsOldValue = txtHex(Index).Text
  If gbChangedField Then
    frmMain.StatusBar.Panels(4).Picture = frmMain.ImgRed.Picture
  Else
    frmMain.StatusBar.Panels(4).Picture = frmMain.ImgGreen.Picture
  End If
End Sub

Private Sub txtHex_Validate(Index As Integer, Cancel As Boolean)
  Dim liX As Integer
  Dim liLine As Integer
  Dim lsChar As String
  Dim lbCancel As Boolean
  
  If Len(txtHex(Index).Text) = 2 Then
    lbCancel = False
    For liX = 1 To Len(txtHex(Index))
      lsChar = UCase(Mid(txtHex(Index), liX, 1))
      Select Case Asc(lsChar)
        Case 48 To 57 ' 0..9
          If lbCancel = False Then lbCancel = False
        Case 65 To 70 ' A..F
          If lbCancel = False Then lbCancel = False
        Case Else
          Beep
          lbCancel = True
      End Select
    Next liX
    If lbCancel Then
      'txtHex(Index).SetFocus
      Cancel = True
      txtHex(Index).ForeColor = &HFF&
    Else
      txtHex(Index).Text = UCase(txtHex(Index).Text)
      If txtHex(Index).Text = gsOldValue Then
        ' set textcolor to black
        txtHex(Index).ForeColor = &H80000008
        ' gbChangedField = False
      Else
        ' set textcolor to green
        txtHex(Index).ForeColor = &H8000&
        gbChangedField = True
      End If
      
      ' Build the ASCII String
      liLine = Index \ 16
      txtAscii(liLine) = ""
      For liX = liLine * 16 To liLine * 16 + 15
        If txtHex(liX).Enabled Then
           If Hex2Dec(txtHex(liX).Text) = 0 Then
             lsChar = "·"
           Else
             lsChar = Chr(Hex2Dec(txtHex(liX).Text))
           End If
         Else
           lsChar = " "
         End If
         txtAscii(liLine) = txtAscii(liLine) + lsChar
      Next liX
    End If
  Else
    ' The HEX field is to short
    Beep
    txtHex(Index).ForeColor = &HFF&
    Cancel = True
  End If
End Sub

Private Sub VScrollEditor_Change()
  If gbChangedField Then
    SaveBlock SearchIndex(txtHexIndex(0).Text)
    gbChangedFile = True
    gbChangedField = False
  End If
  LoadHexBlock VScrollEditor.Value, LstHexView.ListItems.Count
End Sub

Private Sub VScrollEditor_Scroll()
  If gbChangedField Then
    SaveBlock SearchIndex(txtHexIndex(0).Text)
    gbChangedFile = True
    gbChangedField = False
  End If
  LoadHexBlock VScrollEditor.Value, LstHexView.ListItems.Count
End Sub
