VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H008B4801&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Break The Code"
   ClientHeight    =   4905
   ClientLeft      =   3585
   ClientTop       =   1800
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   4905
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   2400
   End
   Begin VB.CommandButton cmdNextGame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next Game"
      Enabled         =   0   'False
      Height          =   270
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   270
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Game"
      Height          =   270
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   6
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdGuess1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guess"
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   3120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblTime1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008B4801&
      Caption         =   "0:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7515
      TabIndex        =   105
      Tag             =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblTitle5 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   104
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblTime2 
      BackColor       =   &H008B4801&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   103
      Tag             =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   840
      TabIndex        =   102
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   34
      Left            =   6120
      TabIndex        =   100
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   33
      Left            =   5640
      TabIndex        =   99
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   32
      Left            =   5160
      TabIndex        =   98
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   97
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   4200
      TabIndex        =   96
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   6120
      TabIndex        =   95
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   5640
      TabIndex        =   94
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   5160
      TabIndex        =   93
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   4680
      TabIndex        =   92
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   4200
      TabIndex        =   91
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   6120
      TabIndex        =   90
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   5640
      TabIndex        =   89
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   5160
      TabIndex        =   88
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   4680
      TabIndex        =   87
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4200
      TabIndex        =   86
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   6120
      TabIndex        =   85
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   84
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   5160
      TabIndex        =   83
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   4680
      TabIndex        =   82
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   4200
      TabIndex        =   81
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   80
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   79
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   5160
      TabIndex        =   78
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   77
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   76
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   75
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   74
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   73
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   72
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   71
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   70
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   69
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   68
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   67
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H008B4801&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   66
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   34
      Left            =   2760
      TabIndex        =   65
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   33
      Left            =   2280
      TabIndex        =   64
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   32
      Left            =   1800
      TabIndex        =   63
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   31
      Left            =   1320
      TabIndex        =   62
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   30
      Left            =   840
      TabIndex        =   61
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   29
      Left            =   2760
      TabIndex        =   60
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   28
      Left            =   2280
      TabIndex        =   59
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   27
      Left            =   1800
      TabIndex        =   58
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   26
      Left            =   1320
      TabIndex        =   57
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   25
      Left            =   840
      TabIndex        =   56
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   24
      Left            =   2760
      TabIndex        =   55
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   23
      Left            =   2280
      TabIndex        =   54
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   22
      Left            =   1800
      TabIndex        =   53
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   21
      Left            =   1320
      TabIndex        =   52
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   20
      Left            =   840
      TabIndex        =   51
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   19
      Left            =   2760
      TabIndex        =   50
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   18
      Left            =   2280
      TabIndex        =   49
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   17
      Left            =   1800
      TabIndex        =   48
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   16
      Left            =   1320
      TabIndex        =   47
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   15
      Left            =   840
      TabIndex        =   46
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   7680
      TabIndex        =   45
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblTitle4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   44
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H008B4801&
      Caption         =   "--->"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   42
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   41
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   1800
      TabIndex        =   40
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   39
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   38
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   37
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   36
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   35
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   34
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   33
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   31
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   30
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   29
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4801&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Height          =   255
      Index           =   7
      Left            =   8520
      TabIndex        =   27
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   26
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   7800
      TabIndex        =   25
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   24
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   22
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   20
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   3960
      X2              =   3960
      Y1              =   1080
      Y2              =   3960
   End
   Begin VB.Label lblTitle3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choices"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1440
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblCoverUpLeftTopCorner 
      BackColor       =   &H008B4801&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hints"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTitle1 
      BackStyle       =   0  'Transparent
      Caption         =   "Guess"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   3960
      X2              =   3960
      Y1              =   480
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   480
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "7.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "6.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "4.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Place
Dim Code1
Dim Code2
Dim Code3
Dim Code4
Dim Code5
Dim Win
Dim Thing

Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdGuess1_Click(Index As Integer)
    
    Win = 0
    
    Select Case Index
     
       Case "0"
            Thing = 0
            cmdGuess1(0).Enabled = False
            If Label3(0).Tag = Code1 Then
               lblHint(0).BackColor = Label3(0).BackColor
               Win = Win + 1
            Else
               lblHint(0).ForeColor = &HFFFF&
               lblHint(0).FontBold = True
               lblHint(0).Caption = "X"
            End If
            
            
             If Label3(1).Tag = Code2 Then
               lblHint(1).BackColor = Label3(1).BackColor
               Win = Win + 1
            Else
               lblHint(1).ForeColor = &HFFFF&
               lblHint(1).FontBold = True
               lblHint(1).Caption = "X"
            End If
            
             If Label3(2).Tag = Code3 Then
               lblHint(2).BackColor = Label3(2).BackColor
               Win = Win + 1
            Else
               lblHint(2).ForeColor = &HFFFF&
               lblHint(2).FontBold = True
               lblHint(2).Caption = "X"
            End If
            
             If Label3(3).Tag = Code4 Then
               lblHint(3).BackColor = Label3(3).BackColor
               Win = Win + 1
            Else
               lblHint(3).ForeColor = &HFFFF&
               lblHint(3).FontBold = True
               lblHint(3).Caption = "X"
            End If
            
             If Label3(4).Tag = Code5 Then
               lblHint(4).BackColor = Label3(4).BackColor
               Win = Win + 1
            Else
               lblHint(4).ForeColor = &HFFFF&
               lblHint(4).FontBold = True
               lblHint(4).Caption = "X"
            End If
            
            lblHint(0).Visible = True
            lblHint(1).Visible = True
            lblHint(2).Visible = True
            lblHint(3).Visible = True
            lblHint(4).Visible = True
               
      Case "1"
            Thing = 1
            cmdGuess1(1).Enabled = False
            If Label3(5).Tag = Code1 Then
               lblHint(5).BackColor = Label3(5).BackColor
               Win = Win + 1
            Else
               lblHint(5).ForeColor = &HFFFF&
               lblHint(5).FontBold = True
               lblHint(5).Caption = "X"
            End If
            
            
             If Label3(6).Tag = Code2 Then
               lblHint(6).BackColor = Label3(6).BackColor
               Win = Win + 1
            Else
               lblHint(6).ForeColor = &HFFFF&
               lblHint(6).FontBold = True
               lblHint(6).Caption = "X"
            End If
            
             If Label3(7).Tag = Code3 Then
               lblHint(7).BackColor = Label3(7).BackColor
               Win = Win + 1
            Else
               lblHint(7).ForeColor = &HFFFF&
               lblHint(7).FontBold = True
               lblHint(7).Caption = "X"
            End If
            
             If Label3(8).Tag = Code4 Then
               lblHint(8).BackColor = Label3(8).BackColor
               Win = Win + 1
            Else
               lblHint(8).ForeColor = &HFFFF&
               lblHint(8).FontBold = True
               lblHint(8).Caption = "X"
            End If
            
             If Label3(9).Tag = Code5 Then
               lblHint(9).BackColor = Label3(9).BackColor
               Win = Win + 1
            Else
               lblHint(9).ForeColor = &HFFFF&
               lblHint(9).FontBold = True
               lblHint(9).Caption = "X"
            End If
            
            lblHint(5).Visible = True
            lblHint(6).Visible = True
            lblHint(7).Visible = True
            lblHint(8).Visible = True
            lblHint(9).Visible = True
            
      Case "2"
            Thing = 2
            cmdGuess1(2).Enabled = False
            If Label3(10).Tag = Code1 Then
               lblHint(10).BackColor = Label3(10).BackColor
               Win = Win + 1
            Else
               lblHint(10).ForeColor = &HFFFF&
               lblHint(10).FontBold = True
               lblHint(10).Caption = "X"
            End If
            
            
             If Label3(11).Tag = Code2 Then
               lblHint(11).BackColor = Label3(11).BackColor
               Win = Win + 1
            Else
               lblHint(11).ForeColor = &HFFFF&
               lblHint(11).FontBold = True
               lblHint(11).Caption = "X"
            End If
            
             If Label3(12).Tag = Code3 Then
               lblHint(12).BackColor = Label3(12).BackColor
               Win = Win + 1
            Else
               lblHint(12).ForeColor = &HFFFF&
               lblHint(12).FontBold = True
               lblHint(12).Caption = "X"
            End If
            
             If Label3(13).Tag = Code4 Then
               lblHint(13).BackColor = Label3(13).BackColor
               Win = Win + 1
            Else
               lblHint(13).ForeColor = &HFFFF&
               lblHint(13).FontBold = True
               lblHint(13).Caption = "X"
            End If
            
             If Label3(14).Tag = Code5 Then
               lblHint(14).BackColor = Label3(14).BackColor
               Win = Win + 1
            Else
               lblHint(14).ForeColor = &HFFFF&
               lblHint(14).FontBold = True
               lblHint(14).Caption = "X"
            End If
            
            lblHint(10).Visible = True
            lblHint(11).Visible = True
            lblHint(12).Visible = True
            lblHint(13).Visible = True
            lblHint(14).Visible = True
      
       Case "3"
            Thing = 3
            cmdGuess1(3).Enabled = False
            If Label3(15).Tag = Code1 Then
               lblHint(15).BackColor = Label3(15).BackColor
               Win = Win + 1
            Else
               lblHint(15).ForeColor = &HFFFF&
               lblHint(15).FontBold = True
               lblHint(15).Caption = "X"
            End If
            
            
             If Label3(16).Tag = Code2 Then
               lblHint(16).BackColor = Label3(16).BackColor
               Win = Win + 1
            Else
               lblHint(16).ForeColor = &HFFFF&
               lblHint(16).FontBold = True
               lblHint(16).Caption = "X"
            End If
            
             If Label3(17).Tag = Code3 Then
               lblHint(17).BackColor = Label3(17).BackColor
               Win = Win + 1
            Else
               lblHint(17).ForeColor = &HFFFF&
               lblHint(17).FontBold = True
               lblHint(17).Caption = "X"
            End If
            
             If Label3(18).Tag = Code4 Then
               lblHint(18).BackColor = Label3(18).BackColor
               Win = Win + 1
            Else
               lblHint(18).ForeColor = &HFFFF&
               lblHint(18).FontBold = True
               lblHint(18).Caption = "X"
            End If
            
             If Label3(19).Tag = Code5 Then
               lblHint(19).BackColor = Label3(19).BackColor
               Win = Win + 1
            Else
               lblHint(19).ForeColor = &HFFFF&
               lblHint(19).FontBold = True
               lblHint(19).Caption = "X"
            End If
            
            lblHint(15).Visible = True
            lblHint(16).Visible = True
            lblHint(17).Visible = True
            lblHint(18).Visible = True
            lblHint(19).Visible = True
            
       Case "4"
            Thing = 4
            cmdGuess1(4).Enabled = False
            If Label3(20).Tag = Code1 Then
               lblHint(20).BackColor = Label3(20).BackColor
               Win = Win + 1
            Else
               lblHint(20).ForeColor = &HFFFF&
               lblHint(20).FontBold = True
               lblHint(20).Caption = "X"
            End If
            
            
             If Label3(21).Tag = Code2 Then
               lblHint(21).BackColor = Label3(21).BackColor
               Win = Win + 1
            Else
               lblHint(21).ForeColor = &HFFFF&
               lblHint(21).FontBold = True
               lblHint(21).Caption = "X"
            End If
            
             If Label3(22).Tag = Code3 Then
               lblHint(22).BackColor = Label3(22).BackColor
               Win = Win + 1
            Else
               lblHint(22).ForeColor = &HFFFF&
               lblHint(22).FontBold = True
               lblHint(22).Caption = "X"
            End If
            
             If Label3(23).Tag = Code4 Then
               lblHint(23).BackColor = Label3(23).BackColor
               Win = Win + 1
            Else
               lblHint(23).ForeColor = &HFFFF&
               lblHint(23).FontBold = True
               lblHint(23).Caption = "X"
            End If
            
             If Label3(24).Tag = Code5 Then
               lblHint(24).BackColor = Label3(24).BackColor
               Win = Win + 1
            Else
               lblHint(24).ForeColor = &HFFFF&
               lblHint(24).FontBold = True
               lblHint(24).Caption = "X"
            End If
            
            lblHint(20).Visible = True
            lblHint(21).Visible = True
            lblHint(22).Visible = True
            lblHint(23).Visible = True
            lblHint(24).Visible = True
            
       Case "5"
            Thing = 5
            cmdGuess1(5).Enabled = False
            If Label3(25).Tag = Code1 Then
               lblHint(25).BackColor = Label3(25).BackColor
               Win = Win + 1
            Else
               lblHint(25).ForeColor = &HFFFF&
               lblHint(25).FontBold = True
               lblHint(25).Caption = "X"
            End If
            
            
             If Label3(26).Tag = Code2 Then
               lblHint(26).BackColor = Label3(26).BackColor
               Win = Win + 1
            Else
               lblHint(26).ForeColor = &HFFFF&
               lblHint(26).FontBold = True
               lblHint(26).Caption = "X"
            End If
            
             If Label3(27).Tag = Code3 Then
               lblHint(27).BackColor = Label3(27).BackColor
               Win = Win + 1
            Else
               lblHint(27).ForeColor = &HFFFF&
               lblHint(27).FontBold = True
               lblHint(27).Caption = "X"
            End If
            
             If Label3(28).Tag = Code4 Then
               lblHint(28).BackColor = Label3(28).BackColor
               Win = Win + 1
            Else
               lblHint(28).ForeColor = &HFFFF&
               lblHint(28).FontBold = True
               lblHint(28).Caption = "X"
            End If
            
             If Label3(29).Tag = Code5 Then
               lblHint(29).BackColor = Label3(29).BackColor
               Win = Win + 1
            Else
               lblHint(29).ForeColor = &HFFFF&
               lblHint(29).FontBold = True
               lblHint(29).Caption = "X"
            End If
            
            lblHint(25).Visible = True
            lblHint(26).Visible = True
            lblHint(27).Visible = True
            lblHint(28).Visible = True
            lblHint(29).Visible = True
            
       Case "6"
            Thing = 6
            cmdGuess1(6).Enabled = False
            If Label3(30).Tag = Code1 Then
               lblHint(30).BackColor = Label3(30).BackColor
               Win = Win + 1
            Else
               lblHint(30).ForeColor = &HFFFF&
               lblHint(30).FontBold = True
               lblHint(30).Caption = "X"
            End If
            
            
             If Label3(31).Tag = Code2 Then
               lblHint(31).BackColor = Label3(31).BackColor
               Win = Win + 1
            Else
               lblHint(31).ForeColor = &HFFFF&
               lblHint(31).FontBold = True
               lblHint(31).Caption = "X"
            End If
            
             If Label3(32).Tag = Code3 Then
               lblHint(32).BackColor = Label3(32).BackColor
               Win = Win + 1
            Else
               lblHint(32).ForeColor = &HFFFF&
               lblHint(32).FontBold = True
               lblHint(32).Caption = "X"
            End If
            
             If Label3(33).Tag = Code4 Then
               lblHint(33).BackColor = Label3(33).BackColor
               Win = Win + 1
            Else
               lblHint(33).ForeColor = &HFFFF&
               lblHint(33).FontBold = True
               lblHint(33).Caption = "X"
            End If
            
             If Label3(34).Tag = Code5 Then
               lblHint(34).BackColor = Label3(34).BackColor
               Win = Win + 1
            Else
               lblHint(34).ForeColor = &HFFFF&
               lblHint(34).FontBold = True
               lblHint(34).Caption = "X"
            End If
            
            lblHint(30).Visible = True
            lblHint(31).Visible = True
            lblHint(32).Visible = True
            lblHint(33).Visible = True
            lblHint(34).Visible = True
            
       End Select
       
       If Index <> 6 Then
       Label6.Top = Label6.Top + 480
       Call Enable
       End If
       
       Call WinGame((Index))
       
End Sub

Private Sub cmdNewGame_Click()
    Timer1.Enabled = False
    Call cmdNextGame_Click
End Sub

Private Sub cmdNextGame_Click()
   Randomize
   
   Place = 0
   Win = 0
   
   Code1 = Int(Rnd * 7)
   Code2 = Int(Rnd * 7)
   Code3 = Int(Rnd * 7)
   Code4 = Int(Rnd * 7)
   Code5 = Int(Rnd * 7)
   
   Call Enable
   
   For I = 0 To 34
       Label3(I).BackColor = &H8B4801
       lblHint(I).BackColor = &H8B4801
       lblHint(I).Caption = ""
   Next I
   
   lblStatus.Caption = ""
   Label6.Move 120, 600
   
   lblTime1.Tag = 0
   lblTime2.Tag = 0
   lblTime1.Caption = "0:"
   lblTime2.Caption = "00"
   
End Sub

Private Sub Form_Load()

   Randomize
   
   Place = 0
   Win = 0
   
   Code1 = Int(Rnd * 7)
   Code2 = Int(Rnd * 7)
   Code3 = Int(Rnd * 7)
   Code4 = Int(Rnd * 7)
   Code5 = Int(Rnd * 7)
      
End Sub

Private Sub Label2_Click(Index As Integer)
    
 Select Case Index
 
   Case "0"
              Label3(Place).BackColor = &HC000&
              Label3(Place).Tag = 0
   Case "1"
              Label3(Place).BackColor = &HC0&
              Label3(Place).Tag = 1
   Case "2"
              Label3(Place).BackColor = &HC00000
              Label3(Place).Tag = 2
   Case "3"
              Label3(Place).BackColor = &HC0C0FF
              Label3(Place).Tag = 3
   Case "4"
              Label3(Place).BackColor = &HFF8080
              Label3(Place).Tag = 4
   Case "5"
              Label3(Place).BackColor = &H80FF&
              Label3(Place).Tag = 5
   Case "6"
              Label3(Place).BackColor = &HFFFF00
              Label3(Place).Tag = 6
   Case "7"
              Label3(Place).BackColor = &H800080
              Label3(Place).Tag = 7
   End Select
   
   If Place = 0 Then
   Timer1.Enabled = True
   End If
   
   
   Place = Place + 1
   
   If Place = "5" Then
     cmdGuess1(0).Enabled = True
     Call Disable
   ElseIf Place = "10" Then
     cmdGuess1(1).Enabled = True
     Call Disable
   ElseIf Place = "15" Then
     cmdGuess1(2).Enabled = True
     Call Disable
   ElseIf Place = "20" Then
     cmdGuess1(3).Enabled = True
     Call Disable
   ElseIf Place = "25" Then
     cmdGuess1(4).Enabled = True
     Call Disable
   ElseIf Place = "30" Then
     cmdGuess1(5).Enabled = True
     Call Disable
   ElseIf Place = "35" Then
     cmdGuess1(6).Enabled = True
     Call Disable
   ElseIf Place = "40" Then
     cmdGuess1(7).Enabled = True
     Call Disable
   End If
   
End Sub

Private Sub Disable()
     Label2(0).Enabled = False
     Label2(1).Enabled = False
     Label2(2).Enabled = False
     Label2(3).Enabled = False
     Label2(4).Enabled = False
     Label2(5).Enabled = False
     Label2(6).Enabled = False
     Label2(7).Enabled = False
End Sub

Private Sub Enable()
     Label2(0).Enabled = True
     Label2(1).Enabled = True
     Label2(2).Enabled = True
     Label2(3).Enabled = True
     Label2(4).Enabled = True
     Label2(5).Enabled = True
     Label2(6).Enabled = True
     Label2(7).Enabled = True
End Sub

Private Sub WinGame(Index As String)
     If Index = 0 Then
        If lblHint(0).Caption = "" And lblHint(1).Caption = "" And lblHint(2).Caption = "" And lblHint(3).Caption = "" And lblHint(4).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
        
     ElseIf Index = 1 Then
        
        If lblHint(5).Caption = "" And lblHint(6).Caption = "" And lblHint(7).Caption = "" And lblHint(8).Caption = "" And lblHint(9).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
        
     ElseIf Index = 2 Then
        If lblHint(10).Caption = "" And lblHint(11).Caption = "" And lblHint(12).Caption = "" And lblHint(13).Caption = "" And lblHint(14).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
        
     ElseIf Index = 3 Then
        
        If lblHint(15).Caption = "" And lblHint(16).Caption = "" And lblHint(17).Caption = "" And lblHint(18).Caption = "" And lblHint(19).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
    
     ElseIf Index = 4 Then
        If lblHint(20).Caption = "" And lblHint(21).Caption = "" And lblHint(22).Caption = "" And lblHint(23).Caption = "" And lblHint(24).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
        
    ElseIf Index = 5 Then
        
        If lblHint(25).Caption = "" And lblHint(26).Caption = "" And lblHint(27).Caption = "" And lblHint(28).Caption = "" And lblHint(29).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        End If
        
    ElseIf Index = 6 Then
        If lblHint(30).Caption = "" And lblHint(31).Caption = "" And lblHint(32).Caption = "" And lblHint(33).Caption = "" And lblHint(34).Caption = "" Then
            Call Disable
            cmdNextGame.Enabled = True
            lblStatus.Caption = "You Win!"
            lblLevel.Caption = lblLevel.Caption + 1
            Timer1.Enabled = False
        Else
            Call Disable
            lblStatus.Caption = "You Lose!"
            cmdNextGame.Enabled = True
            lblLevel.Caption = 1
            Timer1.Enabled = False
        End If
        
    End If
        
End Sub

Private Sub Timer1_Timer()

     lblTime2.Tag = lblTime2.Tag + 1
     
     If lblTime2.Tag = 60 Then
        lblTime1.Caption = lblTime1.Tag + 1 & ":"
        lblTime1.Tag = lblTime1.Caption
        lblTime2.Tag = 0
        lblTime2.Caption = "00"
     Else
         If lblTime2.Tag < 10 Then
            lblTime2.Caption = "0" & lblTime2.Tag
         ElseIf lblTime2.Tag > 10 Then
            lblTime2.Caption = lblTime2.Tag
         End If
     End If
     
End Sub
