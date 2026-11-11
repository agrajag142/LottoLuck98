VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Wheelsystem 
   Caption         =   "Wheeling System"
   BackColor       =   &H00000000&
   FillStyle       =   7
   FillColor       =   &H00404000&
   BorderStyle     =   1
   Icon            =   "Wheelsystem.frx":0000
   LinkTopic       =   "Form1"
   ClientLeft      =   1290
   ClientTop       =   1770
   ClientWidth     =   14355
   ClientHeight    =   8415
   PaletteMode     =   1
   ScaleHeight     =   8415
   ScaleWidth      =   14355
   Begin VB.CommandButton MENU5 
      Caption         =   "WHEELING SYSTEMS"
      Left            =   9600
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   108
      Appearance      =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Append"
      Left            =   360
      Top             =   3240
      Width           =   975
      Height          =   375
      TabIndex        =   107
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Left            =   360
      Top             =   3840
      Width           =   975
      Height          =   375
      TabIndex        =   106
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   9120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton MENU4 
      Caption         =   "GAME DESIGNER"
      Left            =   4800
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   105
      Appearance      =   0
      ToolTipText     =   "Create your own game files"
   End
   Begin VB.CommandButton MENU3 
      Caption         =   "GRAPHING"
      Left            =   7200
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   104
      Appearance      =   0
   End
   Begin VB.CommandButton MENU2 
      Caption         =   "CHECK WINNINGS"
      Left            =   2400
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   103
      Appearance      =   0
   End
   Begin VB.CommandButton MENU1 
      Caption         =   "ADD DRAW"
      Left            =   0
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   102
      Appearance      =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Left            =   360
      Top             =   4440
      Width           =   975
      Height          =   375
      TabIndex        =   101
   End
   Begin VB.ListBox WheeledList 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   120
      Top             =   5160
      Width           =   9375
      Height          =   1230
      TabIndex        =   100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QuickPick"
      BackColor       =   &H00C0C0C0&
      Left            =   2160
      Top             =   2400
      Width           =   975
      Height          =   255
      TabIndex        =   98
   End
   Begin VB.CommandButton button 
      Caption         =   "WHEEL NUMBERS"
      Left            =   4800
      Top             =   3120
      Width           =   2895
      Height          =   375
      Enabled         =   0
      TabIndex        =   97
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Left            =   120
      Top             =   6480
      Width           =   9375
      Height          =   255
      TabIndex        =   99
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VB.Shape Shape4 
      Left            =   240
      Top             =   3000
      Width           =   1215
      Height          =   1935
      Shape           =   4
      FillColor       =   &H00404040&
      FillStyle       =   0
   End
   Begin VB.Label Label97 
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   3840
      Width           =   3735
      Height          =   855
      TabIndex        =   96
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label96 
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   3120
      Width           =   375
      Height          =   255
      TabIndex        =   95
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label95 
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   3480
      Width           =   375
      Height          =   255
      TabIndex        =   94
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label94 
      Caption         =   "NUMBERS TO SELECT:"
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   1920
      Top             =   3120
      Width           =   1935
      Height          =   255
      TabIndex        =   93
      Alignment       =   1
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label93 
      Caption         =   "DESCRIPTION:"
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   1920
      Top             =   3840
      Width           =   1935
      Height          =   255
      TabIndex        =   92
      Alignment       =   1
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label92 
      Caption         =   "NUMBER OF GAMES:"
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Left            =   1920
      Top             =   3480
      Width           =   1935
      Height          =   255
      TabIndex        =   91
      Alignment       =   1
      BackStyle       =   0
      Appearance      =   0
   End
   Begin VB.Label Label91 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   7560
      Top             =   2400
      Width           =   1815
      Height          =   375
      TabIndex        =   90
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label90 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5880
      Top             =   2400
      Width           =   495
      Height          =   375
      TabIndex        =   89
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label89 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5280
      Top             =   2400
      Width           =   495
      Height          =   375
      TabIndex        =   88
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label88 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4680
      Top             =   2400
      Width           =   495
      Height          =   375
      TabIndex        =   87
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label87 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   2400
      Width           =   495
      Height          =   375
      TabIndex        =   86
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label86 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   3480
      Top             =   2400
      Width           =   495
      Height          =   375
      TabIndex        =   85
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label85 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8880
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   84
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label84 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8280
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   83
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label83 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7680
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   82
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label82 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7080
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   81
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label81 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6480
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   80
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label80 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5880
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   79
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label79 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5280
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   78
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label78 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4680
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   77
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label77 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   76
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label76 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   3480
      Top             =   1920
      Width           =   495
      Height          =   375
      TabIndex        =   75
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label75 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8880
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   74
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label74 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8280
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   73
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label73 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7680
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   72
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label72 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7080
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   71
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label71 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6480
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   70
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label70 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5880
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   69
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label69 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5280
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   68
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label68 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4680
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   67
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label67 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   66
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label66 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   3480
      Top             =   1440
      Width           =   495
      Height          =   375
      TabIndex        =   65
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label65 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8880
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   64
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label64 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8280
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   63
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label63 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7680
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   62
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label62 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7080
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   61
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label61 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6480
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   60
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label60 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5880
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   59
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label59 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5280
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   58
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label58 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4680
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   57
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label57 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   56
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label56 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   3480
      Top             =   960
      Width           =   495
      Height          =   375
      TabIndex        =   55
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label55 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8880
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   54
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label54 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8280
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   53
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label53 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7680
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   52
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label52 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7080
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   51
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label51 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6480
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   50
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label50 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5880
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   49
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label49 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   5280
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   48
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label48 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4680
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   47
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label47 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   4080
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   46
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label46 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   3480
      Top             =   480
      Width           =   495
      Height          =   375
      TabIndex        =   45
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label45 
      Caption         =   "45"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   2400
      Width           =   255
      Height          =   255
      TabIndex        =   44
      BorderStyle     =   1
   End
   Begin VB.Label Label44 
      Caption         =   "44"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   2400
      Width           =   255
      Height          =   255
      TabIndex        =   43
      BorderStyle     =   1
   End
   Begin VB.Label Label43 
      Caption         =   "43"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   2400
      Width           =   255
      Height          =   255
      TabIndex        =   42
      BorderStyle     =   1
   End
   Begin VB.Label Label42 
      Caption         =   "42"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   2400
      Width           =   255
      Height          =   255
      TabIndex        =   41
      BorderStyle     =   1
   End
   Begin VB.Label Label41 
      Caption         =   "41"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   2400
      Width           =   255
      Height          =   255
      TabIndex        =   40
      BorderStyle     =   1
   End
   Begin VB.Label Label40 
      Caption         =   "40"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2880
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   39
      BorderStyle     =   1
   End
   Begin VB.Label Label39 
      Caption         =   "39"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2520
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   38
      BorderStyle     =   1
   End
   Begin VB.Label Label38 
      Caption         =   "38"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2160
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   37
      BorderStyle     =   1
   End
   Begin VB.Label Label37 
      Caption         =   "37"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   36
      BorderStyle     =   1
   End
   Begin VB.Label Label36 
      Caption         =   "36"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   35
      BorderStyle     =   1
   End
   Begin VB.Label Label35 
      Caption         =   "35"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   34
      BorderStyle     =   1
   End
   Begin VB.Label Label34 
      Caption         =   "34"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   33
      BorderStyle     =   1
   End
   Begin VB.Label Label33 
      Caption         =   "33"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   2040
      Width           =   255
      Height          =   255
      TabIndex        =   32
      BorderStyle     =   1
   End
   Begin VB.Label Label32 
      Caption         =   "32"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2880
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   31
      BorderStyle     =   1
   End
   Begin VB.Label Label31 
      Caption         =   "31"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2520
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   30
      BorderStyle     =   1
   End
   Begin VB.Label Label30 
      Caption         =   "30"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2160
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   29
      BorderStyle     =   1
   End
   Begin VB.Label Label29 
      Caption         =   "29"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   28
      BorderStyle     =   1
   End
   Begin VB.Label Label28 
      Caption         =   "28"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   27
      BorderStyle     =   1
   End
   Begin VB.Label Label27 
      Caption         =   "27"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   26
      BorderStyle     =   1
   End
   Begin VB.Label Label26 
      Caption         =   "26"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   25
      BorderStyle     =   1
   End
   Begin VB.Label Label25 
      Caption         =   "25"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   1680
      Width           =   255
      Height          =   255
      TabIndex        =   24
      BorderStyle     =   1
   End
   Begin VB.Label Label24 
      Caption         =   "24"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2880
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   23
      BorderStyle     =   1
   End
   Begin VB.Label Label23 
      Caption         =   "23"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2520
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   22
      BorderStyle     =   1
   End
   Begin VB.Label Label22 
      Caption         =   "22"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2160
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   21
      BorderStyle     =   1
   End
   Begin VB.Label Label21 
      Caption         =   "21"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   20
      BorderStyle     =   1
   End
   Begin VB.Label Label20 
      Caption         =   "20"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   19
      BorderStyle     =   1
   End
   Begin VB.Label Label19 
      Caption         =   "19"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   18
      BorderStyle     =   1
   End
   Begin VB.Label Label18 
      Caption         =   "18"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   17
      BorderStyle     =   1
   End
   Begin VB.Label Label17 
      Caption         =   "17"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   1320
      Width           =   255
      Height          =   255
      TabIndex        =   16
      BorderStyle     =   1
   End
   Begin VB.Label Label16 
      Caption         =   "16"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2880
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   15
      BorderStyle     =   1
   End
   Begin VB.Label Label15 
      Caption         =   "15"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2520
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   14
      BorderStyle     =   1
   End
   Begin VB.Label Label14 
      Caption         =   "14"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2160
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   13
      BorderStyle     =   1
   End
   Begin VB.Label Label13 
      Caption         =   "13"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   12
      BorderStyle     =   1
   End
   Begin VB.Label Label12 
      Caption         =   "12"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   11
      BorderStyle     =   1
   End
   Begin VB.Label Label11 
      Caption         =   "11"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   10
      BorderStyle     =   1
   End
   Begin VB.Label Label10 
      Caption         =   "10"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   9
      BorderStyle     =   1
   End
   Begin VB.Label Label9 
      Caption         =   " 9"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   960
      Width           =   255
      Height          =   255
      TabIndex        =   8
      BorderStyle     =   1
   End
   Begin VB.Label Label8 
      Caption         =   " 8"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2880
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   7
      BorderStyle     =   1
   End
   Begin VB.Label Label7 
      Caption         =   " 7"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2520
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   6
      BorderStyle     =   1
   End
   Begin VB.Label Label6 
      Caption         =   " 6"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   2160
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   5
      BorderStyle     =   1
   End
   Begin VB.Label Label5 
      Caption         =   " 5"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1800
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   4
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      Caption         =   " 4"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1440
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   3
      BorderStyle     =   1
   End
   Begin VB.Label Label3 
      Caption         =   " 3"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   1080
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   2
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Caption         =   " 2"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   720
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   1
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Caption         =   " 1"
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Left            =   360
      Top             =   600
      Width           =   255
      Height          =   255
      TabIndex        =   0
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      Left            =   240
      Top             =   480
      Width           =   3015
      Height          =   2295
      Shape           =   4
      FillColor       =   &H00404000&
      FillStyle       =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BorderColor     =   &H00808080&
      Left            =   1680
      Top             =   3000
      Width           =   6375
      Height          =   1935
      Shape           =   4
      BorderStyle     =   0
      FillColor       =   &H00404040&
      FillStyle       =   0
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000040C0&
      BorderColor     =   &H00004080&
      Left            =   120
      Top             =   360
      Width           =   9375
      Height          =   2535
      Shape           =   4
      FillColor       =   &H00000080&
      FillStyle       =   0
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Appendtofile 
         Caption         =   "Append to"
      End
      Begin VB.Menu Autowheel 
         Caption         =   "AutoWheel"
         Checked         =   -1  'True
      End
      Begin VB.Menu PrintWheel 
         Caption         =   "Print"
      End
      Begin VB.Menu Save 
         Caption         =   "Save as"
      End
      Begin VB.Menu quittomain 
         Caption         =   "Quit to Main"
      End
   End
   Begin VB.Menu WheelType 
      Caption         =   "Wheel Type"
      Begin VB.Menu Wheelme 
         Caption         =   "Wheel Current Numbers"
      End
      Begin VB.Menu Space 
         Caption         =   "--------------------------------------"
      End
      Begin VB.Menu Fullwheel 
         Caption         =   "Full Wheels"
         Begin VB.Menu sixone 
            Caption         =   "  6 Numbers/    1 Game"
         End
         Begin VB.Menu sevenseven 
            Caption         =   "  7 Numbers/    7 Games"
         End
         Begin VB.Menu eighttwenty 
            Caption         =   "  8 Numbers/  28 Games"
         End
         Begin VB.Menu nineeightfour 
            Caption         =   "  9 Numbers/  84 Games"
         End
         Begin VB.Menu tentwooneo 
            Caption         =   "10 Numbers/210 Games"
         End
         Begin VB.Menu elevenfoursixtwo 
            Caption         =   "11 Numbers/462 Games"
         End
         Begin VB.Menu twelveninetwofour 
            Caption         =   "12 Numbers/924 Games"
         End
      End
      Begin VB.Menu ninetwelve 
         Caption         =   "  9 Numbers/12 Games"
      End
      Begin VB.Menu TennumbersTen 
         Caption         =   "10 Numbers/10 Games"
      End
      Begin VB.Menu elevenfive 
         Caption         =   "11 Numbers/  5 Games"
      End
      Begin VB.Menu twelvesix 
         Caption         =   "12 Numbers/  6 Games"
      End
      Begin VB.Menu twelveeighteen 
         Caption         =   "12 Numbers/18 Games"
      End
      Begin VB.Menu twelveforty 
         Caption         =   "12 Numbers/42 Games"
      End
      Begin VB.Menu forteenforteen 
         Caption         =   "14 Numbers/14 Games"
      End
      Begin VB.Menu seveneleven 
         Caption         =   "17 Numbers/11 Games"
      End
      Begin VB.Menu ofnftg 
         Caption         =   "18 Numbers/42 Games"
      End
      Begin VB.Menu twentyeight 
         Caption         =   "28 Numbers/36 Games"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu Coldest 
         Caption         =   "Coldest"
         Begin VB.Menu coldlast 
            Caption         =   "Last"
            Begin VB.Menu coldfivedraws 
               Caption         =   "   5 Draws"
            End
            Begin VB.Menu coldtendraws 
               Caption         =   " 10 Draws"
            End
            Begin VB.Menu coldfiftydraws 
               Caption         =   " 50 Draws"
            End
            Begin VB.Menu coldhunddraws 
               Caption         =   "100 Draws"
            End
            Begin VB.Menu coldfivhundraws 
               Caption         =   "500 Draws"
            End
         End
         Begin VB.Menu Coldfour 
            Caption         =   "Since 45 Introduced"
         End
         Begin VB.Menu ColdfourExcl 
            Caption         =   "Since 45 Introduced Excl. Supps."
         End
         Begin VB.Menu coldall 
            Caption         =   "Over All Draws"
         End
         Begin VB.Menu coldallexcl 
            Caption         =   "Over All Draws Excl. Supps."
         End
      End
      Begin VB.Menu hot 
         Caption         =   "Hottest"
         Begin VB.Menu hotlast 
            Caption         =   "Last"
            Begin VB.Menu hotfivedraws 
               Caption         =   "    5 Draws"
            End
            Begin VB.Menu hottendraws 
               Caption         =   "  10 Draws"
            End
            Begin VB.Menu Hotfiftydraws 
               Caption         =   "  50 Draws"
            End
            Begin VB.Menu hothundred 
               Caption         =   "100 Draws"
            End
            Begin VB.Menu hotfivedraws2 
               Caption         =   "500 Draws"
            End
         End
         Begin VB.Menu hottest 
            Caption         =   "Since 45 Introduced"
         End
         Begin VB.Menu Hotfournosup 
            Caption         =   "Since 45 Introduced Excl. Supps."
         End
         Begin VB.Menu Hotall 
            Caption         =   "Over All Draws"
         End
         Begin VB.Menu Hotallnosupp 
            Caption         =   "Over All Draws Excl. Supps."
         End
      End
      Begin VB.Menu powerdue 
         Caption         =   "Power/Due"
         Begin VB.Menu lastone 
            Caption         =   "Last 100"
         End
         Begin VB.Menu lasttwo 
            Caption         =   "Last 200"
         End
         Begin VB.Menu Lastfive 
            Caption         =   "Last 500"
         End
      End
      Begin VB.Menu random 
         Caption         =   "Quick Picks"
      End
      Begin VB.Menu lastdrawn 
         Caption         =   "Use Last Drawn Numbers"
      End
   End
   Begin VB.Menu WheelHelp 
      Caption         =   "Help"
      Begin VB.Menu HelpmeWheel 
         Caption         =   "NO HELP YET!!"
      End
   End
End
Attribute VB_Name = "Wheelsystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Wheelsystem - "WHEELING SYSTEMS": pick a set of numbers (click the 1-45 grid, or use the
' Hottest / Coldest / Power-Due / Quick Pick menus) and expand them through a lotto wheel.
' Grid labels Label1..Label45 = numbers 1..45 (white = free, &H80 = selected).
' Panel labels Label46..Label90 show the selected numbers, Label91 the count,
' Label95/96/97 the wheel's game count, field size and description.

Private Const COLOR_FREE = &HFFFFFF
Private Const COLOR_PICKED = &H80

Private Function GridLabel(ByVal n As Integer) As Label
    Set GridLabel = Me.Controls("Label" & n)
End Function

'---------------------------------------------------------------------------
'  Printing helpers
'---------------------------------------------------------------------------
Private Sub PrintCentered(y As Single, txt As String)
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(txt) / 2
    Printer.CurrentY = y
    Printer.Print txt
End Sub

Private Sub PrintCentered2(y As Single, txt As String)
    ' duplicate of PrintCentered in the original
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(txt) / 2
    Printer.CurrentY = y
    Printer.Print txt
End Sub

'---------------------------------------------------------------------------
'  Number grid
'---------------------------------------------------------------------------
Public Sub Chosenupdate()
    ' toggles the clicked number (ClickedNum) on the grid
    Dim lbl As Label
    If ClickedNum < 1 Or ClickedNum > 45 Then Exit Sub
    Set lbl = GridLabel(ClickedNum)
    If lbl.BackColor = COLOR_FREE Then
        If SelectedCount < WheelNumbers Then lbl.BackColor = COLOR_PICKED
    Else
        lbl.BackColor = COLOR_FREE
    End If
    updateselect
End Sub

Public Sub HotColdPicks()
    ' like Chosenupdate but only ever selects (used by the automatic pickers)
    Dim lbl As Label
    If ClickedNum < 1 Or ClickedNum > 45 Then Exit Sub
    Set lbl = GridLabel(ClickedNum)
    If lbl.BackColor = COLOR_FREE And SelectedCount < WheelNumbers Then
        lbl.BackColor = COLOR_PICKED
        updateselect
    End If
End Sub

Public Sub clearcolor()
    Dim n As Integer
    For n = 1 To 45
        GridLabel(n).BackColor = COLOR_FREE
    Next n
End Sub

Public Sub ClearPanel()
    Dim n As Integer
    For n = 1 To 45
        Chosen(n) = ""
    Next n
    SaveString = ""
End Sub

Public Sub onoffpanel()
    ' shows one panel slot (Label46..Label90) per number in the current wheel
    Dim i As Integer
    For i = 1 To 45
        Me.Controls("Label" & (45 + i)).Visible = (WheelNumbers >= i)
    Next i
End Sub

Public Sub updateselect()
    Dim n As Integer, k As Integer
    For n = 1 To 45
        Chosen(n) = ""
    Next n
    WheeledList.Clear
    k = 0
    For n = 1 To 45
        If GridLabel(n).BackColor = COLOR_PICKED Then
            k = k + 1
            Chosen(k) = CStr(n)
        End If
    Next n
    For n = 1 To 45
        Me.Controls("Label" & (45 + n)).Caption = Chosen(n)
    Next n
    Label91.Caption = "Selected: " & CStr(k)
    SelectedCount = k
    If k = WheelNumbers Then
        button.Enabled = True
        Wheelme.Enabled = True
    Else
        button.Enabled = False
        Wheelme.Enabled = False
    End If
End Sub

Private Sub ShowWheelInfo()
    Label97.Caption = WheelText
    Label95.Caption = CStr(WheelGames)
    Label96.Caption = CStr(WheelNumbers)
End Sub

'---------------------------------------------------------------------------
'  Expand the selection through the wheel
'---------------------------------------------------------------------------
Public Sub Wheelthedata()
    Dim g As Integer, p As Integer, k As Integer
    Dim row(5) As String
    For g = 1 To 1000
        Wheeled(g) = ""
    Next g
    SaveString = ""
    For g = 1 To WheelGames
        For p = 1 To 11 Step 2
            Wheeled(g) = Wheeled(g) & Chosen(CLng(Mid(WheelGame(g), p, 2)))
        Next p
    Next g
    WheeledList.Clear
    For g = 1 To WheelGames Step 5
        For k = 1 To 5
            row(k) = Format(Wheeled(g + k - 1), "## ## ## ## ## ## ## ##")
            If Len(row(k)) = 18 Then row(k) = "0" & Right$(row(k), 16)
            If Len(row(k)) = 19 Then row(k) = Right$(row(k), 17)
        Next k
        WheeledList.AddItem "     " & row(1) & "          " & row(2) & "          " & row(3) & _
                            "          " & row(4) & "          " & row(5)
    Next g
    GameCount = 1
    For g = 1 To WheelGames - 1
        SaveString = SaveString & Wheeled(g) & Chr$(10) & Chr$(13)
        Games(g) = Wheeled(g)
    Next g
    If WheelGames = 1 Then Games(1) = Wheeled(1)
    SaveString = SaveString & Wheeled(WheelGames)
    Games(WheelGames) = Wheeled(WheelGames)
    GameCount = WheelGames
End Sub

Private Sub button_Click()
    Wheelthedata
End Sub

Private Sub Wheelme_Click()
    Wheelthedata
End Sub

'---------------------------------------------------------------------------
'  Automatic pickers
'---------------------------------------------------------------------------
Private Sub PickRandom()
    Dim k As Integer
    If WheelNumbers = SelectedCount Then clearcolor
    Do While WheelNumbers > SelectedCount
        ClickedNum = Int(Rnd * 45) + 1
        HotColdPicks
        For k = 1 To CInt(Format(Time, "ss"))      ' stir the generator
            ClickedNum = CInt(Rnd)
        Next k
    Loop
End Sub

Private Sub Command1_Click()
    ' QuickPick button
    PickRandom
    If Autowheel.Checked = True Then Wheelthedata
End Sub

Private Sub random_Click()
    PickRandom
End Sub

' counts every ball over the last "draws" draws (or back to draw B0413 when since45 = True)
' using balls firstBall..lastBall of each draw, then selects the coldest (ascending=True)
' or hottest numbers until the wheel's field is full
Private Sub PickByFrequency(ByVal draws As Integer, ByVal lastBall As Integer, ByVal since45 As Boolean, ByVal ascending As Boolean)
    Dim counts(45) As Integer, order(46) As Integer
    Dim d As Integer, n As Integer, j As Integer, c As Integer, k As Integer
    ProgressBar1.Visible = True
    DrawIdx = LastDraw
    If since45 Then
        For d = 1 To LastDraw
            If Mid(DrawLine(d), 1, 5) = "B0413" Then draws = LastDraw - d
        Next d
    End If
    ProgressBar1.Min = 0
    ProgressBar1.Max = draws
    ProgressBar1.Value = 0
    For n = 1 To 45: counts(n) = 0: Next n
    For d = 1 To draws
        Update
        For n = 1 To 45
            For j = 1 To lastBall
                If n = Val(Drawn(j)) Then counts(n) = counts(n) + 1
            Next j
        Next n
        DrawIdx = DrawIdx - 1
        ProgressBar1.Value = d
    Next d
    k = 1
    If ascending Then
        For c = 0 To 1000
            For n = 1 To 45
                If counts(n) = c Then order(k) = n: k = k + 1
            Next n
        Next c
    Else
        For c = 1000 To 0 Step -1
            For n = 1 To 45
                If counts(n) = c Then order(k) = n: k = k + 1
            Next n
        Next c
    End If
    k = 1
    Do While WheelNumbers > SelectedCount And k <= 45
        ClickedNum = order(k)
        HotColdPicks
        k = k + 1
    Loop
    ProgressBar1.Visible = False
End Sub

' Power / Due: compares each ball's long-term rate with its rate over the last "recent" draws
Private Sub PickPowerDue(ByVal recent As Integer)
    Dim allTime(45) As Integer, lately(45) As Integer, power(45) As Integer, order(46) As Integer
    Dim d As Integer, n As Integer, j As Integer, c As Integer, k As Integer
    DrawIdx = LastDraw
    For d = 1 To LastDraw
        Update
        For n = 1 To 45
            For j = 1 To 8
                If n = Val(Drawn(j)) Then allTime(n) = allTime(n) + 1
            Next j
        Next n
        DrawIdx = DrawIdx - 1
    Next d
    DrawIdx = LastDraw
    For d = 1 To recent
        Update
        For n = 1 To 45
            For j = 1 To 8
                If n = Val(Drawn(j)) Then lately(n) = lately(n) + 1
            Next j
        Next n
        DrawIdx = DrawIdx - 1
    Next d
    For n = 1 To 45
        power(n) = CInt((allTime(n) / LastDraw) * 100 - (lately(n) / recent) * 100)
    Next n
    k = 1
    For c = 100 To -100 Step -1
        For n = 1 To 45
            If power(n) = c And k <= 45 Then order(k) = n: k = k + 1
        Next n
    Next c
    k = 1
    Do While WheelNumbers > SelectedCount And k <= 45
        ClickedNum = order(k)
        HotColdPicks
        k = k + 1
    Loop
End Sub

Private Sub coldfivedraws_Click()
    PickByFrequency 5, 8, False, True
End Sub
Private Sub coldtendraws_Click()
    PickByFrequency 10, 8, False, True
End Sub
Private Sub coldfiftydraws_Click()
    PickByFrequency 50, 8, False, True
End Sub
Private Sub coldhunddraws_Click()
    PickByFrequency 100, 8, False, True
End Sub
Private Sub coldfivhundraws_Click()
    PickByFrequency 500, 8, False, True
End Sub
Private Sub Coldfour_Click()
    PickByFrequency LastDraw, 8, True, True
End Sub
Private Sub ColdfourExcl_Click()
    PickByFrequency LastDraw, 6, True, True
End Sub
Private Sub coldall_Click()
    PickByFrequency LastDraw, 8, False, True
End Sub
Private Sub coldallexcl_Click()
    PickByFrequency LastDraw, 6, False, True
End Sub
Private Sub hotfivedraws_Click()
    PickByFrequency 5, 8, False, False
End Sub
Private Sub hottendraws_Click()
    PickByFrequency 10, 8, False, False
End Sub
Private Sub Hotfiftydraws_Click()
    PickByFrequency 50, 8, False, False
End Sub
Private Sub hothundred_Click()
    PickByFrequency 100, 8, False, False
End Sub
Private Sub hotfivedraws2_Click()
    PickByFrequency 500, 8, False, False
End Sub
Private Sub hottest_Click()
    PickByFrequency LastDraw, 8, True, False
End Sub
Private Sub Hotfournosup_Click()
    PickByFrequency LastDraw, 6, True, False
End Sub
Private Sub Hotall_Click()
    PickByFrequency LastDraw, 8, False, False
End Sub
Private Sub Hotallnosupp_Click()
    PickByFrequency LastDraw, 6, False, False
End Sub
Private Sub lastone_Click()
    PickPowerDue 100
End Sub
Private Sub lasttwo_Click()
    PickPowerDue 200
End Sub
Private Sub Lastfive_Click()
    PickPowerDue 500
End Sub
Private Sub lastdrawn_Click()
    ' select the numbers from the most recent draw
    Dim j As Integer
    DrawIdx = LastDraw
    Update
    j = 1
    Do While j < 9 And WheelNumbers > SelectedCount
        ClickedNum = CInt(Drawn(j))
        HotColdPicks
        j = j + 1
    Loop
End Sub

' 45 number labels on the grid
Private Sub NumberClick(ByVal n As Integer)
    ClickedNum = n
    Chosenupdate
End Sub
Private Sub Label1_Click()
    NumberClick 1
End Sub
Private Sub label2_Click()
    NumberClick 2
End Sub
Private Sub label3_Click()
    NumberClick 3
End Sub
Private Sub Label4_Click()
    NumberClick 4
End Sub
Private Sub Label5_Click()
    NumberClick 5
End Sub
Private Sub Label6_Click()
    NumberClick 6
End Sub
Private Sub Label7_Click()
    NumberClick 7
End Sub
Private Sub Label8_Click()
    NumberClick 8
End Sub
Private Sub Label9_Click()
    NumberClick 9
End Sub
Private Sub Label10_Click()
    NumberClick 10
End Sub
Private Sub Label11_Click()
    NumberClick 11
End Sub
Private Sub Label12_Click()
    NumberClick 12
End Sub
Private Sub Label13_Click()
    NumberClick 13
End Sub
Private Sub Label14_Click()
    NumberClick 14
End Sub
Private Sub Label15_Click()
    NumberClick 15
End Sub
Private Sub Label16_Click()
    NumberClick 16
End Sub
Private Sub Label17_Click()
    NumberClick 17
End Sub
Private Sub Label18_Click()
    NumberClick 18
End Sub
Private Sub Label19_Click()
    NumberClick 19
End Sub
Private Sub Label20_Click()
    NumberClick 20
End Sub
Private Sub Label21_Click()
    NumberClick 21
End Sub
Private Sub Label22_Click()
    NumberClick 22
End Sub
Private Sub Label23_Click()
    NumberClick 23
End Sub
Private Sub Label24_Click()
    NumberClick 24
End Sub
Private Sub Label25_Click()
    NumberClick 25
End Sub
Private Sub Label26_Click()
    NumberClick 26
End Sub
Private Sub Label27_Click()
    NumberClick 27
End Sub
Private Sub Label28_Click()
    NumberClick 28
End Sub
Private Sub Label29_Click()
    NumberClick 29
End Sub
Private Sub Label30_Click()
    NumberClick 30
End Sub
Private Sub Label31_Click()
    NumberClick 31
End Sub
Private Sub Label32_Click()
    NumberClick 32
End Sub
Private Sub Label33_Click()
    NumberClick 33
End Sub
Private Sub Label34_Click()
    NumberClick 34
End Sub
Private Sub Label35_Click()
    NumberClick 35
End Sub
Private Sub Label36_Click()
    NumberClick 36
End Sub
Private Sub Label37_Click()
    NumberClick 37
End Sub
Private Sub Label38_Click()
    NumberClick 38
End Sub
Private Sub Label39_Click()
    NumberClick 39
End Sub
Private Sub Label40_Click()
    NumberClick 40
End Sub
Private Sub Label41_Click()
    NumberClick 41
End Sub
Private Sub Label42_Click()
    NumberClick 42
End Sub
Private Sub Label43_Click()
    NumberClick 43
End Sub
Private Sub Label44_Click()
    NumberClick 44
End Sub
Private Sub Label45_Click()
    NumberClick 45
End Sub

'---------------------------------------------------------------------------
'  Form
'---------------------------------------------------------------------------
Public Sub Update()
    ' same draw-line parser as Lotto.Update
    Dim s As String, part As String
    Dim pos As Integer, ln As Integer, idx As Integer
    s = Mid(DrawLine(DrawIdx), 16)
    pos = 1: idx = 1: ln = 1
    Do While ln <= 3
        part = Mid(s, pos, ln)
        If Right$(part, 1) = "," Then
            Drawn(idx) = Mid(s, pos, ln - 1)
            pos = pos + ln
            idx = idx + 1
            ln = 0
        End If
        If idx = 8 Then ln = 3
        ln = ln + 1
    Loop
    Drawn(8) = Mid(s, pos, 2)
End Sub

Public Sub Closedown()
    Unload Board
    Unload Wheelsystem
    Unload Occurance
    Unload adddraw
    Unload Lotto
End Sub

Private Sub LayoutForm()
    Wheelsystem.Left = SysInfo1.WorkAreaLeft
    Wheelsystem.Width = SysInfo1.WorkAreaWidth
    Wheelsystem.Top = SysInfo1.WorkAreaTop
    Wheelsystem.Height = SysInfo1.WorkAreaHeight
    ProgressBar1.Top = SysInfo1.WorkAreaHeight - 1520 + 250
    WheeledList.Height = SysInfo1.WorkAreaHeight - WheeledList.Top - 1580 + 200
    MENU1.Width = Wheelsystem.Width / 5
    MENU2.Width = Wheelsystem.Width / 5
    MENU3.Width = Wheelsystem.Width / 5
    MENU4.Width = Wheelsystem.Width / 5
    MENU5.Width = Wheelsystem.Width / 5
    MENU2.Left = MENU1.Width
    MENU4.Left = MENU1.Width * 2
    MENU3.Left = MENU1.Width * 3
    MENU5.Left = MENU1.Width * 4
End Sub

Private Sub Form_Load()
    ClickedNum = 0
    ClickedNum2 = 0
    WheelGames = 0
    WheelNumbers = 0
    WheelText = ""
    Label91.Caption = "Selected: 0"
    clearcolor
    WheeledList.Clear
    LayoutForm
    ' default wheel: 14 numbers / 14 games
    forteenforteen_Click
    WheelText = "THESE 14 COMBINATIONS OF THE FIELD OF 14 NUMBERS' ASSURE A 3rd PRIZE WINNER     IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 14 !"
    ShowWheelInfo
    WheeledList.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Closedown
End Sub

Private Sub SysInfo1_DisplayChanged()
    If Screen.ActiveForm.Width <> SysInfo1.WorkAreaWidth Then
        Screen.ActiveForm.Left = SysInfo1.WorkAreaLeft
        Screen.ActiveForm.Width = SysInfo1.WorkAreaWidth
    End If
    If Screen.ActiveForm.Height <> SysInfo1.WorkAreaHeight Then
        Screen.ActiveForm.Top = SysInfo1.WorkAreaTop
        Screen.ActiveForm.Height = SysInfo1.WorkAreaHeight
    End If
    LayoutForm
End Sub

Private Sub Clear_Click()
    ClearPanel
    clearcolor
    updateselect
End Sub
Private Sub Command2_Click()
    ClearPanel
    clearcolor
    updateselect
End Sub

Private Sub Autowheel_Click()
    If Autowheel.Checked = True Then
        Autowheel.Checked = False
    ElseIf Autowheel.Checked = False Then
        Autowheel.Checked = True
    End If
End Sub

'---------------------------------------------------------------------------
'  Saving / appending the wheeled games as a .lot file
'---------------------------------------------------------------------------
Private Sub WriteGames(ByVal appendMode As Boolean)
    Dim f As Integer
    On Error GoTo WriteErr
    If SaveString = "" Then Exit Sub
    If appendMode Then
        CommonDialog1.DialogTitle = "Append A File"
    Else
        CommonDialog1.DialogTitle = "Save A File"
    End If
    CommonDialog1.Filter = "*.lot"
    CommonDialog1.FileName = "*.lot"
    CommonDialog1.Flags = 2
    CommonDialog1.ShowSave
    If SaveString = "" Then Exit Sub
    If Right$(CommonDialog1.FileName, 4) <> ".lot" Then
        CommonDialog1.FileName = CommonDialog1.FileName & ".lot"
    End If
    f = FreeFile
    If appendMode Then
        Open CommonDialog1.FileName For Append As #f
    Else
        Open CommonDialog1.FileName For Output As #f
    End If
    Print #f, SaveString
    Close #f
WriteErr:
End Sub

Private Sub Save_Click()
    WriteGames False
End Sub
Private Sub Command3_Click()
    WriteGames False
End Sub
Private Sub Appendtofile_Click()
    WriteGames True
End Sub
Private Sub Command4_Click()
    WriteGames True
End Sub

Private Sub PrintWheel_Click()
    Dim hdr As String
    Dim i As Integer, g As Integer, p As Integer, k As Integer
    If SaveString = "" Then Exit Sub
    PrintCentered 1440, "LOTTO LUCK 97"
    PrintCentered 1440, "(C) 1997 Justin Little"
    Printer.Print Chr$(13)
    hdr = ""
    For i = 46 To 90
        hdr = hdr & Me.Controls("Label" & i).Caption & " "
    Next i
    For i = 1 To Len(hdr) Step 70
        Printer.Print Mid(hdr, i, 70)
    Next i
    For i = 1 To Len(WheelText) Step 70
        Printer.Print Mid(WheelText, i, 70)
    Next i
    Printer.Print Chr$(13)
    For g = 1 To WheelGames Step 3
        For k = 0 To 2
            If g + k <= WheelGames Then
                For p = 1 To 10 Step 2
                    Printer.Print Mid$(Wheeled(g + k), p, 2) & ",";
                Next p
                Printer.Print Mid$(Wheeled(g + k), 11, 2);
                Printer.Print Spc(6);
            End If
        Next k
        Printer.Print
    Next g
    Printer.EndDoc
End Sub

'---------------------------------------------------------------------------
'  Navigation
'---------------------------------------------------------------------------
Private Sub MENU1_Click()
    adddraw.Show
    Wheelsystem.Hide
End Sub
Private Sub MENU2_Click()
    Lotto.Show
    Wheelsystem.Hide
End Sub
Private Sub MENU3_Click()
    Occurance.Show
    Wheelsystem.Hide
End Sub
Private Sub MENU4_Click()
    Board.Show
    Wheelsystem.Hide
End Sub
Private Sub Quittomain_Click()
    Lotto.Show
    Wheelsystem.Hide
End Sub

'---------------------------------------------------------------------------
'  Wheel definitions (combination tables recovered from the EXE)
'---------------------------------------------------------------------------
Private Sub sixone_Click()
    WheelText = "THIS 1 COMBINATION OF THE FIELD OF 6 NUMBERS ASSURE A 1ST PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 6 !"
    WheelGames = 1
    WheelNumbers = 6
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub sevenseven_Click()
    WheelText = "THESE 7 COMBINATIONS OF THE FIELD OF 7 NUMBERS ASSURE A 1ST PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 7 !"
    WheelGames = 7
    WheelNumbers = 7
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040507"
    WheelGame(3) = "010203040607"
    WheelGame(4) = "010203050607"
    WheelGame(5) = "010204050607"
    WheelGame(6) = "010304050607"
    WheelGame(7) = "020304050607"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub eighttwenty_Click()
    WheelText = "THESE 28 COMBINATIONS OF THE FIELD OF 8 NUMBERS ASSURE' A 1st PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 8 !"
    WheelGames = 28
    WheelNumbers = 8
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040507"
    WheelGame(3) = "010203040508"
    WheelGame(4) = "010203040607"
    WheelGame(5) = "010203040608"
    WheelGame(6) = "010203040708"
    WheelGame(7) = "010203050607"
    WheelGame(8) = "010203050608"
    WheelGame(9) = "010203050708"
    WheelGame(10) = "010203060708"
    WheelGame(11) = "010204050607"
    WheelGame(12) = "010204050608"
    WheelGame(13) = "010204050708"
    WheelGame(14) = "010204060708"
    WheelGame(15) = "010205060708"
    WheelGame(16) = "010304050607"
    WheelGame(17) = "010304050608"
    WheelGame(18) = "010304050708"
    WheelGame(19) = "010304060708"
    WheelGame(20) = "010305060708"
    WheelGame(21) = "010405060708"
    WheelGame(22) = "020304050607"
    WheelGame(23) = "020304050608"
    WheelGame(24) = "020304050708"
    WheelGame(25) = "020304060708"
    WheelGame(26) = "020305060708"
    WheelGame(27) = "020405060708"
    WheelGame(28) = "030405060708"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub nineeightfour_Click()
    WheelText = "THESE 84 COMBINATIONS OF THE FIELD OF 9 ASSURE A LOTTO WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 9 !"
    WheelGames = 84
    WheelNumbers = 9
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040507"
    WheelGame(3) = "010203040508"
    WheelGame(4) = "010203040509"
    WheelGame(5) = "010203040607"
    WheelGame(6) = "010203040608"
    WheelGame(7) = "010203040609"
    WheelGame(8) = "010203040708"
    WheelGame(9) = "010203040709"
    WheelGame(10) = "010203040809"
    WheelGame(11) = "010203050607"
    WheelGame(12) = "010203050608"
    WheelGame(13) = "010203050609"
    WheelGame(14) = "010203050708"
    WheelGame(15) = "010203050709"
    WheelGame(16) = "010203050809"
    WheelGame(17) = "010203060708"
    WheelGame(18) = "010203060709"
    WheelGame(19) = "010203060809"
    WheelGame(20) = "010203070809"
    WheelGame(21) = "010204050607"
    WheelGame(22) = "010204050608"
    WheelGame(23) = "010204050609"
    WheelGame(24) = "010204050708"
    WheelGame(25) = "010204050709"
    WheelGame(26) = "010204050809"
    WheelGame(27) = "010204060708"
    WheelGame(28) = "010204060709"
    WheelGame(29) = "010204060809"
    WheelGame(30) = "010204070809"
    WheelGame(31) = "010205060708"
    WheelGame(32) = "010205060709"
    WheelGame(33) = "010205060809"
    WheelGame(34) = "010205070809"
    WheelGame(35) = "010206070809"
    WheelGame(36) = "010304050607"
    WheelGame(37) = "010304050608"
    WheelGame(38) = "010304050609"
    WheelGame(39) = "010304050708"
    WheelGame(40) = "010304050709"
    WheelGame(41) = "010304050809"
    WheelGame(42) = "010304060708"
    WheelGame(43) = "010304060709"
    WheelGame(44) = "010304060809"
    WheelGame(45) = "010304070809"
    WheelGame(46) = "010305060708"
    WheelGame(47) = "010305060709"
    WheelGame(48) = "010305060809"
    WheelGame(49) = "010305070809"
    WheelGame(50) = "010306070809"
    WheelGame(51) = "010405060708"
    WheelGame(52) = "010405060709"
    WheelGame(53) = "010405060809"
    WheelGame(54) = "010405070809"
    WheelGame(55) = "010406070809"
    WheelGame(56) = "010506070809"
    WheelGame(57) = "020304050607"
    WheelGame(58) = "020304050608"
    WheelGame(59) = "020304050609"
    WheelGame(60) = "020304050708"
    WheelGame(61) = "020304050709"
    WheelGame(62) = "020304050809"
    WheelGame(63) = "020304060708"
    WheelGame(64) = "020304060709"
    WheelGame(65) = "020304060809"
    WheelGame(66) = "020304070809"
    WheelGame(67) = "020305060708"
    WheelGame(68) = "020305060709"
    WheelGame(69) = "020305060809"
    WheelGame(70) = "020305070809"
    WheelGame(71) = "020306070809"
    WheelGame(72) = "020405060708"
    WheelGame(73) = "020405060709"
    WheelGame(74) = "020405060809"
    WheelGame(75) = "020405070809"
    WheelGame(76) = "020406070809"
    WheelGame(77) = "020506070809"
    WheelGame(78) = "030405060708"
    WheelGame(79) = "030405060709"
    WheelGame(80) = "030405060809"
    WheelGame(81) = "030405070809"
    WheelGame(82) = "030406070809"
    WheelGame(83) = "030506070809"
    WheelGame(84) = "040506070809"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub tentwooneo_Click()
    WheelText = "THESE 210 COMBINATIONS OF THE FIELD OF 10 NUMBERS ASSURE A 1ST PRIZE WINNER." & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 10 !"
    WheelGames = 210
    WheelNumbers = 10
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040507"
    WheelGame(3) = "010203040508"
    WheelGame(4) = "010203040509"
    WheelGame(5) = "010203040510"
    WheelGame(6) = "010203040607"
    WheelGame(7) = "010203040608"
    WheelGame(8) = "010203040609"
    WheelGame(9) = "010203040610"
    WheelGame(10) = "010203040708"
    WheelGame(11) = "010203040709"
    WheelGame(12) = "010203040710"
    WheelGame(13) = "010203040809"
    WheelGame(14) = "010203040810"
    WheelGame(15) = "010203040910"
    WheelGame(16) = "010203050607"
    WheelGame(17) = "010203050608"
    WheelGame(18) = "010203050609"
    WheelGame(19) = "010203050610"
    WheelGame(20) = "010203050708"
    WheelGame(21) = "010203050709"
    WheelGame(22) = "010203050710"
    WheelGame(23) = "010203050809"
    WheelGame(24) = "010203050810"
    WheelGame(25) = "010203050910"
    WheelGame(26) = "010203060708"
    WheelGame(27) = "010203060709"
    WheelGame(28) = "010203060710"
    WheelGame(29) = "010203060809"
    WheelGame(30) = "010203060810"
    WheelGame(31) = "010203060910"
    WheelGame(32) = "010203070809"
    WheelGame(33) = "010203070810"
    WheelGame(34) = "010203070910"
    WheelGame(35) = "010203080910"
    WheelGame(36) = "010204050607"
    WheelGame(37) = "010204050608"
    WheelGame(38) = "010204050609"
    WheelGame(39) = "010204050610"
    WheelGame(40) = "010204050708"
    WheelGame(41) = "010204050709"
    WheelGame(42) = "010204050710"
    WheelGame(43) = "010204050809"
    WheelGame(44) = "010204050810"
    WheelGame(45) = "010204050910"
    WheelGame(46) = "010204060708"
    WheelGame(47) = "010204060709"
    WheelGame(48) = "010204060710"
    WheelGame(49) = "010204060809"
    WheelGame(50) = "010204060810"
    WheelGame(51) = "010204060910"
    WheelGame(52) = "010204070809"
    WheelGame(53) = "010204070810"
    WheelGame(54) = "010204070910"
    WheelGame(55) = "010204080910"
    WheelGame(56) = "010205060708"
    WheelGame(57) = "010205060709"
    WheelGame(58) = "010205060710"
    WheelGame(59) = "010205060809"
    WheelGame(60) = "010205060810"
    WheelGame(61) = "010205060910"
    WheelGame(62) = "010205070809"
    WheelGame(63) = "010205070810"
    WheelGame(64) = "010205070910"
    WheelGame(65) = "010205080910"
    WheelGame(66) = "010206070809"
    WheelGame(67) = "010206070810"
    WheelGame(68) = "010206070910"
    WheelGame(69) = "010206080910"
    WheelGame(70) = "010207080910"
    WheelGame(71) = "010304050607"
    WheelGame(72) = "010304050608"
    WheelGame(73) = "010304050609"
    WheelGame(74) = "010304050610"
    WheelGame(75) = "010304050708"
    WheelGame(76) = "010304050709"
    WheelGame(77) = "010304050710"
    WheelGame(78) = "010304050809"
    WheelGame(79) = "010304050810"
    WheelGame(80) = "010304050910"
    WheelGame(81) = "010304060708"
    WheelGame(82) = "010304060709"
    WheelGame(83) = "010304060710"
    WheelGame(84) = "010304060809"
    WheelGame(85) = "010304060810"
    WheelGame(86) = "010304060910"
    WheelGame(87) = "010304070809"
    WheelGame(88) = "010304070810"
    WheelGame(89) = "010304070910"
    WheelGame(90) = "010304080910"
    WheelGame(91) = "010305060708"
    WheelGame(92) = "010305060709"
    WheelGame(93) = "010305060710"
    WheelGame(94) = "010305060809"
    WheelGame(95) = "010305060810"
    WheelGame(96) = "010305060910"
    WheelGame(97) = "010305070809"
    WheelGame(98) = "010305070810"
    WheelGame(99) = "010305070910"
    WheelGame(100) = "010305080910"
    WheelGame(101) = "010306070809"
    WheelGame(102) = "010306070810"
    WheelGame(103) = "010306070910"
    WheelGame(104) = "010306080910"
    WheelGame(105) = "010307080910"
    WheelGame(106) = "010405060708"
    WheelGame(107) = "010405060709"
    WheelGame(108) = "010405060710"
    WheelGame(109) = "010405060809"
    WheelGame(110) = "010405060810"
    WheelGame(111) = "010405060910"
    WheelGame(112) = "010405070809"
    WheelGame(113) = "010405070810"
    WheelGame(114) = "010405070910"
    WheelGame(115) = "010405080910"
    WheelGame(116) = "010406070809"
    WheelGame(117) = "010406070810"
    WheelGame(118) = "010406070910"
    WheelGame(119) = "010406080910"
    WheelGame(120) = "010407080910"
    WheelGame(121) = "010506070809"
    WheelGame(122) = "010506070810"
    WheelGame(123) = "010506070910"
    WheelGame(124) = "010506080910"
    WheelGame(125) = "010507080910"
    WheelGame(126) = "010607080910"
    WheelGame(127) = "020304050607"
    WheelGame(128) = "020304050608"
    WheelGame(129) = "020304050609"
    WheelGame(130) = "020304050610"
    WheelGame(131) = "020304050708"
    WheelGame(132) = "020304050709"
    WheelGame(133) = "020304050710"
    WheelGame(134) = "020304050809"
    WheelGame(135) = "020304050810"
    WheelGame(136) = "020304050910"
    WheelGame(137) = "020304060708"
    WheelGame(138) = "020304060709"
    WheelGame(139) = "020304060710"
    WheelGame(140) = "020304060809"
    WheelGame(141) = "020304060810"
    WheelGame(142) = "020304060910"
    WheelGame(143) = "020304070809"
    WheelGame(144) = "020304070810"
    WheelGame(145) = "020304070910"
    WheelGame(146) = "020304080910"
    WheelGame(147) = "020305060708"
    WheelGame(148) = "020305060709"
    WheelGame(149) = "020305060710"
    WheelGame(150) = "020305060809"
    WheelGame(151) = "020305060810"
    WheelGame(152) = "020305060910"
    WheelGame(153) = "020305070809"
    WheelGame(154) = "020305070810"
    WheelGame(155) = "020305070910"
    WheelGame(156) = "020305080910"
    WheelGame(157) = "020306070809"
    WheelGame(158) = "020306070810"
    WheelGame(159) = "020306070910"
    WheelGame(160) = "020306080910"
    WheelGame(161) = "020307080910"
    WheelGame(162) = "020405060708"
    WheelGame(163) = "020405060709"
    WheelGame(164) = "020405060710"
    WheelGame(165) = "020405060809"
    WheelGame(166) = "020405060810"
    WheelGame(167) = "020405060910"
    WheelGame(168) = "020405070809"
    WheelGame(169) = "020405070810"
    WheelGame(170) = "020405070910"
    WheelGame(171) = "020405080910"
    WheelGame(172) = "020406070809"
    WheelGame(173) = "020406070810"
    WheelGame(174) = "020406070910"
    WheelGame(175) = "020406080910"
    WheelGame(176) = "020407080910"
    WheelGame(177) = "020506070809"
    WheelGame(178) = "020506070810"
    WheelGame(179) = "020506070910"
    WheelGame(180) = "020506080910"
    WheelGame(181) = "020507080910"
    WheelGame(182) = "020607080910"
    WheelGame(183) = "030405060708"
    WheelGame(184) = "030405060709"
    WheelGame(185) = "030405060710"
    WheelGame(186) = "030405060809"
    WheelGame(187) = "030405060810"
    WheelGame(188) = "030405060910"
    WheelGame(189) = "030405070809"
    WheelGame(190) = "030405070810"
    WheelGame(191) = "030405070910"
    WheelGame(192) = "030405080910"
    WheelGame(193) = "030406070809"
    WheelGame(194) = "030406070810"
    WheelGame(195) = "030406070910"
    WheelGame(196) = "030406080910"
    WheelGame(197) = "030407080910"
    WheelGame(198) = "030506070809"
    WheelGame(199) = "030506070810"
    WheelGame(200) = "030506070910"
    WheelGame(201) = "030506080910"
    WheelGame(202) = "030507080910"
    WheelGame(203) = "030607080910"
    WheelGame(204) = "040506070809"
    WheelGame(205) = "040506070810"
    WheelGame(206) = "040506070910"
    WheelGame(207) = "040506080910"
    WheelGame(208) = "040507080910"
    WheelGame(209) = "040607080910"
    WheelGame(210) = "050607080910"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub elevenfoursixtwo_Click()
    WheelText = "THESE 462 COMBINATIONS OF THE FIELD OF 11 NUMBERS ASSURE A 1ST PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 11 !"
    WheelGames = 462
    WheelNumbers = 11
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040507"
    WheelGame(3) = "010203040508"
    WheelGame(4) = "010203040509"
    WheelGame(5) = "010203040510"
    WheelGame(6) = "010203040511"
    WheelGame(7) = "010203040607"
    WheelGame(8) = "010203040608"
    WheelGame(9) = "010203040609"
    WheelGame(10) = "010203040610"
    WheelGame(11) = "010203040611"
    WheelGame(12) = "010203040708"
    WheelGame(13) = "010203040709"
    WheelGame(14) = "010203040710"
    WheelGame(15) = "010203040711"
    WheelGame(16) = "010203040809"
    WheelGame(17) = "010203040810"
    WheelGame(18) = "010203040811"
    WheelGame(19) = "010203040910"
    WheelGame(20) = "010203040911"
    WheelGame(21) = "010203041011"
    WheelGame(22) = "010203050607"
    WheelGame(23) = "010203050608"
    WheelGame(24) = "010203050609"
    WheelGame(25) = "010203050610"
    WheelGame(26) = "010203050611"
    WheelGame(27) = "010203050708"
    WheelGame(28) = "010203050709"
    WheelGame(29) = "010203050710"
    WheelGame(30) = "010203050711"
    WheelGame(31) = "010203050809"
    WheelGame(32) = "010203050810"
    WheelGame(33) = "010203050811"
    WheelGame(34) = "010203050910"
    WheelGame(35) = "010203050911"
    WheelGame(36) = "010203051011"
    WheelGame(37) = "010203060708"
    WheelGame(38) = "010203060709"
    WheelGame(39) = "010203060710"
    WheelGame(40) = "010203060711"
    WheelGame(41) = "010203060809"
    WheelGame(42) = "010203060810"
    WheelGame(43) = "010203060811"
    WheelGame(44) = "010203060910"
    WheelGame(45) = "010203060911"
    WheelGame(46) = "010203061011"
    WheelGame(47) = "010203070809"
    WheelGame(48) = "010203070810"
    WheelGame(49) = "010203070811"
    WheelGame(50) = "010203070910"
    WheelGame(51) = "010203070911"
    WheelGame(52) = "010203071011"
    WheelGame(53) = "010203080910"
    WheelGame(54) = "010203080911"
    WheelGame(55) = "010203081011"
    WheelGame(56) = "010203091011"
    WheelGame(57) = "010204050607"
    WheelGame(58) = "010204050608"
    WheelGame(59) = "010204050609"
    WheelGame(60) = "010204050610"
    WheelGame(61) = "010204050611"
    WheelGame(62) = "010204050708"
    WheelGame(63) = "010204050709"
    WheelGame(64) = "010204050710"
    WheelGame(65) = "010204050711"
    WheelGame(66) = "010204050809"
    WheelGame(67) = "010204050810"
    WheelGame(68) = "010204050811"
    WheelGame(69) = "010204050910"
    WheelGame(70) = "010204050911"
    WheelGame(71) = "010204051011"
    WheelGame(72) = "010204060708"
    WheelGame(73) = "010204060709"
    WheelGame(74) = "010204060710"
    WheelGame(75) = "010204060711"
    WheelGame(76) = "010204060809"
    WheelGame(77) = "010204060810"
    WheelGame(78) = "010204060811"
    WheelGame(79) = "010204060910"
    WheelGame(80) = "010204060911"
    WheelGame(81) = "010204061011"
    WheelGame(82) = "010204070809"
    WheelGame(83) = "010204070810"
    WheelGame(84) = "010204070811"
    WheelGame(85) = "010204070910"
    WheelGame(86) = "010204070911"
    WheelGame(87) = "010204071011"
    WheelGame(88) = "010204080910"
    WheelGame(89) = "010204080911"
    WheelGame(90) = "010204081011"
    WheelGame(91) = "010204091011"
    WheelGame(92) = "010205060708"
    WheelGame(93) = "010205060709"
    WheelGame(94) = "010205060710"
    WheelGame(95) = "010205060711"
    WheelGame(96) = "010205060809"
    WheelGame(97) = "010205060810"
    WheelGame(98) = "010205060811"
    WheelGame(99) = "010205060910"
    WheelGame(100) = "010205060911"
    WheelGame(101) = "010205061011"
    WheelGame(102) = "010205070809"
    WheelGame(103) = "010205070810"
    WheelGame(104) = "010205070811"
    WheelGame(105) = "010205070910"
    WheelGame(106) = "010205070911"
    WheelGame(107) = "010205071011"
    WheelGame(108) = "010205080910"
    WheelGame(109) = "010205080911"
    WheelGame(110) = "010205081011"
    WheelGame(111) = "010205091011"
    WheelGame(112) = "010206070809"
    WheelGame(113) = "010206070810"
    WheelGame(114) = "010206070811"
    WheelGame(115) = "010206070910"
    WheelGame(116) = "010206070911"
    WheelGame(117) = "010206071011"
    WheelGame(118) = "010206080910"
    WheelGame(119) = "010206080911"
    WheelGame(120) = "010206081011"
    WheelGame(121) = "010206091011"
    WheelGame(122) = "010207080910"
    WheelGame(123) = "010207080911"
    WheelGame(124) = "010207081011"
    WheelGame(125) = "010207091011"
    WheelGame(126) = "010208091011"
    WheelGame(127) = "010304050607"
    WheelGame(128) = "010304050608"
    WheelGame(129) = "010304050609"
    WheelGame(130) = "010304050610"
    WheelGame(131) = "010304050611"
    WheelGame(132) = "010304050708"
    WheelGame(133) = "010304050709"
    WheelGame(134) = "010304050710"
    WheelGame(135) = "010304050711"
    WheelGame(136) = "010304050809"
    WheelGame(137) = "010304050810"
    WheelGame(138) = "010304050811"
    WheelGame(139) = "010304050910"
    WheelGame(140) = "010304050911"
    WheelGame(141) = "010304051011"
    WheelGame(142) = "010304060708"
    WheelGame(143) = "010304060709"
    WheelGame(144) = "010304060710"
    WheelGame(145) = "010304060711"
    WheelGame(146) = "010304060809"
    WheelGame(147) = "010304060810"
    WheelGame(148) = "010304060811"
    WheelGame(149) = "010304060910"
    WheelGame(150) = "010304060911"
    WheelGame(151) = "010304061011"
    WheelGame(152) = "010304070809"
    WheelGame(153) = "010304070810"
    WheelGame(154) = "010304070811"
    WheelGame(155) = "010304070910"
    WheelGame(156) = "010304070911"
    WheelGame(157) = "010304071011"
    WheelGame(158) = "010304080910"
    WheelGame(159) = "010304080911"
    WheelGame(160) = "010304081011"
    WheelGame(161) = "010304091011"
    WheelGame(162) = "010305060708"
    WheelGame(163) = "010305060709"
    WheelGame(164) = "010305060710"
    WheelGame(165) = "010305060711"
    WheelGame(166) = "010305060809"
    WheelGame(167) = "010305060810"
    WheelGame(168) = "010305060811"
    WheelGame(169) = "010305060910"
    WheelGame(170) = "010305060911"
    WheelGame(171) = "010305061011"
    WheelGame(172) = "010305070809"
    WheelGame(173) = "010305070810"
    WheelGame(174) = "010305070811"
    WheelGame(175) = "010305070910"
    WheelGame(176) = "010305070911"
    WheelGame(177) = "010305071011"
    WheelGame(178) = "010305080910"
    WheelGame(179) = "010305080911"
    WheelGame(180) = "010305081011"
    WheelGame(181) = "010305091011"
    WheelGame(182) = "010306070809"
    WheelGame(183) = "010306070810"
    WheelGame(184) = "010306070811"
    WheelGame(185) = "010306070910"
    WheelGame(186) = "010306070911"
    WheelGame(187) = "010306071011"
    WheelGame(188) = "010306080910"
    WheelGame(189) = "010306080911"
    WheelGame(190) = "010306081011"
    WheelGame(191) = "010306091011"
    WheelGame(192) = "010307080910"
    WheelGame(193) = "010307080911"
    WheelGame(194) = "010307081011"
    WheelGame(195) = "010307091011"
    WheelGame(196) = "010308091011"
    WheelGame(197) = "010405060708"
    WheelGame(198) = "010405060709"
    WheelGame(199) = "010405060710"
    WheelGame(200) = "010405060711"
    WheelGame(201) = "010405060809"
    WheelGame(202) = "010405060810"
    WheelGame(203) = "010405060811"
    WheelGame(204) = "010405060910"
    WheelGame(205) = "010405060911"
    WheelGame(206) = "010405061011"
    WheelGame(207) = "010405070809"
    WheelGame(208) = "010405070810"
    WheelGame(209) = "010405070811"
    WheelGame(210) = "010405070910"
    WheelGame(211) = "010405070911"
    WheelGame(212) = "010405071011"
    WheelGame(213) = "010405080910"
    WheelGame(214) = "010405080911"
    WheelGame(215) = "010405081011"
    WheelGame(216) = "010405091011"
    WheelGame(217) = "010406070809"
    WheelGame(218) = "010406070810"
    WheelGame(219) = "010406070811"
    WheelGame(220) = "010406070910"
    WheelGame(221) = "010406070911"
    WheelGame(222) = "010406071011"
    WheelGame(223) = "010406080910"
    WheelGame(224) = "010406080911"
    WheelGame(225) = "010406081011"
    WheelGame(226) = "010406091011"
    WheelGame(227) = "010407080910"
    WheelGame(228) = "010407080911"
    WheelGame(229) = "010407081011"
    WheelGame(230) = "010407091011"
    WheelGame(231) = "010408091011"
    WheelGame(232) = "010506070809"
    WheelGame(233) = "010506070810"
    WheelGame(234) = "010506070811"
    WheelGame(235) = "010506070910"
    WheelGame(236) = "010506070911"
    WheelGame(237) = "010506071011"
    WheelGame(238) = "010506080910"
    WheelGame(239) = "010506080911"
    WheelGame(240) = "010506081011"
    WheelGame(241) = "010506091011"
    WheelGame(242) = "010507080910"
    WheelGame(243) = "010507080911"
    WheelGame(244) = "010507081011"
    WheelGame(245) = "010507091011"
    WheelGame(246) = "010508091011"
    WheelGame(247) = "010607080910"
    WheelGame(248) = "010607080911"
    WheelGame(249) = "010607081011"
    WheelGame(250) = "010607091011"
    WheelGame(251) = "010608091011"
    WheelGame(252) = "010708091011"
    WheelGame(253) = "020304050607"
    WheelGame(254) = "020304050608"
    WheelGame(255) = "020304050609"
    WheelGame(256) = "020304050610"
    WheelGame(257) = "020304050611"
    WheelGame(258) = "020304050708"
    WheelGame(259) = "020304050709"
    WheelGame(260) = "020304050710"
    WheelGame(261) = "020304050711"
    WheelGame(262) = "020304050809"
    WheelGame(263) = "020304050810"
    WheelGame(264) = "020304050811"
    WheelGame(265) = "020304050910"
    WheelGame(266) = "020304050911"
    WheelGame(267) = "020304051011"
    WheelGame(268) = "020304060708"
    WheelGame(269) = "020304060709"
    WheelGame(270) = "020304060710"
    WheelGame(271) = "020304060711"
    WheelGame(272) = "020304060809"
    WheelGame(273) = "020304060810"
    WheelGame(274) = "020304060811"
    WheelGame(275) = "020304060910"
    WheelGame(276) = "020304060911"
    WheelGame(277) = "020304061011"
    WheelGame(278) = "020304070809"
    WheelGame(279) = "020304070810"
    WheelGame(280) = "020304070811"
    WheelGame(281) = "020304070910"
    WheelGame(282) = "020304070911"
    WheelGame(283) = "020304071011"
    WheelGame(284) = "020304080910"
    WheelGame(285) = "020304080911"
    WheelGame(286) = "020304081011"
    WheelGame(287) = "020304091011"
    WheelGame(288) = "020305060708"
    WheelGame(289) = "020305060709"
    WheelGame(290) = "020305060710"
    WheelGame(291) = "020305060711"
    WheelGame(292) = "020305060809"
    WheelGame(293) = "020305060810"
    WheelGame(294) = "020305060811"
    WheelGame(295) = "020305060910"
    WheelGame(296) = "020305060911"
    WheelGame(297) = "020305061011"
    WheelGame(298) = "020305070809"
    WheelGame(299) = "020305070810"
    WheelGame(300) = "020305070811"
    WheelGame(301) = "020305070910"
    WheelGame(302) = "020305070911"
    WheelGame(303) = "020305071011"
    WheelGame(304) = "020305080910"
    WheelGame(305) = "020305080911"
    WheelGame(306) = "020305081011"
    WheelGame(307) = "020305091011"
    WheelGame(308) = "020306070809"
    WheelGame(309) = "020306070810"
    WheelGame(310) = "020306070811"
    WheelGame(311) = "020306070910"
    WheelGame(312) = "020306070911"
    WheelGame(313) = "020306071011"
    WheelGame(314) = "020306080910"
    WheelGame(315) = "020306080911"
    WheelGame(316) = "020306081011"
    WheelGame(317) = "020306091011"
    WheelGame(318) = "020307080910"
    WheelGame(319) = "020307080911"
    WheelGame(320) = "020307081011"
    WheelGame(321) = "020307091011"
    WheelGame(322) = "020308091011"
    WheelGame(323) = "020405060708"
    WheelGame(324) = "020405060709"
    WheelGame(325) = "020405060710"
    WheelGame(326) = "020405060711"
    WheelGame(327) = "020405060809"
    WheelGame(328) = "020405060810"
    WheelGame(329) = "020405060811"
    WheelGame(330) = "020405060910"
    WheelGame(331) = "020405060911"
    WheelGame(332) = "020405061011"
    WheelGame(333) = "020405070809"
    WheelGame(334) = "020405070810"
    WheelGame(335) = "020405070811"
    WheelGame(336) = "020405070910"
    WheelGame(337) = "020405070911"
    WheelGame(338) = "020405071011"
    WheelGame(339) = "020405080910"
    WheelGame(340) = "020405080911"
    WheelGame(341) = "020405081011"
    WheelGame(342) = "020405091011"
    WheelGame(343) = "020406070809"
    WheelGame(344) = "020406070810"
    WheelGame(345) = "020406070811"
    WheelGame(346) = "020406070910"
    WheelGame(347) = "020406070911"
    WheelGame(348) = "020406071011"
    WheelGame(349) = "020406080910"
    WheelGame(350) = "020406080911"
    WheelGame(351) = "020406081011"
    WheelGame(352) = "020406091011"
    WheelGame(353) = "020407080910"
    WheelGame(354) = "020407080911"
    WheelGame(355) = "020407081011"
    WheelGame(356) = "020407091011"
    WheelGame(357) = "020408091011"
    WheelGame(358) = "020506070809"
    WheelGame(359) = "020506070810"
    WheelGame(360) = "020506070811"
    WheelGame(361) = "020506070910"
    WheelGame(362) = "020506070911"
    WheelGame(363) = "020506071011"
    WheelGame(364) = "020506080910"
    WheelGame(365) = "020506080911"
    WheelGame(366) = "020506081011"
    WheelGame(367) = "020506091011"
    WheelGame(368) = "020507080910"
    WheelGame(369) = "020507080911"
    WheelGame(370) = "020507081011"
    WheelGame(371) = "020507091011"
    WheelGame(372) = "020508091011"
    WheelGame(373) = "020607080910"
    WheelGame(374) = "020607080911"
    WheelGame(375) = "020607081011"
    WheelGame(376) = "020607091011"
    WheelGame(377) = "020608091011"
    WheelGame(378) = "020708091011"
    WheelGame(379) = "030405060708"
    WheelGame(380) = "030405060709"
    WheelGame(381) = "030405060710"
    WheelGame(382) = "030405060711"
    WheelGame(383) = "030405060809"
    WheelGame(384) = "030405060810"
    WheelGame(385) = "030405060811"
    WheelGame(386) = "030405060910"
    WheelGame(387) = "030405060911"
    WheelGame(388) = "030405061011"
    WheelGame(389) = "030405070809"
    WheelGame(390) = "030405070810"
    WheelGame(391) = "030405070811"
    WheelGame(392) = "030405070910"
    WheelGame(393) = "030405070911"
    WheelGame(394) = "030405071011"
    WheelGame(395) = "030405080910"
    WheelGame(396) = "030405080911"
    WheelGame(397) = "030405081011"
    WheelGame(398) = "030405091011"
    WheelGame(399) = "030406070809"
    WheelGame(400) = "030406070810"
    WheelGame(401) = "030406070811"
    WheelGame(402) = "030406070910"
    WheelGame(403) = "030406070911"
    WheelGame(404) = "030406071011"
    WheelGame(405) = "030406080910"
    WheelGame(406) = "030406080911"
    WheelGame(407) = "030406081011"
    WheelGame(408) = "030406091011"
    WheelGame(409) = "030407080910"
    WheelGame(410) = "030407080911"
    WheelGame(411) = "030407081011"
    WheelGame(412) = "030407091011"
    WheelGame(413) = "030408091011"
    WheelGame(414) = "030506070809"
    WheelGame(415) = "030506070810"
    WheelGame(416) = "030506070811"
    WheelGame(417) = "030506070910"
    WheelGame(418) = "030506070911"
    WheelGame(419) = "030506071011"
    WheelGame(420) = "030506080910"
    WheelGame(421) = "030506080911"
    WheelGame(422) = "030506081011"
    WheelGame(423) = "030506091011"
    WheelGame(424) = "030507080910"
    WheelGame(425) = "030507080911"
    WheelGame(426) = "030507081011"
    WheelGame(427) = "030507091011"
    WheelGame(428) = "030508091011"
    WheelGame(429) = "030607080910"
    WheelGame(430) = "030607080911"
    WheelGame(431) = "030607081011"
    WheelGame(432) = "030607091011"
    WheelGame(433) = "030608091011"
    WheelGame(434) = "030708091011"
    WheelGame(435) = "040506070809"
    WheelGame(436) = "040506070810"
    WheelGame(437) = "040506070811"
    WheelGame(438) = "040506070910"
    WheelGame(439) = "040506070911"
    WheelGame(440) = "040506071011"
    WheelGame(441) = "040506080910"
    WheelGame(442) = "040506080911"
    WheelGame(443) = "040506081011"
    WheelGame(444) = "040506091011"
    WheelGame(445) = "040507080910"
    WheelGame(446) = "040507080911"
    WheelGame(447) = "040507081011"
    WheelGame(448) = "040507091011"
    WheelGame(449) = "040508091011"
    WheelGame(450) = "040607080910"
    WheelGame(451) = "040607080911"
    WheelGame(452) = "040607081011"
    WheelGame(453) = "040607091011"
    WheelGame(454) = "040608091011"
    WheelGame(455) = "040708091011"
    WheelGame(456) = "050607080910"
    WheelGame(457) = "050607080911"
    WheelGame(458) = "050607081011"
    WheelGame(459) = "050607091011"
    WheelGame(460) = "050608091011"
    WheelGame(461) = "050708091011"
    WheelGame(462) = "060708091011"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub twelveninetwofour_Click()
    WheelText = "THESE 924 COMBINATIONS OF THE FIELD OF 12 NUMBERS ASSURE A 1ST PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 12 !"
    WheelGames = 924
    WheelNumbers = 12
    ShowWheelInfo
    WheelGame(1) = "010205060910"
    WheelGame(2) = "010203040506"
    WheelGame(3) = "010203040507"
    WheelGame(4) = "010203040508"
    WheelGame(5) = "010203040509"
    WheelGame(6) = "010203040510"
    WheelGame(7) = "010203040511"
    WheelGame(8) = "010203040512"
    WheelGame(9) = "010203040607"
    WheelGame(10) = "010203040608"
    WheelGame(11) = "010203040609"
    WheelGame(12) = "010203040610"
    WheelGame(13) = "010203040611"
    WheelGame(14) = "010203040612"
    WheelGame(15) = "010203040708"
    WheelGame(16) = "010203040709"
    WheelGame(17) = "010203040710"
    WheelGame(18) = "010203040711"
    WheelGame(19) = "010203040712"
    WheelGame(20) = "010203040809"
    WheelGame(21) = "010203040810"
    WheelGame(22) = "010203040811"
    WheelGame(23) = "010203040812"
    WheelGame(24) = "010203040910"
    WheelGame(25) = "010203040911"
    WheelGame(26) = "010203040912"
    WheelGame(27) = "010203041011"
    WheelGame(28) = "010203041012"
    WheelGame(29) = "010203041112"
    WheelGame(30) = "010203050607"
    WheelGame(31) = "010203050608"
    WheelGame(32) = "010203050609"
    WheelGame(33) = "010203050610"
    WheelGame(34) = "010203050611"
    WheelGame(35) = "010203050612"
    WheelGame(36) = "010203050708"
    WheelGame(37) = "010203050709"
    WheelGame(38) = "010203050710"
    WheelGame(39) = "010203050711"
    WheelGame(40) = "010203050712"
    WheelGame(41) = "010203050809"
    WheelGame(42) = "010203050810"
    WheelGame(43) = "010203050811"
    WheelGame(44) = "010203050812"
    WheelGame(45) = "010203050910"
    WheelGame(46) = "010203050911"
    WheelGame(47) = "010203050912"
    WheelGame(48) = "010203051011"
    WheelGame(49) = "010203051012"
    WheelGame(50) = "010203051112"
    WheelGame(51) = "010203060708"
    WheelGame(52) = "010203060709"
    WheelGame(53) = "010203060710"
    WheelGame(54) = "010203060711"
    WheelGame(55) = "010203060712"
    WheelGame(56) = "010203060809"
    WheelGame(57) = "010203060810"
    WheelGame(58) = "010203060811"
    WheelGame(59) = "010203060812"
    WheelGame(60) = "010203060910"
    WheelGame(61) = "010203060911"
    WheelGame(62) = "010203060912"
    WheelGame(63) = "010203061011"
    WheelGame(64) = "010203061012"
    WheelGame(65) = "010203061112"
    WheelGame(66) = "010203070809"
    WheelGame(67) = "010203070810"
    WheelGame(68) = "010203070811"
    WheelGame(69) = "010203070812"
    WheelGame(70) = "010203070910"
    WheelGame(71) = "010203070911"
    WheelGame(72) = "010203070912"
    WheelGame(73) = "010203071011"
    WheelGame(74) = "010203071012"
    WheelGame(75) = "010203071112"
    WheelGame(76) = "010203080910"
    WheelGame(77) = "010203080911"
    WheelGame(78) = "010203080912"
    WheelGame(79) = "010203081011"
    WheelGame(80) = "010203081012"
    WheelGame(81) = "010203081112"
    WheelGame(82) = "010203091011"
    WheelGame(83) = "010203091012"
    WheelGame(84) = "010203091112"
    WheelGame(85) = "010203101112"
    WheelGame(86) = "010204050607"
    WheelGame(87) = "010204050608"
    WheelGame(88) = "010204050609"
    WheelGame(89) = "010204050610"
    WheelGame(90) = "010204050611"
    WheelGame(91) = "010204050612"
    WheelGame(92) = "010204050708"
    WheelGame(93) = "010204050709"
    WheelGame(94) = "010204050710"
    WheelGame(95) = "010204050711"
    WheelGame(96) = "010204050712"
    WheelGame(97) = "010204050809"
    WheelGame(98) = "010204050810"
    WheelGame(99) = "010204050811"
    WheelGame(100) = "010204050812"
    WheelGame(101) = "010204050910"
    WheelGame(102) = "010204050911"
    WheelGame(103) = "010204050912"
    WheelGame(104) = "010204051011"
    WheelGame(105) = "010204051012"
    WheelGame(106) = "010204051112"
    WheelGame(107) = "010204060708"
    WheelGame(108) = "010204060709"
    WheelGame(109) = "010204060710"
    WheelGame(110) = "010204060711"
    WheelGame(111) = "010204060712"
    WheelGame(112) = "010204060809"
    WheelGame(113) = "010204060810"
    WheelGame(114) = "010204060811"
    WheelGame(115) = "010204060812"
    WheelGame(116) = "010204060910"
    WheelGame(117) = "010204060911"
    WheelGame(118) = "010204060912"
    WheelGame(119) = "010204061011"
    WheelGame(120) = "010204061012"
    WheelGame(121) = "010204061112"
    WheelGame(122) = "010204070809"
    WheelGame(123) = "010204070810"
    WheelGame(124) = "010204070811"
    WheelGame(125) = "010204070812"
    WheelGame(126) = "010204070910"
    WheelGame(127) = "010204070911"
    WheelGame(128) = "010204070912"
    WheelGame(129) = "010204071011"
    WheelGame(130) = "010204071012"
    WheelGame(131) = "010204071112"
    WheelGame(132) = "010204080910"
    WheelGame(133) = "010204080911"
    WheelGame(134) = "010204080912"
    WheelGame(135) = "010204081011"
    WheelGame(136) = "010204081012"
    WheelGame(137) = "010204081112"
    WheelGame(138) = "010204091011"
    WheelGame(139) = "010204091012"
    WheelGame(140) = "010204091112"
    WheelGame(141) = "010204101112"
    WheelGame(142) = "010205060708"
    WheelGame(143) = "010205060709"
    WheelGame(144) = "010205060710"
    WheelGame(145) = "010205060711"
    WheelGame(146) = "010205060712"
    WheelGame(147) = "010205060809"
    WheelGame(148) = "010205060810"
    WheelGame(149) = "010205060811"
    WheelGame(150) = "010205060812"
    WheelGame(151) = "010205060911"
    WheelGame(152) = "010205060912"
    WheelGame(153) = "010205061011"
    WheelGame(154) = "010205061012"
    WheelGame(155) = "010205061112"
    WheelGame(156) = "010205070809"
    WheelGame(157) = "010205070810"
    WheelGame(158) = "010205070811"
    WheelGame(159) = "010205070812"
    WheelGame(160) = "010205070910"
    WheelGame(161) = "010205070911"
    WheelGame(162) = "010205070912"
    WheelGame(163) = "010205071011"
    WheelGame(164) = "010205071012"
    WheelGame(165) = "010205071112"
    WheelGame(166) = "010205080910"
    WheelGame(167) = "010205080911"
    WheelGame(168) = "010205080912"
    WheelGame(169) = "010205081011"
    WheelGame(170) = "010205081012"
    WheelGame(171) = "010205081112"
    WheelGame(172) = "010205091011"
    WheelGame(173) = "010205091012"
    WheelGame(174) = "010205091112"
    WheelGame(175) = "010205101112"
    WheelGame(176) = "010206070809"
    WheelGame(177) = "010206070810"
    WheelGame(178) = "010206070811"
    WheelGame(179) = "010206070812"
    WheelGame(180) = "010206070910"
    WheelGame(181) = "010206070911"
    WheelGame(182) = "010206070912"
    WheelGame(183) = "010206071011"
    WheelGame(184) = "010206071012"
    WheelGame(185) = "010206071112"
    WheelGame(186) = "010206080910"
    WheelGame(187) = "010206080911"
    WheelGame(188) = "010206080912"
    WheelGame(189) = "010206081011"
    WheelGame(190) = "010206081012"
    WheelGame(191) = "010206081112"
    WheelGame(192) = "010206091011"
    WheelGame(193) = "010206091012"
    WheelGame(194) = "010206091112"
    WheelGame(195) = "010206101112"
    WheelGame(196) = "010207080910"
    WheelGame(197) = "010207080911"
    WheelGame(198) = "010207080912"
    WheelGame(199) = "010207081011"
    WheelGame(200) = "010207081012"
    WheelGame(201) = "010207081112"
    WheelGame(202) = "010207091011"
    WheelGame(203) = "010207091012"
    WheelGame(204) = "010207091112"
    WheelGame(205) = "010207101112"
    WheelGame(206) = "010208091011"
    WheelGame(207) = "010208091012"
    WheelGame(208) = "010208091112"
    WheelGame(209) = "010208101112"
    WheelGame(210) = "010209101112"
    WheelGame(211) = "010304050607"
    WheelGame(212) = "010304050608"
    WheelGame(213) = "010304050609"
    WheelGame(214) = "010304050610"
    WheelGame(215) = "010304050611"
    WheelGame(216) = "010304050612"
    WheelGame(217) = "010304050708"
    WheelGame(218) = "010304050709"
    WheelGame(219) = "010304050710"
    WheelGame(220) = "010304050711"
    WheelGame(221) = "010304050712"
    WheelGame(222) = "010304050809"
    WheelGame(223) = "010304050810"
    WheelGame(224) = "010304050811"
    WheelGame(225) = "010304050812"
    WheelGame(226) = "010304050910"
    WheelGame(227) = "010304050911"
    WheelGame(228) = "010304050912"
    WheelGame(229) = "010304051011"
    WheelGame(230) = "010304051012"
    WheelGame(231) = "010304051112"
    WheelGame(232) = "010304060708"
    WheelGame(233) = "010304060709"
    WheelGame(234) = "010304060710"
    WheelGame(235) = "010304060711"
    WheelGame(236) = "010304060712"
    WheelGame(237) = "010304060809"
    WheelGame(238) = "010304060810"
    WheelGame(239) = "010304060811"
    WheelGame(240) = "010304060812"
    WheelGame(241) = "010304060910"
    WheelGame(242) = "010304060911"
    WheelGame(243) = "010304060912"
    WheelGame(244) = "010304061011"
    WheelGame(245) = "010304061012"
    WheelGame(246) = "010304061112"
    WheelGame(247) = "010304070809"
    WheelGame(248) = "010304070810"
    WheelGame(249) = "010304070811"
    WheelGame(250) = "010304070812"
    WheelGame(251) = "010304070910"
    WheelGame(252) = "010304070911"
    WheelGame(253) = "010304070912"
    WheelGame(254) = "010304071011"
    WheelGame(255) = "010304071012"
    WheelGame(256) = "010304071112"
    WheelGame(257) = "010304080910"
    WheelGame(258) = "010304080911"
    WheelGame(259) = "010304080912"
    WheelGame(260) = "010304081011"
    WheelGame(261) = "010304081012"
    WheelGame(262) = "010304081112"
    WheelGame(263) = "010304091011"
    WheelGame(264) = "010304091012"
    WheelGame(265) = "010304091112"
    WheelGame(266) = "010304101112"
    WheelGame(267) = "010305060708"
    WheelGame(268) = "010305060709"
    WheelGame(269) = "010305060710"
    WheelGame(270) = "010305060711"
    WheelGame(271) = "010305060712"
    WheelGame(272) = "010305060809"
    WheelGame(273) = "010305060810"
    WheelGame(274) = "010305060811"
    WheelGame(275) = "010305060812"
    WheelGame(276) = "010305060910"
    WheelGame(277) = "010305060911"
    WheelGame(278) = "010305060912"
    WheelGame(279) = "010305061011"
    WheelGame(280) = "010305061012"
    WheelGame(281) = "010305061112"
    WheelGame(282) = "010305070809"
    WheelGame(283) = "010305070810"
    WheelGame(284) = "010305070811"
    WheelGame(285) = "010305070812"
    WheelGame(286) = "010305070910"
    WheelGame(287) = "010305070911"
    WheelGame(288) = "010305070912"
    WheelGame(289) = "010305071011"
    WheelGame(290) = "010305071012"
    WheelGame(291) = "010305071112"
    WheelGame(292) = "010305080910"
    WheelGame(293) = "010305080911"
    WheelGame(294) = "010305080912"
    WheelGame(295) = "010305081011"
    WheelGame(296) = "010305081012"
    WheelGame(297) = "010305081112"
    WheelGame(298) = "010305091011"
    WheelGame(299) = "010305091012"
    WheelGame(300) = "010305091112"
    WheelGame(301) = "010305101112"
    WheelGame(302) = "010306070809"
    WheelGame(303) = "010306070810"
    WheelGame(304) = "010306070811"
    WheelGame(305) = "010306070812"
    WheelGame(306) = "010306070910"
    WheelGame(307) = "010306070911"
    WheelGame(308) = "010306070912"
    WheelGame(309) = "010306071011"
    WheelGame(310) = "010306071012"
    WheelGame(311) = "010306071112"
    WheelGame(312) = "010306080910"
    WheelGame(313) = "010306080911"
    WheelGame(314) = "010306080912"
    WheelGame(315) = "010306081011"
    WheelGame(316) = "010306081012"
    WheelGame(317) = "010306081112"
    WheelGame(318) = "010306091011"
    WheelGame(319) = "010306091012"
    WheelGame(320) = "010306091112"
    WheelGame(321) = "010306101112"
    WheelGame(322) = "010307080910"
    WheelGame(323) = "010307080911"
    WheelGame(324) = "010307080912"
    WheelGame(325) = "010307081011"
    WheelGame(326) = "010307081012"
    WheelGame(327) = "010307081112"
    WheelGame(328) = "010307091011"
    WheelGame(329) = "010307091012"
    WheelGame(330) = "010307091112"
    WheelGame(331) = "010307101112"
    WheelGame(332) = "010308091011"
    WheelGame(333) = "010308091012"
    WheelGame(334) = "010308091112"
    WheelGame(335) = "010308101112"
    WheelGame(336) = "010309101112"
    WheelGame(337) = "010405060708"
    WheelGame(338) = "010405060709"
    WheelGame(339) = "010405060710"
    WheelGame(340) = "010405060711"
    WheelGame(341) = "010405060712"
    WheelGame(342) = "010405060809"
    WheelGame(343) = "010405060810"
    WheelGame(344) = "010405060811"
    WheelGame(345) = "010405060812"
    WheelGame(346) = "010405060910"
    WheelGame(347) = "010405060911"
    WheelGame(348) = "010405060912"
    WheelGame(349) = "010405061011"
    WheelGame(350) = "010405061012"
    WheelGame(351) = "010405061112"
    WheelGame(352) = "010405070809"
    WheelGame(353) = "010405070810"
    WheelGame(354) = "010405070811"
    WheelGame(355) = "010405070812"
    WheelGame(356) = "010405070910"
    WheelGame(357) = "010405070911"
    WheelGame(358) = "010405070912"
    WheelGame(359) = "010405071011"
    WheelGame(360) = "010405071012"
    WheelGame(361) = "010405071112"
    WheelGame(362) = "010405080910"
    WheelGame(363) = "010405080911"
    WheelGame(364) = "010405080912"
    WheelGame(365) = "010405081011"
    WheelGame(366) = "010405081012"
    WheelGame(367) = "010405081112"
    WheelGame(368) = "010405091011"
    WheelGame(369) = "010405091012"
    WheelGame(370) = "010405091112"
    WheelGame(371) = "010405101112"
    WheelGame(372) = "010406070809"
    WheelGame(373) = "010406070810"
    WheelGame(374) = "010406070811"
    WheelGame(375) = "010406070812"
    WheelGame(376) = "010406070910"
    WheelGame(377) = "010406070911"
    WheelGame(378) = "010406070912"
    WheelGame(379) = "010406071011"
    WheelGame(380) = "010406071012"
    WheelGame(381) = "010406071112"
    WheelGame(382) = "010406080910"
    WheelGame(383) = "010406080911"
    WheelGame(384) = "010406080912"
    WheelGame(385) = "010406081011"
    WheelGame(386) = "010406081012"
    WheelGame(387) = "010406081112"
    WheelGame(388) = "010406091011"
    WheelGame(389) = "010406091012"
    WheelGame(390) = "010406091112"
    WheelGame(391) = "010406101112"
    WheelGame(392) = "010407080910"
    WheelGame(393) = "010407080911"
    WheelGame(394) = "010407080912"
    WheelGame(395) = "010407081011"
    WheelGame(396) = "010407081012"
    WheelGame(397) = "010407081112"
    WheelGame(398) = "010407091011"
    WheelGame(399) = "010407091012"
    WheelGame(400) = "010407091112"
    WheelGame(401) = "010407101112"
    WheelGame(402) = "010408091011"
    WheelGame(403) = "010408091012"
    WheelGame(404) = "010408091112"
    WheelGame(405) = "010408101112"
    WheelGame(406) = "010409101112"
    WheelGame(407) = "010506070809"
    WheelGame(408) = "010506070810"
    WheelGame(409) = "010506070811"
    WheelGame(410) = "010506070812"
    WheelGame(411) = "010506070910"
    WheelGame(412) = "010506070911"
    WheelGame(413) = "010506070912"
    WheelGame(414) = "010506071011"
    WheelGame(415) = "010506071012"
    WheelGame(416) = "010506071112"
    WheelGame(417) = "010506080910"
    WheelGame(418) = "010506080911"
    WheelGame(419) = "010506080912"
    WheelGame(420) = "010506081011"
    WheelGame(421) = "010506081012"
    WheelGame(422) = "010506081112"
    WheelGame(423) = "010506091011"
    WheelGame(424) = "010506091012"
    WheelGame(425) = "010506091112"
    WheelGame(426) = "010506101112"
    WheelGame(427) = "010507080910"
    WheelGame(428) = "010507080911"
    WheelGame(429) = "010507080912"
    WheelGame(430) = "010507081011"
    WheelGame(431) = "010507081012"
    WheelGame(432) = "010507081112"
    WheelGame(433) = "010507091011"
    WheelGame(434) = "010507091012"
    WheelGame(435) = "010507091112"
    WheelGame(436) = "010507101112"
    WheelGame(437) = "010508091011"
    WheelGame(438) = "010508091012"
    WheelGame(439) = "010508091112"
    WheelGame(440) = "010508101112"
    WheelGame(441) = "010509101112"
    WheelGame(442) = "010607080910"
    WheelGame(443) = "010607080911"
    WheelGame(444) = "010607080912"
    WheelGame(445) = "010607081011"
    WheelGame(446) = "010607081012"
    WheelGame(447) = "010607081112"
    WheelGame(448) = "010607091011"
    WheelGame(449) = "010607091012"
    WheelGame(450) = "010607091112"
    WheelGame(451) = "010607101112"
    WheelGame(452) = "010608091011"
    WheelGame(453) = "010608091012"
    WheelGame(454) = "010608091112"
    WheelGame(455) = "010608101112"
    WheelGame(456) = "010609101112"
    WheelGame(457) = "010708091011"
    WheelGame(458) = "010708091012"
    WheelGame(459) = "010708091112"
    WheelGame(460) = "010708101112"
    WheelGame(461) = "010709101112"
    WheelGame(462) = "010809101112"
    WheelGame(463) = "020304050607"
    WheelGame(464) = "020304050608"
    WheelGame(465) = "020304050609"
    WheelGame(466) = "020304050610"
    WheelGame(467) = "020304050611"
    WheelGame(468) = "020304050612"
    WheelGame(469) = "020304050708"
    WheelGame(470) = "020304050709"
    WheelGame(471) = "020304050710"
    WheelGame(472) = "020304050711"
    WheelGame(473) = "020304050712"
    WheelGame(474) = "020304050809"
    WheelGame(475) = "020304050810"
    WheelGame(476) = "020304050811"
    WheelGame(477) = "020304050812"
    WheelGame(478) = "020304050910"
    WheelGame(479) = "020304050911"
    WheelGame(480) = "020304050912"
    WheelGame(481) = "020304051011"
    WheelGame(482) = "020304051012"
    WheelGame(483) = "020304051112"
    WheelGame(484) = "020304060708"
    WheelGame(485) = "020304060709"
    WheelGame(486) = "020304060710"
    WheelGame(487) = "020304060711"
    WheelGame(488) = "020304060712"
    WheelGame(489) = "020304060809"
    WheelGame(490) = "020304060810"
    WheelGame(491) = "020304060811"
    WheelGame(492) = "020304060812"
    WheelGame(493) = "020304060910"
    WheelGame(494) = "020304060911"
    WheelGame(495) = "020304060912"
    WheelGame(496) = "020304061011"
    WheelGame(497) = "020304061012"
    WheelGame(498) = "020304061112"
    WheelGame(499) = "020304070809"
    WheelGame(500) = "020304070810"
    WheelGame(501) = "020304070811"
    WheelGame(502) = "020304070812"
    WheelGame(503) = "020304070910"
    WheelGame(504) = "020304070911"
    WheelGame(505) = "020304070912"
    WheelGame(506) = "020304071011"
    WheelGame(507) = "020304071012"
    WheelGame(508) = "020304071112"
    WheelGame(509) = "020304080910"
    WheelGame(510) = "020304080911"
    WheelGame(511) = "020304080912"
    WheelGame(512) = "020304081011"
    WheelGame(513) = "020304081012"
    WheelGame(514) = "020304081112"
    WheelGame(515) = "020304091011"
    WheelGame(516) = "020304091012"
    WheelGame(517) = "020304091112"
    WheelGame(518) = "020304101112"
    WheelGame(519) = "020305060708"
    WheelGame(520) = "020305060709"
    WheelGame(521) = "020305060710"
    WheelGame(522) = "020305060711"
    WheelGame(523) = "020305060712"
    WheelGame(524) = "020305060809"
    WheelGame(525) = "020305060810"
    WheelGame(526) = "020305060811"
    WheelGame(527) = "020305060812"
    WheelGame(528) = "020305060910"
    WheelGame(529) = "020305060911"
    WheelGame(530) = "020305060912"
    WheelGame(531) = "020305061011"
    WheelGame(532) = "020305061012"
    WheelGame(533) = "020305061112"
    WheelGame(534) = "020305070809"
    WheelGame(535) = "020305070810"
    WheelGame(536) = "020305070811"
    WheelGame(537) = "020305070812"
    WheelGame(538) = "020305070910"
    WheelGame(539) = "020305070911"
    WheelGame(540) = "020305070912"
    WheelGame(541) = "020305071011"
    WheelGame(542) = "020305071012"
    WheelGame(543) = "020305071112"
    WheelGame(544) = "020305080910"
    WheelGame(545) = "020305080911"
    WheelGame(546) = "020305080912"
    WheelGame(547) = "020305081011"
    WheelGame(548) = "020305081012"
    WheelGame(549) = "020305081112"
    WheelGame(550) = "020305091011"
    WheelGame(551) = "020305091012"
    WheelGame(552) = "020305091112"
    WheelGame(553) = "020305101112"
    WheelGame(554) = "020306070809"
    WheelGame(555) = "020306070810"
    WheelGame(556) = "020306070811"
    WheelGame(557) = "020306070812"
    WheelGame(558) = "020306070910"
    WheelGame(559) = "020306070911"
    WheelGame(560) = "020306070912"
    WheelGame(561) = "020306071011"
    WheelGame(562) = "020306071012"
    WheelGame(563) = "020306071112"
    WheelGame(564) = "020306080910"
    WheelGame(565) = "020306080911"
    WheelGame(566) = "020306080912"
    WheelGame(567) = "020306081011"
    WheelGame(568) = "020306081012"
    WheelGame(569) = "020306081112"
    WheelGame(570) = "020306091011"
    WheelGame(571) = "020306091012"
    WheelGame(572) = "020306091112"
    WheelGame(573) = "020306101112"
    WheelGame(574) = "020307080910"
    WheelGame(575) = "020307080911"
    WheelGame(576) = "020307080912"
    WheelGame(577) = "020307081011"
    WheelGame(578) = "020307081012"
    WheelGame(579) = "020307081112"
    WheelGame(580) = "020307091011"
    WheelGame(581) = "020307091012"
    WheelGame(582) = "020307091112"
    WheelGame(583) = "020307101112"
    WheelGame(584) = "020308091011"
    WheelGame(585) = "020308091012"
    WheelGame(586) = "020308091112"
    WheelGame(587) = "020308101112"
    WheelGame(588) = "020309101112"
    WheelGame(589) = "020405060708"
    WheelGame(590) = "020405060709"
    WheelGame(591) = "020405060710"
    WheelGame(592) = "020405060711"
    WheelGame(593) = "020405060712"
    WheelGame(594) = "020405060809"
    WheelGame(595) = "020405060810"
    WheelGame(596) = "020405060811"
    WheelGame(597) = "020405060812"
    WheelGame(598) = "020405060910"
    WheelGame(599) = "020405060911"
    WheelGame(600) = "020405060912"
    WheelGame(601) = "020405061011"
    WheelGame(602) = "020405061012"
    WheelGame(603) = "020405061112"
    WheelGame(604) = "020405070809"
    WheelGame(605) = "020405070810"
    WheelGame(606) = "020405070811"
    WheelGame(607) = "020405070812"
    WheelGame(608) = "020405070910"
    WheelGame(609) = "020405070911"
    WheelGame(610) = "020405070912"
    WheelGame(611) = "020405071011"
    WheelGame(612) = "020405071012"
    WheelGame(613) = "020405071112"
    WheelGame(614) = "020405080910"
    WheelGame(615) = "020405080911"
    WheelGame(616) = "020405080912"
    WheelGame(617) = "020405081011"
    WheelGame(618) = "020405081012"
    WheelGame(619) = "020405081112"
    WheelGame(620) = "020405091011"
    WheelGame(621) = "020405091012"
    WheelGame(622) = "020405091112"
    WheelGame(623) = "020405101112"
    WheelGame(624) = "020406070809"
    WheelGame(625) = "020406070810"
    WheelGame(626) = "020406070811"
    WheelGame(627) = "020406070812"
    WheelGame(628) = "020406070910"
    WheelGame(629) = "020406070911"
    WheelGame(630) = "020406070912"
    WheelGame(631) = "020406071011"
    WheelGame(632) = "020406071012"
    WheelGame(633) = "020406071112"
    WheelGame(634) = "020406080910"
    WheelGame(635) = "020406080911"
    WheelGame(636) = "020406080912"
    WheelGame(637) = "020406081011"
    WheelGame(638) = "020406081012"
    WheelGame(639) = "020406081112"
    WheelGame(640) = "020406091011"
    WheelGame(641) = "020406091012"
    WheelGame(642) = "020406091112"
    WheelGame(643) = "020406101112"
    WheelGame(644) = "020407080910"
    WheelGame(645) = "020407080911"
    WheelGame(646) = "020407080912"
    WheelGame(647) = "020407081011"
    WheelGame(648) = "020407081012"
    WheelGame(649) = "020407081112"
    WheelGame(650) = "020407091011"
    WheelGame(651) = "020407091012"
    WheelGame(652) = "020407091112"
    WheelGame(653) = "020407101112"
    WheelGame(654) = "020408091011"
    WheelGame(655) = "020408091012"
    WheelGame(656) = "020408091112"
    WheelGame(657) = "020408101112"
    WheelGame(658) = "020409101112"
    WheelGame(659) = "020506070809"
    WheelGame(660) = "020506070810"
    WheelGame(661) = "020506070811"
    WheelGame(662) = "020506070812"
    WheelGame(663) = "020506070910"
    WheelGame(664) = "020506070911"
    WheelGame(665) = "020506070912"
    WheelGame(666) = "020506071011"
    WheelGame(667) = "020506071012"
    WheelGame(668) = "020506071112"
    WheelGame(669) = "020506080910"
    WheelGame(670) = "020506080911"
    WheelGame(671) = "020506080912"
    WheelGame(672) = "020506081011"
    WheelGame(673) = "020506081012"
    WheelGame(674) = "020506081112"
    WheelGame(675) = "020506091011"
    WheelGame(676) = "020506091012"
    WheelGame(677) = "020506091112"
    WheelGame(678) = "020506101112"
    WheelGame(679) = "020507080910"
    WheelGame(680) = "020507080911"
    WheelGame(681) = "020507080912"
    WheelGame(682) = "020507081011"
    WheelGame(683) = "020507081012"
    WheelGame(684) = "020507081112"
    WheelGame(685) = "020507091011"
    WheelGame(686) = "020507091012"
    WheelGame(687) = "020507091112"
    WheelGame(688) = "020507101112"
    WheelGame(689) = "020508091011"
    WheelGame(690) = "020508091012"
    WheelGame(691) = "020508091112"
    WheelGame(692) = "020508101112"
    WheelGame(693) = "020509101112"
    WheelGame(694) = "020607080910"
    WheelGame(695) = "020607080911"
    WheelGame(696) = "020607080912"
    WheelGame(697) = "020607081011"
    WheelGame(698) = "020607081012"
    WheelGame(699) = "020607081112"
    WheelGame(700) = "020607091011"
    WheelGame(701) = "020607091012"
    WheelGame(702) = "020607091112"
    WheelGame(703) = "020607101112"
    WheelGame(704) = "020608091011"
    WheelGame(705) = "020608091012"
    WheelGame(706) = "020608091112"
    WheelGame(707) = "020608101112"
    WheelGame(708) = "020609101112"
    WheelGame(709) = "020708091011"
    WheelGame(710) = "020708091012"
    WheelGame(711) = "020708091112"
    WheelGame(712) = "020708101112"
    WheelGame(713) = "020709101112"
    WheelGame(714) = "020809101112"
    WheelGame(715) = "030405060708"
    WheelGame(716) = "030405060709"
    WheelGame(717) = "030405060710"
    WheelGame(718) = "030405060711"
    WheelGame(719) = "030405060712"
    WheelGame(720) = "030405060809"
    WheelGame(721) = "030405060810"
    WheelGame(722) = "030405060811"
    WheelGame(723) = "030405060812"
    WheelGame(724) = "030405060910"
    WheelGame(725) = "030405060911"
    WheelGame(726) = "030405060912"
    WheelGame(727) = "030405061011"
    WheelGame(728) = "030405061012"
    WheelGame(729) = "030405061112"
    WheelGame(730) = "030405070809"
    WheelGame(731) = "030405070810"
    WheelGame(732) = "030405070811"
    WheelGame(733) = "030405070812"
    WheelGame(734) = "030405070910"
    WheelGame(735) = "030405070911"
    WheelGame(736) = "030405070912"
    WheelGame(737) = "030405071011"
    WheelGame(738) = "030405071012"
    WheelGame(739) = "030405071112"
    WheelGame(740) = "030405080910"
    WheelGame(741) = "030405080911"
    WheelGame(742) = "030405080912"
    WheelGame(743) = "030405081011"
    WheelGame(744) = "030405081012"
    WheelGame(745) = "030405081112"
    WheelGame(746) = "030405091011"
    WheelGame(747) = "030405091012"
    WheelGame(748) = "030405091112"
    WheelGame(749) = "030405101112"
    WheelGame(750) = "030406070809"
    WheelGame(751) = "030406070810"
    WheelGame(752) = "030406070811"
    WheelGame(753) = "030406070812"
    WheelGame(754) = "030406070910"
    WheelGame(755) = "030406070911"
    WheelGame(756) = "030406070912"
    WheelGame(757) = "030406071011"
    WheelGame(758) = "030406071012"
    WheelGame(759) = "030406071112"
    WheelGame(760) = "030406080910"
    WheelGame(761) = "030406080911"
    WheelGame(762) = "030406080912"
    WheelGame(763) = "030406081011"
    WheelGame(764) = "030406081012"
    WheelGame(765) = "030406081112"
    WheelGame(766) = "030406091011"
    WheelGame(767) = "030406091012"
    WheelGame(768) = "030406091112"
    WheelGame(769) = "030406101112"
    WheelGame(770) = "030407080910"
    WheelGame(771) = "030407080911"
    WheelGame(772) = "030407080912"
    WheelGame(773) = "030407081011"
    WheelGame(774) = "030407081012"
    WheelGame(775) = "030407081112"
    WheelGame(776) = "030407091011"
    WheelGame(777) = "030407091012"
    WheelGame(778) = "030407091112"
    WheelGame(779) = "030407101112"
    WheelGame(780) = "030408091011"
    WheelGame(781) = "030408091012"
    WheelGame(782) = "030408091112"
    WheelGame(783) = "030408101112"
    WheelGame(784) = "030409101112"
    WheelGame(785) = "030506070809"
    WheelGame(786) = "030506070810"
    WheelGame(787) = "030506070811"
    WheelGame(788) = "030506070812"
    WheelGame(789) = "030506070910"
    WheelGame(790) = "030506070911"
    WheelGame(791) = "030506070912"
    WheelGame(792) = "030506071011"
    WheelGame(793) = "030506071012"
    WheelGame(794) = "030506071112"
    WheelGame(795) = "030506080910"
    WheelGame(796) = "030506080911"
    WheelGame(797) = "030506080912"
    WheelGame(798) = "030506081011"
    WheelGame(799) = "030506081012"
    WheelGame(800) = "030506081112"
    WheelGame(801) = "030506091011"
    WheelGame(802) = "030506091012"
    WheelGame(803) = "030506091112"
    WheelGame(804) = "030506101112"
    WheelGame(805) = "030507080910"
    WheelGame(806) = "030507080911"
    WheelGame(807) = "030507080912"
    WheelGame(808) = "030507081011"
    WheelGame(809) = "030507081012"
    WheelGame(810) = "030507081112"
    WheelGame(811) = "030507091011"
    WheelGame(812) = "030507091012"
    WheelGame(813) = "030507091112"
    WheelGame(814) = "030507101112"
    WheelGame(815) = "030508091011"
    WheelGame(816) = "030508091012"
    WheelGame(817) = "030508091112"
    WheelGame(818) = "030508101112"
    WheelGame(819) = "030509101112"
    WheelGame(820) = "030607080910"
    WheelGame(821) = "030607080911"
    WheelGame(822) = "030607080912"
    WheelGame(823) = "030607081011"
    WheelGame(824) = "030607081012"
    WheelGame(825) = "030607081112"
    WheelGame(826) = "030607091011"
    WheelGame(827) = "030607091012"
    WheelGame(828) = "030607091112"
    WheelGame(829) = "030607101112"
    WheelGame(830) = "030608091011"
    WheelGame(831) = "030608091012"
    WheelGame(832) = "030608091112"
    WheelGame(833) = "030608101112"
    WheelGame(834) = "030609101112"
    WheelGame(835) = "030708091011"
    WheelGame(836) = "030708091012"
    WheelGame(837) = "030708091112"
    WheelGame(838) = "030708101112"
    WheelGame(839) = "030709101112"
    WheelGame(840) = "030809101112"
    WheelGame(841) = "040506070809"
    WheelGame(842) = "040506070810"
    WheelGame(843) = "040506070811"
    WheelGame(844) = "040506070812"
    WheelGame(845) = "040506070910"
    WheelGame(846) = "040506070911"
    WheelGame(847) = "040506070912"
    WheelGame(848) = "040506071011"
    WheelGame(849) = "040506071012"
    WheelGame(850) = "040506071112"
    WheelGame(851) = "040506080910"
    WheelGame(852) = "040506080911"
    WheelGame(853) = "040506080912"
    WheelGame(854) = "040506081011"
    WheelGame(855) = "040506081012"
    WheelGame(856) = "040506081112"
    WheelGame(857) = "040506091011"
    WheelGame(858) = "040506091012"
    WheelGame(859) = "040506091112"
    WheelGame(860) = "040506101112"
    WheelGame(861) = "040507080910"
    WheelGame(862) = "040507080911"
    WheelGame(863) = "040507080912"
    WheelGame(864) = "040507081011"
    WheelGame(865) = "040507081012"
    WheelGame(866) = "040507081112"
    WheelGame(867) = "040507091011"
    WheelGame(868) = "040507091012"
    WheelGame(869) = "040507091112"
    WheelGame(870) = "040507101112"
    WheelGame(871) = "040508091011"
    WheelGame(872) = "040508091012"
    WheelGame(873) = "040508091112"
    WheelGame(874) = "040508101112"
    WheelGame(875) = "040509101112"
    WheelGame(876) = "040607080910"
    WheelGame(877) = "040607080911"
    WheelGame(878) = "040607080912"
    WheelGame(879) = "040607081011"
    WheelGame(880) = "040607081012"
    WheelGame(881) = "040607081112"
    WheelGame(882) = "040607091011"
    WheelGame(883) = "040607091012"
    WheelGame(884) = "040607091112"
    WheelGame(885) = "040607101112"
    WheelGame(886) = "040608091011"
    WheelGame(887) = "040608091012"
    WheelGame(888) = "040608091112"
    WheelGame(889) = "040608101112"
    WheelGame(890) = "040609101112"
    WheelGame(891) = "040708091011"
    WheelGame(892) = "040708091012"
    WheelGame(893) = "040708091112"
    WheelGame(894) = "040708101112"
    WheelGame(895) = "040709101112"
    WheelGame(896) = "040809101112"
    WheelGame(897) = "050607080910"
    WheelGame(898) = "050607080911"
    WheelGame(899) = "050607080912"
    WheelGame(900) = "050607081011"
    WheelGame(901) = "050607081012"
    WheelGame(902) = "050607081112"
    WheelGame(903) = "050607091011"
    WheelGame(904) = "050607091012"
    WheelGame(905) = "050607091112"
    WheelGame(906) = "050607101112"
    WheelGame(907) = "050608091011"
    WheelGame(908) = "050608091012"
    WheelGame(909) = "050608091112"
    WheelGame(910) = "050608101112"
    WheelGame(911) = "050609101112"
    WheelGame(912) = "050708091011"
    WheelGame(913) = "050708091012"
    WheelGame(914) = "050708091112"
    WheelGame(915) = "050708101112"
    WheelGame(916) = "050709101112"
    WheelGame(917) = "050809101112"
    WheelGame(918) = "060708091011"
    WheelGame(919) = "060708091012"
    WheelGame(920) = "060708091112"
    WheelGame(921) = "060708101112"
    WheelGame(922) = "060709101112"
    WheelGame(923) = "060809101112"
    WheelGame(924) = "070809101112"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub ninetwelve_Click()
    WheelText = "THESE 12 COMBINATIONS OF THE FIELD OF 9 NUMBERS' ASSURE A 3rd PRIZE WINNER" & Chr$(13) & "IF ANY 4 NUMBERS DRAWN FALL WITHIN THE FIELD OF 9 NUMBERS !"
    WheelGames = 12
    WheelNumbers = 9
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203040708"
    WheelGame(3) = "010203050809"
    WheelGame(4) = "010204060809"
    WheelGame(5) = "010205060709"
    WheelGame(6) = "010304060709"
    WheelGame(7) = "010305060708"
    WheelGame(8) = "010405070809"
    WheelGame(9) = "020304050709"
    WheelGame(10) = "020306070809"
    WheelGame(11) = "020405060708"
    WheelGame(12) = "030405060809"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub TennumbersTen_Click()
    WheelText = "THESE 10 COMBINATIONS OF THE FIELD OF 10 NUMBERS' ASSURE A 4th PRIZE WINNER" & Chr$(13) & "IF ANY 3 NUMBERS DRAWN FALL WITHIN THE FIELD OF 10 ! "
    WheelGames = 10
    WheelNumbers = 10
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010207080910"
    WheelGame(3) = "010307080910"
    WheelGame(4) = "010407080910"
    WheelGame(5) = "010507080910"
    WheelGame(6) = "010607080910"
    WheelGame(7) = "020304050607"
    WheelGame(8) = "020304050608"
    WheelGame(9) = "020304050609"
    WheelGame(10) = "020304050610"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub elevenfive_Click()
    WheelText = "THESE 5 COMBINATIONS OF THE FIELD OF 11 NUMBERS' ASSURE A 4th PRIZE WINNERS" & Chr$(13) & "IF ANY 4 NUMBERS DRAWN FALL WITHIN THE FIELD OF 11 !"
    WheelGames = 5
    WheelNumbers = 11
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010305070811"
    WheelGame(3) = "010305091011"
    WheelGame(4) = "020407080910"
    WheelGame(5) = "060708091011"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub twelvesix_Click()
    WheelText = "THESE 6 COMBINATIONS OF THE FIELD OF 12 NUMBERS' ASSURE A 3rd PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 12 NUMBERS!"
    WheelGames = 6
    WheelNumbers = 12
    ShowWheelInfo
    WheelGame(1) = "010204050912"
    WheelGame(2) = "010304060910"
    WheelGame(3) = "010507081011"
    WheelGame(4) = "020305061012"
    WheelGame(5) = "020607080911"
    WheelGame(6) = "030407081112"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub twelveeighteen_Click()
    WheelText = "THESE 18 COMBINATIONS OF THE FIELD OF 12 NUMBERS, ASSURE A 3rd PRIZE WINNER" & Chr$(13) & "IF ANY 5 NUMBERS DRAWN FALL WITHIN THE FIELD OF 12 !"
    WheelGames = 18
    WheelNumbers = 12
    ShowWheelInfo
    WheelGame(1) = "010203050811"
    WheelGame(2) = "010205080911"
    WheelGame(3) = "010205081112"
    WheelGame(4) = "010304060912"
    WheelGame(5) = "010304070912"
    WheelGame(6) = "010304091012"
    WheelGame(7) = "010306070912"
    WheelGame(8) = "010306091012"
    WheelGame(9) = "010307091012"
    WheelGame(10) = "020305080911"
    WheelGame(11) = "020305081112"
    WheelGame(12) = "020405060710"
    WheelGame(13) = "020406070810"
    WheelGame(14) = "020406071011"
    WheelGame(15) = "020508091112"
    WheelGame(16) = "040506070810"
    WheelGame(17) = "040506071011"
    WheelGame(18) = "040607081011"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub twelveforty_Click()
    WheelText = "THESE 42 COMBINATIONS OF THE FIELD OF 12 NUMBERS' ASSURE A 2nd PRIZE WINNER" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 12 !"
    WheelGames = 42
    WheelNumbers = 12
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010203071011"
    WheelGame(3) = "010203080912"
    WheelGame(4) = "010204070911"
    WheelGame(5) = "010204081012"
    WheelGame(6) = "010205070910"
    WheelGame(7) = "010205081112"
    WheelGame(8) = "010206070812"
    WheelGame(9) = "010206091011"
    WheelGame(10) = "010304070811"
    WheelGame(11) = "010304091012"
    WheelGame(12) = "010305070810"
    WheelGame(13) = "010305091112"
    WheelGame(14) = "010306070912"
    WheelGame(15) = "010306081011"
    WheelGame(16) = "010405070809"
    WheelGame(17) = "010405101112"
    WheelGame(18) = "010406071012"
    WheelGame(19) = "010406080911"
    WheelGame(20) = "010506071112"
    WheelGame(21) = "010506080910"
    WheelGame(22) = "020304071112"
    WheelGame(23) = "020304080910"
    WheelGame(24) = "020305071012"
    WheelGame(25) = "020305080911"
    WheelGame(26) = "020306070809"
    WheelGame(27) = "020306101112"
    WheelGame(28) = "020405070912"
    WheelGame(29) = "020405081011"
    WheelGame(30) = "020406070810"
    WheelGame(31) = "020406091112"
    WheelGame(32) = "020506070811"
    WheelGame(33) = "020506091012"
    WheelGame(34) = "030405070812"
    WheelGame(35) = "030405091011"
    WheelGame(36) = "030406070910"
    WheelGame(37) = "030406081112"
    WheelGame(38) = "030506070911"
    WheelGame(39) = "030506081012"
    WheelGame(40) = "040506071011"
    WheelGame(41) = "040506080912"
    WheelGame(42) = "070809101112"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub forteenforteen_Click()
    WheelText = "THESE 14 COMBINATIONS OF THE FIELD OF 14 NUMBERS' ASSURE A 3rd PRIZE WINNER  IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 14 !"
    WheelGames = 14
    WheelNumbers = 14
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010207080910"
    WheelGame(3) = "010304071112"
    WheelGame(4) = "010306081013"
    WheelGame(5) = "010507081114"
    WheelGame(6) = "010609121314"
    WheelGame(7) = "020305070913"
    WheelGame(8) = "020410111314"
    WheelGame(9) = "020506081112"
    WheelGame(10) = "020607101214"
    WheelGame(11) = "030408091214"
    WheelGame(12) = "030509101114"
    WheelGame(13) = "040508101213"
    WheelGame(14) = "040607091113"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub seveneleven_Click()
    WheelText = "THESE 11 COMBINATIONS OF THE FIELD OF 17 NUMBERS' ASSURE A 4th PRIZE WINNER" & Chr$(13) & "IF ANY 5 NUMBERS DRAWN FALL WITHIN A FIELD OF 17 NUMBERS !"
    WheelGames = 11
    WheelNumbers = 17
    ShowWheelInfo
    WheelGame(1) = "010305060709"
    WheelGame(2) = "010306111217"
    WheelGame(3) = "010506070911"
    WheelGame(4) = "010507091217"
    WheelGame(5) = "020408101316"
    WheelGame(6) = "020408141516"
    WheelGame(7) = "020810131415"
    WheelGame(8) = "030507091112"
    WheelGame(9) = "030507091117"
    WheelGame(10) = "041013141516"
    WheelGame(11) = "050607091217"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub ofnftg_Click()
    WheelText = "THESE 42 COMBINATIONS OF THE FIELD OF 18 NUMBERS' ASSURE A 3rd PRIZE WINNER !" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 18 !"
    WheelGames = 42
    WheelNumbers = 18
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010207111516"
    WheelGame(3) = "010208121317"
    WheelGame(4) = "010209101418"
    WheelGame(5) = "010307101518"
    WheelGame(6) = "010308111316"
    WheelGame(7) = "010309121417"
    WheelGame(8) = "010405070809"
    WheelGame(9) = "010405101112"
    WheelGame(10) = "010405131415"
    WheelGame(11) = "010405161718"
    WheelGame(12) = "010607121517"
    WheelGame(13) = "010608101318"
    WheelGame(14) = "010609111416"
    WheelGame(15) = "020306070809"
    WheelGame(16) = "020306101112"
    WheelGame(17) = "020306131415"
    WheelGame(18) = "020306161718"
    WheelGame(19) = "020407121318"
    WheelGame(20) = "020408101416"
    WheelGame(21) = "020409111517"
    WheelGame(22) = "020507101417"
    WheelGame(23) = "020508111518"
    WheelGame(24) = "020509121316"
    WheelGame(25) = "030407111317"
    WheelGame(26) = "030408121418"
    WheelGame(27) = "030409101516"
    WheelGame(28) = "030507121416"
    WheelGame(29) = "030508101517"
    WheelGame(30) = "030509111318"
    WheelGame(31) = "040607101316"
    WheelGame(32) = "040608111417"
    WheelGame(33) = "040609121518"
    WheelGame(34) = "050607111418"
    WheelGame(35) = "050608121516"
    WheelGame(36) = "050609101317"
    WheelGame(37) = "070809101112"
    WheelGame(38) = "070809131415"
    WheelGame(39) = "070809161718"
    WheelGame(40) = "101112131415"
    WheelGame(41) = "101112161718"
    WheelGame(42) = "131415161718"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

Private Sub twentyeight_Click()
    WheelText = "THESE 36 COMBINATIONS OF THE FIELD OF 28 NUMBERS' ASSURE A 3rd PRIZE WINNER !" & Chr$(13) & "IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 28 !"
    WheelGames = 36
    WheelNumbers = 28
    ShowWheelInfo
    WheelGame(1) = "010203040708"
    WheelGame(2) = "010203050910"
    WheelGame(3) = "010203192021"
    WheelGame(4) = "010204060910"
    WheelGame(5) = "010205061718"
    WheelGame(6) = "010304051718"
    WheelGame(7) = "010305060708"
    WheelGame(8) = "010419202122"
    WheelGame(9) = "020304061718"
    WheelGame(10) = "020405060708"
    WheelGame(11) = "020519202122"
    WheelGame(12) = "030405060910"
    WheelGame(13) = "030619202122"
    WheelGame(14) = "040506072122"
    WheelGame(15) = "070809101718"
    WheelGame(16) = "070809101921"
    WheelGame(17) = "070809102022"
    WheelGame(18) = "070817181921"
    WheelGame(19) = "070817182022"
    WheelGame(20) = "091017181921"
    WheelGame(21) = "091017182022"
    WheelGame(22) = "111213141516"
    WheelGame(23) = "111215162526"
    WheelGame(24) = "111223242728"
    WheelGame(25) = "111225262728"
    WheelGame(26) = "111315232527"
    WheelGame(27) = "111316242628"
    WheelGame(28) = "111415242528"
    WheelGame(29) = "111416232627"
    WheelGame(30) = "121315242627"
    WheelGame(31) = "121316232528"
    WheelGame(32) = "121415232628"
    WheelGame(33) = "121416242527"
    WheelGame(34) = "131415162728"
    WheelGame(35) = "131423242526"
    WheelGame(36) = "151623242526"
    onoffpanel
    ClearPanel
    clearcolor
End Sub

' (not attached to any menu in the original - kept for completeness)
Private Sub TenOfFive_Click()
    WheelText = "THESE 10 COMBINATIONS OF THE FIELD OF 5 NUMBERS' ASSURE A 3rd PRIZE WINNER !   IF THE 6 NUMBERS DRAWN FALL WITHIN THE FIELD OF 18 !"
    WheelGames = 10
    WheelNumbers = 5
    ShowWheelInfo
    WheelGame(1) = "010203040506"
    WheelGame(2) = "010207111516"
    WheelGame(3) = "010208121317"
    WheelGame(4) = "010209101418"
    WheelGame(5) = "010307101518"
    onoffpanel
    ClearPanel
    clearcolor
End Sub
