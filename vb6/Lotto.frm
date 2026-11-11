VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Lotto 
   Caption         =   "Lottoluck '98 (c) Justin Little "
   BackColor       =   &H00400000&
   Picture         =   "Lotto.frx":0000
   BorderStyle     =   1
   Icon            =   "Lotto.frx":75742
   LinkTopic       =   "form1"
   MaxButton       =   0
   MinButton       =   0
   ClientLeft      =   1830
   ClientTop       =   2070
   ClientWidth     =   15150
   ClientHeight    =   10860
   PaletteMode     =   1
   ScaleHeight     =   10860
   ScaleWidth      =   15150
   Begin VB.CommandButton MENU5 
      Caption         =   "CHECK WINNINGS"
      Left            =   9600
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   1
      Appearance      =   0
      ToolTipText     =   "Check your game files"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALL +"
      Left            =   3720
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   51
      ToolTipText     =   "All to when 45 numbers introduced"
   End
   Begin MCI.MMControl MMControl1 
      Left            =   1560
      Top             =   8040
      Width           =   4215
      Height          =   330
      TabIndex        =   33
      _ExtentX        =   7435
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "WaveAudio"
      FileName        =   "sound1.wav"
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   10800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.FileListBox GameListbox 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Left            =   12000
      Top             =   240
      Width           =   2895
      Height          =   3990
      Visible         =   0
      TabIndex        =   26
      Pattern         =   "*.lot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton MENU4 
      Caption         =   "WHEELING SYSTEMS"
      Left            =   7200
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   4
      Appearance      =   0
      ToolTipText     =   "Wheel a selection of numbers"
   End
   Begin VB.CommandButton MENU3 
      Caption         =   "GAME DESIGNER"
      Left            =   4800
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   2
      Appearance      =   0
      ToolTipText     =   "Create your own game files"
   End
   Begin VB.CommandButton MENU2 
      Caption         =   "GRAPHING"
      BackColor       =   &H00C0C0C0&
      Left            =   2400
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   3
      Appearance      =   0
      ToolTipText     =   "See how draws have been performing"
   End
   Begin VB.CommandButton MENU1 
      Caption         =   "ADD DRAW"
      Left            =   0
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   0
      Appearance      =   0
      ToolTipText     =   "Add latest draw to database"
   End
   Begin ComctlLib.Slider Slider2 
      Left            =   360
      Top             =   6360
      Width           =   5055
      Height          =   255
      TabIndex        =   25
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      SmallChange     =   1
      Min             =   1
      Max             =   10
      ToolTipText     =   "Choose number of draws to compare with"
   End
   Begin ComctlLib.Slider Slider1 
      Left            =   6000
      Top             =   6120
      Width           =   5055
      Height          =   255
      TabIndex        =   23
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      SmallChange     =   1
      Min             =   0
      Max             =   10
      ToolTipText     =   "Draw history"
   End
   Begin VB.ListBox WinList 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FFFF&
      Left            =   5760
      Top             =   3240
      Width           =   5535
      Height          =   1710
      TabIndex        =   20
      Sorted          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ALL"
      Left            =   4560
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   10
      ToolTipText     =   "All to draw number 1"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "52"
      Left            =   2880
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   9
      ToolTipText     =   "One year"
   End
   Begin VB.CommandButton Command7 
      Caption         =   "8"
      Left            =   2040
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   8
      ToolTipText     =   "Two months"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "4"
      Left            =   1200
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   7
      ToolTipText     =   "One month"
   End
   Begin VB.CommandButton Command5 
      Caption         =   "1"
      Left            =   360
      Top             =   5880
      Width           =   855
      Height          =   375
      TabIndex        =   6
      ToolTipText     =   "One week"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Left            =   360
      Top             =   5160
      Width           =   5055
      Height          =   255
      TabIndex        =   12
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6840
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.tattersalls.com.au"
      URL             =   "http://www.tattersalls.com.au/files/tattslot.zip"
      Document        =   "/files/tattslot.zip"
   End
   Begin VB.Label eight 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000FF&
      Left            =   10440
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   59
      BorderStyle     =   1
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
   Begin VB.Label seven 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000FF&
      Left            =   9960
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   58
      BorderStyle     =   1
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
   Begin VB.Label six 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   9120
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   57
      BorderStyle     =   1
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
   Begin VB.Label five 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   8520
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   56
      BorderStyle     =   1
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
   Begin VB.Label four 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7920
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   55
      BorderStyle     =   1
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
   Begin VB.Label three 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   7320
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   54
      BorderStyle     =   1
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
   Begin VB.Label two 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6720
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   53
      BorderStyle     =   1
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
   Begin VB.Label one 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Left            =   6120
      Top             =   5520
      Width           =   495
      Height          =   420
      TabIndex        =   52
      BorderStyle     =   1
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
   Begin VB.Label Label24 
      Caption         =   "DIVISION"
      BackColor       =   &H00404000&
      ForeColor       =   &H0000FFFF&
      Left            =   600
      Top             =   1440
      Width           =   1455
      Height          =   255
      TabIndex        =   50
      Alignment       =   2
   End
   Begin VB.Label Label23 
      Caption         =   "WINS"
      BackColor       =   &H00404000&
      ForeColor       =   &H0000FFFF&
      Left            =   2160
      Top             =   1440
      Width           =   735
      Height          =   255
      TabIndex        =   49
      Alignment       =   2
   End
   Begin VB.Label Label22 
      Caption         =   "Stat. Value"
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Left            =   3000
      Top             =   1440
      Width           =   855
      Height          =   255
      TabIndex        =   48
      Alignment       =   2
   End
   Begin VB.Label Label21 
      Caption         =   "ODDS"
      BackColor       =   &H00404000&
      ForeColor       =   &H0000FFFF&
      Left            =   3960
      Top             =   1440
      Width           =   1215
      Height          =   255
      TabIndex        =   47
      Alignment       =   2
   End
   Begin VB.Label Duration 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Left            =   3120
      Top             =   4440
      Width           =   1935
      Height          =   495
      TabIndex        =   46
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "Time taken to compare"
   End
   Begin VB.Label Wins1 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   3000
      Top             =   1800
      Width           =   855
      Height          =   375
      TabIndex        =   45
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Wins2 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   3000
      Top             =   2280
      Width           =   855
      Height          =   375
      TabIndex        =   44
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Wins3 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   3000
      Top             =   2760
      Width           =   855
      Height          =   375
      TabIndex        =   43
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Wins4 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   3000
      Top             =   3240
      Width           =   855
      Height          =   375
      TabIndex        =   42
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Wins5 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   3000
      Top             =   3720
      Width           =   855
      Height          =   375
      TabIndex        =   41
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label18 
      Caption         =   "2 With 1 Sup."
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   4680
      Width           =   1455
      Height          =   375
      TabIndex        =   40
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label17 
      Caption         =   "3 With 0 Sup."
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   4200
      Width           =   1455
      Height          =   375
      TabIndex        =   39
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label16 
      Caption         =   "Division 5"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   3720
      Width           =   1455
      Height          =   375
      TabIndex        =   38
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label15 
      Caption         =   "Division 4"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   3240
      Width           =   1455
      Height          =   375
      TabIndex        =   37
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label14 
      Caption         =   "Division 3"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   2760
      Width           =   1455
      Height          =   375
      TabIndex        =   36
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label13 
      Caption         =   "Division 2"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   2280
      Width           =   1455
      Height          =   375
      TabIndex        =   35
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label12 
      Caption         =   "Division 1"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   600
      Top             =   1800
      Width           =   1455
      Height          =   375
      TabIndex        =   34
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label9 
      Caption         =   "1:733"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   3960
      Top             =   3240
      Width           =   1215
      Height          =   375
      TabIndex        =   31
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label8 
      Caption         =   "1:36,690"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   3960
      Top             =   2760
      Width           =   1215
      Height          =   375
      TabIndex        =   30
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label7 
      Caption         =   "1:678,755"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   3960
      Top             =   2280
      Width           =   1215
      Height          =   375
      TabIndex        =   29
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label GameloadDetails 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Left            =   240
      Top             =   7080
      Width           =   5295
      Height          =   255
      TabIndex        =   27
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Winfrequency 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FF00&
      Left            =   720
      Top             =   6720
      Width           =   4455
      Height          =   255
      TabIndex        =   24
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label twonumbersup 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   4680
      Width           =   735
      Height          =   375
      TabIndex        =   22
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000A&
      BorderColor     =   &H00404000&
      Left            =   480
      Top             =   1440
      Width           =   4815
      Height          =   2655
      BorderWidth     =   3
   End
   Begin VB.Label threenumbers 
      BackColor       =   &H00004000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   4200
      Width           =   735
      Height          =   375
      TabIndex        =   21
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Outlay 
      BackColor       =   &H000000C0&
      ForeColor       =   &H0080FFFF&
      Left            =   960
      Top             =   960
      Width           =   3855
      Height          =   375
      TabIndex        =   19
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label5 
      Caption         =   "         CHECK WINNINGS"
      BackColor       =   &H80000005&
      ForeColor       =   &H0000FFFF&
      Left            =   0
      Top             =   480
      Width           =   5295
      Height          =   375
      TabIndex        =   18
      Alignment       =   2
      BackStyle       =   0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   3720
      Width           =   735
      Height          =   375
      TabIndex        =   17
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   3240
      Width           =   735
      Height          =   375
      TabIndex        =   16
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   2760
      Width           =   735
      Height          =   375
      TabIndex        =   15
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label19 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   2280
      Width           =   735
      Height          =   375
      TabIndex        =   14
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label20 
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   2160
      Top             =   1800
      Width           =   735
      Height          =   375
      TabIndex        =   13
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   360
      Top             =   5400
      Width           =   5055
      Height          =   495
      TabIndex        =   11
      BorderStyle     =   1
      WordWrap        =   -1  'True
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   6000
      Top             =   5160
      Width           =   5055
      Height          =   255
      TabIndex        =   5
      BorderStyle     =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      Left            =   5760
      Top             =   5040
      Width           =   5535
      Height          =   1575
      Shape           =   4
      FillColor       =   &H00404040&
      FillStyle       =   0
   End
   Begin VB.Label Label6 
      Caption         =   "1:8,145,060"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   3960
      Top             =   1800
      Width           =   1215
      Height          =   375
      TabIndex        =   28
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Label Label10 
      Caption         =   "1:298"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FFFF&
      Left            =   3960
      Top             =   3720
      Width           =   1215
      Height          =   375
      TabIndex        =   32
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      Left            =   120
      Top             =   360
      Width           =   5535
      Height          =   6735
      Shape           =   4
      BorderStyle     =   0
      FillColor       =   &H00400000&
      FillStyle       =   0
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu loadgame 
         Caption         =   "Load Game File"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Add 
         Caption         =   "Add Latest Draw"
      End
      Begin VB.Menu CouponScreen 
         Caption         =   "Coupon Entry"
         Enabled         =   0
      End
      Begin VB.Menu Compare 
         Caption         =   "Check Winnings"
      End
      Begin VB.Menu Graph1 
         Caption         =   "Graphing"
      End
      Begin VB.Menu BoardMenu 
         Caption         =   "Random Number Game Creator"
      End
      Begin VB.Menu soundonof 
         Caption         =   "Sound"
         Begin VB.Menu Soundon 
            Caption         =   "On"
         End
         Begin VB.Menu SoundOff 
            Caption         =   "Off"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu wheeling 
         Caption         =   "Wheeling Systems"
      End
   End
   Begin VB.Menu Internet 
      Caption         =   "Internet"
      Begin VB.Menu updatedraw 
         Caption         =   "Download Tattslot.txt update"
      End
   End
   Begin VB.Menu Helpme 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu What 
         Caption         =   "No Help Yet!"
      End
   End
End
Attribute VB_Name = "Lotto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
'  Lotto  -  main "CHECK WINNINGS" screen of Lottoluck '98
'  Reconstructed from the compiled Xlotto.exe (VB5 native code).
'  Control names, captions, strings and logic come from the binary.
'  Comments marked "RECON:" flag places where the exact original wording
'  could not be recovered and an equivalent was written instead.
' ---------------------------------------------------------------------------

' SysInfo1 / Slider / ProgressBar / MCI / Inet / CommonDialog were all used
' late-bound in the original (dispatch IDs below were decoded from the EXE).

'===========================================================================
'  General procedures
'===========================================================================

Public Sub loadgamedata()
    Dim f As Integer
    Dim GameIndex As Variant           ' RECON: never initialised in the original (always Empty -> 0)
    GameCount = 1
    DrawLine(GameIndex) = File         ' original quirk: "File" resolves to the File menu control (writes "True" into DrawLine(0))
    f = FreeFile
    Open "number.lot" For Input As #f
    Do While Not EOF(f)
        Line Input #f, Games(GameCount)
        GameCount = GameCount + 1
    Loop
    GameCount = GameCount - 1
    Close #f
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & " draws. "
End Sub

Public Sub Update()
    ' Splits the current draw line (DrawLine(DrawIdx)) into Drawn(1..8).
    ' Line format:  A0001,19720624,37,39,1,14,18,31,30,0
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
    For idx = 1 To 8
        If Len(Drawn(idx)) < 2 Then Drawn(idx) = "0" & Drawn(idx)
    Next idx
End Sub

Private Sub GetTattsUpdate()
    Inet1.Execute "http://www.tattersalls.com.au" & "/files/tattslot.zip", "GET"
End Sub

Public Sub PlayWinSound1()
    MMControl1.Command = "close"
    MMControl1.FileName = "sound1.wav"
    MMControl1.Command = "open"
    MMControl1.Command = "stop"
    MMControl1.Command = "prev"
    MMControl1.Command = "play"
End Sub

Public Sub PlayWinSound2()
    MMControl1.Command = "close"
    MMControl1.FileName = "sound2.wav"
    MMControl1.Command = "open"
    MMControl1.Command = "stop"
    MMControl1.Command = "prev"
    MMControl1.Command = "play"
End Sub

Public Sub PlayWinSound3()
    MMControl1.Command = "close"
    MMControl1.FileName = "sound3.wav"
    MMControl1.Command = "open"
    MMControl1.Command = "stop"
    MMControl1.Command = "prev"
    MMControl1.Command = "play"
End Sub

Public Sub ODDsCalc()
    ' "Stat. Value" column: expected number of wins for the games/draws selected
    Wins1.Caption = CStr(Int((DrawsToCompare - 1) / (8145060 / GameCount) * 10) / 10)
    Wins2.Caption = CStr(Int((DrawsToCompare - 1) / (678755 / GameCount) * 10) / 10)
    Wins3.Caption = CStr(Int((DrawsToCompare - 1) / (36690 / GameCount) * 10) / 10)
    Wins4.Caption = CStr(Int((DrawsToCompare - 1) / (733 / GameCount) * 10) / 10)
    Wins5.Caption = CStr(Int((DrawsToCompare - 1) / (298 / GameCount) * 10) / 10)
End Sub

Public Sub Closedown()
    MMControl1.Command = "close"
    Unload Board
    Unload Wheelsystem
    Unload Occurance
    Unload adddraw
    Unload Lotto
End Sub

Public Sub Compareit()
    ' Checks every loaded game against the last (DrawsToCompare - 1) draws.
    Dim hits() As Integer, supps() As Integer
    Dim before(6) As Integer, divs(6) As Integer
    Dim startTime As Variant, drawsChecked As Variant
    Dim lastWin As String
    Dim d As Integer, x As Integer, j As Integer, p As Integer
    Dim winCount As Integer
    Dim dv As Variant
    Dim num As String

    ReDim hits(1250): ReDim supps(1250)
    startTime = Time
    lastWin = ""
    drawsChecked = DrawsToCompare - 1
    DrawsCompared = CInt(drawsChecked)
    WinList.Clear
    DrawIdx = LastDraw
    Update
    For j = 0 To 6: divs(j) = 0: Next j
    ProgressBar1.Min = 0
    ProgressBar1.Max = CSng(DrawsToCompare)
    ODDsCalc

    ' clear the history used by the graphing screen
    For x = 1 To UBound(WinHistory, 1) + 1
        For j = 0 To 4
            WinHistory(x - 1, j) = 0
        Next j
    Next x

    d = 0
    Do Until d = DrawsToCompare - 1
        d = d + 1
        ProgressBar1.Value = d + 1
        For x = 1 To GameCount
            hits(x) = 0
            supps(x) = 0
        Next x
        For x = 1 To GameCount
            For dv = 0 To 6
                before(dv) = divs(dv)
            Next dv
            ' each game is 12 characters: six 2-digit numbers
            For p = 1 To 11 Step 2
                num = Mid(Games(x), p, 2)
                For j = 1 To 6
                    If num = Drawn(j) Then hits(x) = hits(x) + 1
                Next j
                For j = 7 To 8
                    If num = Drawn(j) Then supps(x) = supps(x) + 1
                Next j
            Next p
            If hits(x) = 2 And supps(x) = 1 Then divs(6) = divs(6) + 1      ' 2 + supp
            If hits(x) = 3 And supps(x) = 0 Then divs(0) = divs(0) + 1      ' 3, no supp
            If hits(x) = 3 And supps(x) > 0 Then divs(5) = divs(5) + 1      ' Division 5
            If hits(x) = 4 And supps(x) >= 0 Then divs(4) = divs(4) + 1     ' Division 4
            If hits(x) = 5 And supps(x) = 0 Then divs(3) = divs(3) + 1      ' Division 3
            If hits(x) = 5 And supps(x) > 0 Then divs(2) = divs(2) + 1      ' Division 2
            If hits(x) = 6 And supps(x) >= 0 Then divs(1) = divs(1) + 1     ' Division 1
            For dv = 1 To 5
                If before(dv) <> divs(dv) Then
                    WinList.AddItem "Division " & dv & " win on Draw " & Mid(DrawLine(DrawIdx), 1, 5) & " Game " & x
                    winCount = winCount + 1
                    WinHistory(d - 1, dv - 1) = 6 - dv
                    If lastWin = "" Then lastWin = Right$(Mid$(DrawLine(DrawIdx), 1, 5), 4)
                End If
            Next dv
        Next x
        If DrawIdx <= 1 Then DrawIdx = 2
        DrawIdx = DrawIdx - 1
        Update
    Loop

    Label2.Caption = CStr(divs(5))
    Label3.Caption = CStr(divs(4))
    Label4.Caption = CStr(divs(3))
    Label19.Caption = CStr(divs(2))
    Label20.Caption = CStr(divs(1))
    threenumbers.Caption = CStr(divs(0))
    twonumbersup.Caption = CStr(divs(6))

    startTime = Time - startTime
    Duration.Caption = Format(startTime, "HH:MM:SS")

    If Soundon.Checked = True And divs(1) > 0 Then PlayWinSound1
    If Soundon.Checked = True And divs(1) = 0 And divs(2) > 0 Then PlayWinSound2     ' RECON: exact test
    If Soundon.Checked = True And divs(1) = 0 And divs(2) = 0 And divs(3) > 0 Then PlayWinSound3

    If winCount > 0 Then
        Winfrequency.Caption = "Win: 1 in " & Int(1 / (winCount / drawsChecked)) & _
            "  Next Win: " & (Int(1 / (winCount / drawsChecked)) * 2 + lastWin)
    Else
        Winfrequency.Caption = ""
    End If
End Sub

' Re-centres the screen inside the Windows work area and lays out the
' five navigation buttons across the top (shared by Form_Load / DisplayChanged)
Private Sub LayoutMenus()
    MENU1.Width = Lotto.Width / 5
    MENU2.Width = Lotto.Width / 5
    MENU3.Width = Lotto.Width / 5
    MENU4.Width = Lotto.Width / 5
    MENU5.Width = Lotto.Width / 5
    MENU5.Left = MENU1.Width
    MENU3.Left = MENU1.Width * 2
    MENU2.Left = MENU1.Width * 3
    MENU4.Left = MENU1.Width * 4
End Sub

Private Sub ShowOutlay(ByVal draws As Variant)
    Outlay.Caption = "Outlay = $" & CStr(Int(GameCount * 6 / 14 * draws))  ' RECON: rounding
End Sub

Private Sub ShowComparing()
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & _
        " draws. Equiv to " & CStr((DrawsToCompare - 1) * GameCount) & " Games."
End Sub

Private Sub ShowDrawNumber()
    one.Caption = Drawn(1)
    two.Caption = Drawn(2)
    three.Caption = Drawn(3)
    four.Caption = Drawn(4)
    five.Caption = Drawn(5)
    six.Caption = Drawn(6)
    seven.Caption = Drawn(7)
    eight.Caption = Drawn(8)
    Label1.Caption = "DRAW NUMBER: " & Mid(DrawLine(DrawIdx), 2, 4)
End Sub

'===========================================================================
'  Form events
'===========================================================================

Private Sub Form_Load()
    Dim f As Integer
    DrawsCompared = 0
    frmSplash.Show
    Lotto.Hide
    GamesWanted = 1
    DrawIdx = 0

    Lotto.Left = SysInfo1.WorkAreaLeft
    Lotto.Width = SysInfo1.WorkAreaWidth
    Lotto.Top = SysInfo1.WorkAreaTop
    Lotto.Height = SysInfo1.WorkAreaHeight

    If Screen.Width >= 15360 Then GameListbox.Visible = True    ' 1024+ pixels wide: show the game file list
    GameListbox.Width = Screen.Width - GameListbox.Left
    GameListbox.Height = Lotto.Height - 1220
    LayoutMenus

    ' load the Tattslotto results history
    f = FreeFile
    Open "tattslot.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, DrawLine(DrawIdx)
        DrawIdx = DrawIdx + 1
    Loop
    DrawIdx = DrawIdx - 1
    LastDraw = DrawIdx - 1
    Close #f

    Slider1.Min = 1
    Slider1.Max = LastDraw
    Slider1.Value = LastDraw

    loadgamedata
    Update
    ShowDrawNumber

    DrawsToCompare = 2
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & " draws. "
    Slider2.Min = 1
    Slider2.Max = LastDraw
    Slider2.Value = 1
End Sub

Private Sub Form_Activate()
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & " draws. "
    Slider1.Min = 1
    Slider1.Max = LastDraw
    Slider1.Value = LastDraw
    Update
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
    LayoutMenus
End Sub

'---------------------------------------------------------------------------
'  Draw-range buttons:  1 / 4 / 8 / 52 / ALL / ALL+
'---------------------------------------------------------------------------
Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawsToCompare = 2
    ShowComparing
    ODDsCalc
    ShowOutlay 1
End Sub
Private Sub Command5_Click()
    If GameCount > 0 Then Compareit
    ShowOutlay 1
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawsToCompare = 5
    ODDsCalc
    ShowComparing
    ShowOutlay 4
End Sub
Private Sub Command6_Click()
    If GameCount > 0 Then Compareit
    ShowOutlay 5          ' (sic - original multiplies by 5 here and by 4 on MouseMove)
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawsToCompare = 9
    ShowComparing
    ODDsCalc
    ShowOutlay 8
End Sub
Private Sub Command7_Click()
    If GameCount > 0 Then Compareit
    ShowOutlay 10         ' (sic)
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawsToCompare = 53
    ShowComparing
    ODDsCalc
    ShowOutlay 52
End Sub
Private Sub Command8_Click()
    If GameCount > 0 Then Compareit
    ShowOutlay 50         ' (sic)
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' ALL - every draw back to draw number 1
    DrawIdx = LastDraw
    DrawsToCompare = LastDraw + 1
    ShowComparing
    ODDsCalc
    ShowOutlay LastDraw
End Sub
Private Sub Command9_Click()
    DrawIdx = LastDraw
    Update
    DrawsToCompare = DrawIdx + 1
    If GameCount > 0 Then Compareit
    ShowOutlay LastDraw
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' ALL + : back to when 45 numbers were introduced (659 draws from the start)
    DrawIdx = LastDraw - 659
    DrawsToCompare = (LastDraw - 659) + 1
    ShowComparing
    ODDsCalc
    ShowOutlay LastDraw - 659
End Sub
Private Sub Command1_Click()
    DrawIdx = LastDraw - 659
    Update
    DrawsToCompare = DrawIdx + 1
    If GameCount > 0 Then Compareit
    ShowOutlay LastDraw - 659
End Sub

'---------------------------------------------------------------------------
'  Sliders
'---------------------------------------------------------------------------
Private Sub Slider1_Change()
    DrawIdx = Slider1.Value
    Update
    ShowDrawNumber
End Sub
Private Sub Slider1_Click()
    DrawIdx = Slider1.Value
    Update
    ShowDrawNumber
End Sub
Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawIdx = Slider1.Value
    Update
    ShowDrawNumber
End Sub

Private Sub Slider2_Change()
    DrawsToCompare = Slider2.Value + 1
    ShowComparing
End Sub
Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawsToCompare = Slider2.Value + 1
    ShowComparing
    ODDsCalc
    ShowOutlay Slider2.Value
End Sub
Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GameCount > 0 Then Compareit
    ShowOutlay Slider2.Value
End Sub

'---------------------------------------------------------------------------
'  Game files
'---------------------------------------------------------------------------
Private Sub GameListbox_DblClick()
    Dim f As Integer
    On Error Resume Next
    GameloadDetails.Caption = "GAME LOADED: " & GameListbox.FileName
    GameCount = 1
    f = FreeFile
    Open GameListbox.FileName For Input As #f
    Do While Not EOF(f)
        Line Input #f, Games(GameCount)
        GameCount = GameCount + 1
    Loop
    GameCount = GameCount - 1
    Close #f
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & " draws. "
End Sub

Private Sub loadgame_Click()
    Dim f As Integer
    On Error Resume Next
    CommonDialog1.DialogTitle = "Load A Game File"
    CommonDialog1.Filter = "*.lot"
    CommonDialog1.FileName = "*.lot"
    CommonDialog1.ShowOpen
    If Right$(CommonDialog1.FileName, 4) = ".lot" And Left$(CommonDialog1.FileName, 1) <> "*" Then
        GameloadDetails.Caption = "GAME LOADED: " & CommonDialog1.FileName
        GameCount = 1
        f = FreeFile
        Open CommonDialog1.FileName For Input As #f
        Do While Not EOF(f)
            Line Input #f, Games(GameCount)
            GameCount = GameCount + 1
        Loop
        GameCount = GameCount - 1
        Close #f
    End If
    Label11.Caption = "Comparing " & CStr(GameCount) & " games with last " & CStr(DrawsToCompare - 1) & " draws. "
End Sub

'---------------------------------------------------------------------------
'  Menus / navigation
'---------------------------------------------------------------------------
Private Sub Compare_Click()
    If GameCount > 0 Then Compareit
End Sub

Private Sub About_Click()
    frmSplash.Show
End Sub

Private Sub Add_Click()
    adddraw.Show
    Unload Lotto
End Sub

Private Sub BoardMenu_Click()
    Board.Show
    Lotto.Visible = False
End Sub

Private Sub Graph1_Click()
    Occurance.Visible = True
    Lotto.Visible = False
End Sub

Private Sub wheeling_Click()
    Wheelsystem.Show
    Lotto.Hide
End Sub

Private Sub MENU1_Click()
    adddraw.Show
    Lotto.Hide
End Sub

Private Sub MENU2_Click()
    Occurance.Show
    Lotto.Hide
End Sub

Private Sub MENU3_Click()
    Board.Show
    Lotto.Hide
End Sub

Private Sub MENU4_Click()
    Wheelsystem.Show
    Lotto.Hide
End Sub

Private Sub Soundon_Click()
    Soundon.Checked = True
    SoundOff.Checked = False
End Sub

Private Sub SoundOff_Click()
    Soundon.Checked = False
    SoundOff.Checked = True
End Sub

Private Sub Quit_Click()
    Unload Lotto
    Unload Board
    Unload Wheelsystem
    Unload Occurance
    Unload adddraw
End Sub

'---------------------------------------------------------------------------
'  Internet update
'---------------------------------------------------------------------------
Private Sub updatedraw_Click()
    Dim b() As Byte
    Dim f As Integer
    b() = Inet1.OpenURL("http://www.tattersalls.com.au/files/tattslot.zip", icByteArray)
    f = FreeFile
    Open "tattslot.zip" For Binary Access Write As #f
    Put #f, , b()
    Close #f
    MsgBox "Download Complete"
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    ' (empty in the original)
End Sub
