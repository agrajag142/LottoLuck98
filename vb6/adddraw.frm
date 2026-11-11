VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form adddraw 
   Caption         =   "Add Draw Data"
   BackColor       =   &H00400000&
   Icon            =   "adddraw.frx":0000
   LinkTopic       =   "Form1"
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   ClientHeight    =   7845
   StartUpPosition =   3
   ScaleHeight     =   7845
   ScaleWidth      =   10650
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   7680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton MENU4 
      Caption         =   "WHEELING SYSTEMS"
      Left            =   7200
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   64
      Appearance      =   0
   End
   Begin VB.CommandButton MENU3 
      Caption         =   "GAME DESIGNER"
      Left            =   4800
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   63
      Appearance      =   0
      ToolTipText     =   "Create your own game files"
   End
   Begin VB.CommandButton MENU2 
      Caption         =   "GRAPHING"
      Left            =   2400
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   62
      Appearance      =   0
   End
   Begin VB.CommandButton MENU1 
      Caption         =   "CHECK WINNINGS"
      Left            =   0
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   61
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Left            =   1200
      Top             =   960
      Width           =   6375
      Height          =   5535
      TabIndex        =   0
      BorderStyle     =   0
      Begin VB.CommandButton Command1 
         Caption         =   "EXIT"
         Left            =   4680
         Top             =   4680
         Width           =   1335
         Height          =   495
         TabIndex        =   13
      End
      Begin VB.TextBox year 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   1680
         Top             =   4680
         Width           =   1215
         Height          =   555
         TabIndex        =   11
         Alignment       =   2
         MaxLength       =   4
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
      Begin VB.TextBox month 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   960
         Top             =   4680
         Width           =   615
         Height          =   555
         TabIndex        =   10
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox day 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   4680
         Width           =   615
         Height          =   555
         TabIndex        =   9
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw8 
         BackColor       =   &H00808000&
         ForeColor       =   &H0000FFFF&
         Left            =   5400
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   8
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw7 
         BackColor       =   &H00808000&
         ForeColor       =   &H0000FFFF&
         Left            =   4680
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   7
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw6 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   3840
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   6
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw5 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   3120
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   5
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw4 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   2400
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   4
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw3 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   1680
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   3
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw2 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   960
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   2
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.TextBox Draw1 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   3840
         Width           =   615
         Height          =   555
         TabIndex        =   1
         Alignment       =   2
         MaxLength       =   2
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
      Begin VB.CommandButton ADD 
         Caption         =   "ADD"
         BackColor       =   &H00C0C0C0&
         Left            =   3000
         Top             =   4680
         Width           =   1575
         Height          =   495
         TabIndex        =   12
      End
      Begin VB.Label Message 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   5520
         Width           =   5775
         Height          =   375
         TabIndex        =   60
         BorderStyle     =   1
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin VB.Label draw 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   3240
         Width           =   5775
         Height          =   375
         TabIndex        =   59
         BorderStyle     =   1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   " 1"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   58
         BorderStyle     =   1
      End
      Begin VB.Label Label2 
         Caption         =   " 2"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   57
         BorderStyle     =   1
      End
      Begin VB.Label Label3 
         Caption         =   " 3"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   56
         BorderStyle     =   1
      End
      Begin VB.Label Label4 
         Caption         =   " 4"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   55
         BorderStyle     =   1
      End
      Begin VB.Label Label5 
         Caption         =   " 5"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   54
         BorderStyle     =   1
      End
      Begin VB.Label Label6 
         Caption         =   " 6"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3480
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   53
         BorderStyle     =   1
      End
      Begin VB.Label Label7 
         Caption         =   " 7"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3840
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   52
         BorderStyle     =   1
      End
      Begin VB.Label Label8 
         Caption         =   " 8"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   4200
         Top             =   720
         Width           =   255
         Height          =   255
         TabIndex        =   51
         BorderStyle     =   1
      End
      Begin VB.Label Label9 
         Caption         =   " 9"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   50
         BorderStyle     =   1
      End
      Begin VB.Label Label10 
         Caption         =   "10"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   49
         BorderStyle     =   1
      End
      Begin VB.Label Label11 
         Caption         =   "11"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   48
         BorderStyle     =   1
      End
      Begin VB.Label Label12 
         Caption         =   "12"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   47
         BorderStyle     =   1
      End
      Begin VB.Label Label13 
         Caption         =   "13"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   46
         BorderStyle     =   1
      End
      Begin VB.Label Label14 
         Caption         =   "14"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3480
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   45
         BorderStyle     =   1
      End
      Begin VB.Label Label15 
         Caption         =   "15"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3840
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   44
         BorderStyle     =   1
      End
      Begin VB.Label Label16 
         Caption         =   "16"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   4200
         Top             =   1080
         Width           =   255
         Height          =   255
         TabIndex        =   43
         BorderStyle     =   1
      End
      Begin VB.Label Label17 
         Caption         =   "17"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   42
         BorderStyle     =   1
      End
      Begin VB.Label Label18 
         Caption         =   "18"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   41
         BorderStyle     =   1
      End
      Begin VB.Label Label19 
         Caption         =   "19"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   40
         BorderStyle     =   1
      End
      Begin VB.Label Label20 
         Caption         =   "20"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   39
         BorderStyle     =   1
      End
      Begin VB.Label Label21 
         Caption         =   "21"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   38
         BorderStyle     =   1
      End
      Begin VB.Label Label22 
         Caption         =   "22"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3480
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   37
         BorderStyle     =   1
      End
      Begin VB.Label Label23 
         Caption         =   "23"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3840
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   36
         BorderStyle     =   1
      End
      Begin VB.Label Label24 
         Caption         =   "24"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   4200
         Top             =   1440
         Width           =   255
         Height          =   255
         TabIndex        =   35
         BorderStyle     =   1
      End
      Begin VB.Label Label25 
         Caption         =   "25"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   34
         BorderStyle     =   1
      End
      Begin VB.Label Label26 
         Caption         =   "26"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   33
         BorderStyle     =   1
      End
      Begin VB.Label Label27 
         Caption         =   "27"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   32
         BorderStyle     =   1
      End
      Begin VB.Label Label28 
         Caption         =   "28"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   31
         BorderStyle     =   1
      End
      Begin VB.Label Label29 
         Caption         =   "29"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   30
         BorderStyle     =   1
      End
      Begin VB.Label Label30 
         Caption         =   "30"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3480
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   29
         BorderStyle     =   1
      End
      Begin VB.Label Label31 
         Caption         =   "31"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3840
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   28
         BorderStyle     =   1
      End
      Begin VB.Label Label32 
         Caption         =   "32"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   4200
         Top             =   1800
         Width           =   255
         Height          =   255
         TabIndex        =   27
         BorderStyle     =   1
      End
      Begin VB.Label Label33 
         Caption         =   "33"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   26
         BorderStyle     =   1
      End
      Begin VB.Label Label34 
         Caption         =   "34"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   25
         BorderStyle     =   1
      End
      Begin VB.Label Label35 
         Caption         =   "35"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   24
         BorderStyle     =   1
      End
      Begin VB.Label Label36 
         Caption         =   "36"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   23
         BorderStyle     =   1
      End
      Begin VB.Label Label37 
         Caption         =   "37"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   22
         BorderStyle     =   1
      End
      Begin VB.Label Label38 
         Caption         =   "38"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3480
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   21
         BorderStyle     =   1
      End
      Begin VB.Label Label39 
         Caption         =   "39"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3840
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   20
         BorderStyle     =   1
      End
      Begin VB.Label Label40 
         Caption         =   "40"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   4200
         Top             =   2160
         Width           =   255
         Height          =   255
         TabIndex        =   19
         BorderStyle     =   1
      End
      Begin VB.Label Label41 
         Caption         =   "41"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   1680
         Top             =   2520
         Width           =   255
         Height          =   255
         TabIndex        =   18
         BorderStyle     =   1
      End
      Begin VB.Label Label42 
         Caption         =   "42"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2040
         Top             =   2520
         Width           =   255
         Height          =   255
         TabIndex        =   17
         BorderStyle     =   1
      End
      Begin VB.Label Label43 
         Caption         =   "43"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2400
         Top             =   2520
         Width           =   255
         Height          =   255
         TabIndex        =   16
         BorderStyle     =   1
      End
      Begin VB.Label Label44 
         Caption         =   "44"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   2760
         Top             =   2520
         Width           =   255
         Height          =   255
         TabIndex        =   15
         BorderStyle     =   1
      End
      Begin VB.Label Label45 
         Caption         =   "45"
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000C0&
         Left            =   3120
         Top             =   2520
         Width           =   255
         Height          =   255
         TabIndex        =   14
         BorderStyle     =   1
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         Left            =   1560
         Top             =   600
         Width           =   3015
         Height          =   2295
         Shape           =   4
         FillColor       =   &H00404000&
         FillStyle       =   0
      End
   End
   Begin VB.Label Label46 
      Caption         =   "Don't key single numbers with an 0 in front eg. key 2 not 02 - this screen is temp."
      BackColor       =   &H0000FFFF&
      Left            =   1200
      Top             =   600
      Width           =   6255
      Height          =   255
      TabIndex        =   65
   End
End
Attribute VB_Name = "adddraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' adddraw - "ADD DRAW" screen: type in the latest Tattslotto result and append it to tattslot.txt

Public Sub Closedown()
    Unload Board
    Unload Wheelsystem
    Unload Occurance
    Unload adddraw
    Unload Lotto
End Sub

Private Sub LoadDraws()
    Dim f As Integer
    DrawIdx = 0
    f = FreeFile
    Open "tattslot.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, DrawLine(DrawIdx)
        DrawIdx = DrawIdx + 1
    Loop
    DrawIdx = DrawIdx - 1
    Close #f
End Sub

Private Sub LayoutMenus()
    MENU1.Width = adddraw.Width / 4
    MENU2.Width = adddraw.Width / 4
    MENU3.Width = adddraw.Width / 4
    MENU4.Width = adddraw.Width / 4
    MENU2.Left = MENU1.Width
    MENU3.Left = MENU1.Width * 2
    MENU4.Left = MENU1.Width * 3
End Sub

Private Sub Form_Load()
    AddDrawOpen = 1
    DrawIdx = 0
    adddraw.Left = SysInfo1.WorkAreaLeft
    adddraw.Width = SysInfo1.WorkAreaWidth
    adddraw.Top = SysInfo1.WorkAreaTop
    adddraw.Height = SysInfo1.WorkAreaHeight
    LayoutMenus
    Frame1.Top = Int((adddraw.Height - Frame1.Height) / 2)
    Frame1.Left = Int((adddraw.Width - Frame1.Width) / 2)

    LoadDraws
    LastDraw = DrawIdx - 1

    ' pre-fill the boxes from the most recent draw
    LastDrawNumber = CInt(Mid(DrawLine(LastDraw), 2, 4))
    year.Text = CStr(Mid(DrawLine(LastDraw), 7, 4))
    month.Text = CStr(Mid(DrawLine(LastDraw), 11, 2))
    day.Text = CStr(Mid(DrawLine(LastDraw), 13, 2) + 7)       ' next week's draw
    draw.Caption = "DRAW: " & CStr(LastDrawNumber + 2)
End Sub

Private Sub Add_Click()
    Dim f As Integer
    On Error GoTo AddErr
    SaveString = "B" & CStr(LastDrawNumber + 2) & "," & year.Text & month.Text & day.Text & "," & _
        Draw1.Text & "," & Draw2.Text & "," & Draw3.Text & "," & Draw4.Text & "," & _
        Draw5.Text & "," & Draw6.Text & "," & Draw7.Text & "," & Draw8.Text
    If SaveString <> "" Then
        f = FreeFile
        Open "tattslot.txt" For Append As #f
        Print #f, SaveString
        Close #f
    End If
    Message.Caption = "UPDATE COMPLETE"
    ADD.Enabled = False
    LoadDraws
    LastDraw = DrawIdx
AddErr:
End Sub

Private Sub Command1_Click()
    Lotto.Show
    adddraw.Hide
End Sub

' Enter moves to the next box
Private Sub day_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then month.SetFocus
End Sub
Private Sub month_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then year.SetFocus
End Sub
Private Sub year_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ADD.SetFocus
End Sub
Private Sub Draw1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw2.SetFocus
End Sub
Private Sub Draw2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw3.SetFocus
End Sub
Private Sub Draw3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw4.SetFocus
End Sub
Private Sub Draw4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw5.SetFocus
End Sub
Private Sub Draw5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw6.SetFocus
End Sub
Private Sub Draw6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw7.SetFocus
End Sub
Private Sub Draw7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Draw8.SetFocus
End Sub
Private Sub Draw8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then day.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Closedown
End Sub

Private Sub MENU1_Click()
    Lotto.Show
    adddraw.Hide
End Sub
Private Sub MENU2_Click()
    Occurance.Show
    adddraw.Hide
End Sub
Private Sub MENU3_Click()
    Board.Show
    adddraw.Hide
End Sub
Private Sub MENU4_Click()
    Wheelsystem.Show
    adddraw.Hide
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
    Frame1.Top = Int((adddraw.Height - Frame1.Height) / 2)
    Frame1.Left = Int((adddraw.Width - Frame1.Width) / 2)
End Sub
