VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Board 
   Caption         =   "Game Designer"
   BackColor       =   &H80000008&
   BorderStyle     =   1
   Icon            =   "Board.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   ClientLeft      =   150
   ClientTop       =   1005
   ClientWidth     =   12435
   ClientHeight    =   8565
   StartUpPosition =   3
   ScaleHeight     =   8565
   ScaleWidth      =   12435
   Begin VB.CommandButton ClearListbox 
      Caption         =   "CLEAR "
      Left            =   9000
      Top             =   6720
      Width           =   2535
      Height          =   375
      TabIndex        =   48
   End
   Begin VB.Frame Frame3 
      Caption         =   "MANUAL ENTRY"
      Left            =   120
      Top             =   3240
      Width           =   4455
      Height          =   2175
      TabIndex        =   35
      Begin VB.TextBox Select6 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   3720
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   47
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Select5 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   3000
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   46
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Select4 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   2280
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   45
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Select3 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   1560
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   44
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Select2 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   840
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   43
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Select1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Left            =   120
         Top             =   600
         Width           =   615
         Height          =   495
         TabIndex        =   42
         Alignment       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         Caption         =   "6"
         Left            =   3720
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   41
         Alignment       =   2
      End
      Begin VB.Label Label14 
         Caption         =   "5"
         Left            =   3000
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   40
         Alignment       =   2
      End
      Begin VB.Label Label13 
         Caption         =   "4"
         Left            =   2280
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   39
         Alignment       =   2
      End
      Begin VB.Label Label12 
         Caption         =   "3"
         Left            =   1560
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   38
         Alignment       =   2
      End
      Begin VB.Label Label11 
         Caption         =   "2"
         Left            =   840
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   37
         Alignment       =   2
      End
      Begin VB.Label Label10 
         Caption         =   "1"
         Left            =   120
         Top             =   360
         Width           =   615
         Height          =   255
         TabIndex        =   36
         Alignment       =   2
      End
   End
   Begin VB.CommandButton MENU5 
      Caption         =   "GAME DESIGNER"
      Left            =   4800
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   33
      Appearance      =   0
      ToolTipText     =   "Create your own game files"
   End
   Begin ComctlLib.Slider SliderMin 
      Left            =   240
      Top             =   2040
      Width           =   2535
      Height          =   495
      TabIndex        =   19
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      SmallChange     =   1
      Min             =   21
      Max             =   254
      ToolTipText     =   "Minimum Total"
   End
   Begin VB.Frame Frame2 
      Caption         =   "PREFERRED"
      BackColor       =   &H00C0C0C0&
      Left            =   4800
      Top             =   2280
      Width           =   3975
      Height          =   2295
      TabIndex        =   11
      ToolTipText     =   "Enter numbers to be present in all games"
      Begin VB.OptionButton Option3 
         Caption         =   "POWER/DUE"
         Left            =   240
         Top             =   1800
         Width           =   1695
         Height          =   255
         TabIndex        =   34
         ToolTipText     =   "Coldest since 45 introduced"
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CLEAR ALL"
         Left            =   2400
         Top             =   1800
         Width           =   1335
         Height          =   315
         TabIndex        =   31
         ToolTipText     =   "Clear all preferred boxes"
      End
      Begin VB.CommandButton Command6 
         Left            =   3120
         Top             =   960
         Width           =   615
         Height          =   195
         TabIndex        =   30
         ToolTipText     =   "CLEAR"
      End
      Begin VB.CommandButton Command5 
         Left            =   2400
         Top             =   960
         Width           =   615
         Height          =   195
         TabIndex        =   29
         ToolTipText     =   "CLEAR"
      End
      Begin VB.CommandButton Command4 
         Left            =   1680
         Top             =   960
         Width           =   615
         Height          =   195
         TabIndex        =   28
         ToolTipText     =   "CLEAR"
      End
      Begin VB.CommandButton Command3 
         Left            =   960
         Top             =   960
         Width           =   615
         Height          =   195
         TabIndex        =   27
         ToolTipText     =   "CLEAR"
      End
      Begin VB.CommandButton Command2 
         Left            =   240
         Top             =   960
         Width           =   615
         Height          =   195
         TabIndex        =   26
         ToolTipText     =   "CLEAR"
      End
      Begin VB.OptionButton Option2 
         Caption         =   "HOTTEST"
         Left            =   240
         Top             =   1560
         Width           =   1095
         Height          =   255
         TabIndex        =   25
         ToolTipText     =   "Hottest since 45 introduced"
      End
      Begin VB.OptionButton Option1 
         Caption         =   "COLDEST"
         Left            =   240
         Top             =   1320
         Width           =   1095
         Height          =   255
         TabIndex        =   24
         ToolTipText     =   "Coldest since 45 introduced"
      End
      Begin VB.TextBox Prefer5 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   3120
         Top             =   360
         Width           =   615
         Height          =   615
         TabIndex        =   16
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Prefer4 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   2400
         Top             =   360
         Width           =   615
         Height          =   615
         TabIndex        =   15
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Prefer3 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   1680
         Top             =   360
         Width           =   615
         Height          =   615
         TabIndex        =   14
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Prefer2 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   960
         Top             =   360
         Width           =   615
         Height          =   615
         TabIndex        =   13
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Prefer1 
         BackColor       =   &H00404000&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   360
         Width           =   615
         Height          =   615
         TabIndex        =   12
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FILTERING"
      Left            =   4800
      Top             =   960
      Width           =   3975
      Height          =   1215
      TabIndex        =   8
      Begin VB.CheckBox Check1 
         Caption         =   "No Duplicates"
         Left            =   1080
         Top             =   360
         Width           =   1935
         Height          =   255
         Enabled         =   0
         TabIndex        =   32
      End
      Begin VB.CheckBox Thirteen 
         Caption         =   "13"
         Left            =   240
         Top             =   720
         Width           =   1215
         Height          =   255
         TabIndex        =   10
      End
      Begin VB.CheckBox Seven 
         Caption         =   "7"
         Left            =   240
         Top             =   360
         Width           =   1215
         Height          =   255
         TabIndex        =   9
      End
   End
   Begin ComctlLib.Slider Slider1 
      Left            =   240
      Top             =   840
      Width           =   2535
      Height          =   495
      TabIndex        =   7
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      SmallChange     =   1
      Min             =   1
      Max             =   1000
      SelStart        =   70
      Value           =   70
      ToolTipText     =   "Number Of Games To Create"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE GAMES"
      Left            =   4800
      Top             =   4920
      Width           =   3975
      Height          =   495
      TabIndex        =   6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "Create games with filters"
   End
   Begin VB.CommandButton MENU1 
      Caption         =   "ADD DRAW"
      Left            =   0
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   5
      Appearance      =   0
   End
   Begin VB.CommandButton MENU2 
      Caption         =   "CHECK WINNINGS"
      Left            =   2400
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   4
      Appearance      =   0
   End
   Begin VB.CommandButton MENU3 
      Caption         =   "GRAPHING"
      Left            =   7200
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   3
      Appearance      =   0
   End
   Begin VB.CommandButton MENU4 
      Caption         =   "WHEELING SYSTEMS"
      Left            =   9600
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   2
      Appearance      =   0
   End
   Begin VB.ListBox GameList 
      BackColor       =   &H00404000&
      ForeColor       =   &H0000FFFF&
      Left            =   9000
      Top             =   960
      Width           =   2535
      Height          =   5520
      TabIndex        =   0
      TabStop         =   0
      ToolTipText     =   "Generated Number List"
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   9240
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.Slider SliderMax 
      Left            =   240
      Top             =   2520
      Width           =   2535
      Height          =   495
      TabIndex        =   20
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      SmallChange     =   1
      Min             =   22
      Max             =   255
      SelStart        =   255
      Value           =   255
      ToolTipText     =   "Maximum Total"
   End
   Begin VB.Label Label3 
      Caption         =   "SUM"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FF80&
      Left            =   10440
      Top             =   720
      Width           =   1095
      Height          =   255
      TabIndex        =   23
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Caption         =   "NUMBERS"
      BackColor       =   &H00404000&
      ForeColor       =   &H0080FF80&
      Left            =   9000
      Top             =   720
      Width           =   1455
      Height          =   255
      TabIndex        =   22
      BorderStyle     =   1
      Alignment       =   2
   End
   Begin VB.Shape Shape2 
      Left            =   4680
      Top             =   720
      Width           =   4215
      Height          =   4095
      Shape           =   4
      FillColor       =   &H00404000&
      FillStyle       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "SUM OF NUMBERS FILTER"
      ForeColor       =   &H00000000&
      Left            =   240
      Top             =   1680
      Width           =   3975
      Height          =   375
      TabIndex        =   21
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label SumMax 
      Caption         =   "255"
      BackColor       =   &H00004080&
      ForeColor       =   &H0000FFFF&
      Left            =   2760
      Top             =   2520
      Width           =   1455
      Height          =   495
      TabIndex        =   18
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "Maximum Total"
   End
   Begin VB.Label SumMin 
      Caption         =   "21"
      BackColor       =   &H00004080&
      ForeColor       =   &H0000FFFF&
      Left            =   2760
      Top             =   2040
      Width           =   1455
      Height          =   495
      TabIndex        =   17
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "Minimum Total"
   End
   Begin VB.Label ToGet 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Left            =   2760
      Top             =   840
      Width           =   1455
      Height          =   495
      TabIndex        =   1
      BorderStyle     =   1
      Alignment       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "Number Of Games To Create"
   End
   Begin VB.Shape Shape1 
      Left            =   120
      Top             =   1560
      Width           =   4215
      Height          =   1575
      Shape           =   4
      FillColor       =   &H00404000&
      FillStyle       =   0
   End
   Begin VB.Shape Shape3 
      Left            =   120
      Top             =   720
      Width           =   4215
      Height          =   735
      Shape           =   4
      FillColor       =   &H00404000&
      FillStyle       =   0
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu Quittomain 
         Caption         =   "Quit to main"
      End
   End
   Begin VB.Menu Games 
      Caption         =   "Games"
      Begin VB.Menu fourteen 
         Caption         =   "14"
      End
      Begin VB.Menu twentyeight 
         Caption         =   "28"
      End
      Begin VB.Menu fortytwo 
         Caption         =   "42"
      End
      Begin VB.Menu Seventy 
         Caption         =   "70"
      End
      Begin VB.Menu hundred 
         Caption         =   "100"
      End
      Begin VB.Menu twonineeight 
         Caption         =   "298"
      End
      Begin VB.Menu seventhreethree 
         Caption         =   "733"
      End
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Board - "GAME DESIGNER": generates random games with filters and saves them as .lot files

Public Sub GetSixBalls()
    ' picks six different random balls into Balls(1..6), honouring the PREFERRED boxes
    ' and the "7" / "13" exclusion check boxes, then pads and sorts them
    Dim i As Integer, j As Integer, k As Integer, good As Integer
    Dim t As String
    Do
        For i = 1 To 6
            Balls(i) = CStr(Int(Rnd * 45) + 1)
            For k = 1 To CInt(Format(Time, "ss"))       ' stir the generator
                ClickedNum = Rnd
            Next k
        Next i
        If Prefer1.Text <> "" Then Balls(1) = Prefer1.Text
        If Prefer2.Text <> "" Then Balls(2) = Prefer2.Text
        If Prefer3.Text <> "" Then Balls(3) = Prefer3.Text
        If Prefer4.Text <> "" Then Balls(4) = Prefer4.Text
        If Prefer5.Text <> "" Then Balls(5) = Prefer5.Text
        good = 0
        For i = 1 To 6
            If Seven.Value = 1 And Val(Balls(i)) = 7 Then good = good + 1
            If Thirteen.Value = 1 And Val(Balls(i)) = 13 Then good = good + 1
            For j = 1 To 6
                If Val(Balls(i)) = Val(Balls(j)) Then good = good + 1
            Next j
        Next i
    Loop Until good = 6           ' each ball matched only itself -> six unique, none excluded
    For i = 1 To 6
        If Len(Balls(i)) = 1 Then Balls(i) = "0" & Balls(i)
    Next i
    For i = 1 To 6
        For j = 1 To 6
            If Balls(i) < Balls(j) Then
                t = Balls(i): Balls(i) = Balls(j): Balls(j) = t
            End If
        Next j
    Next i
End Sub

Private Function BallSum() As Integer
    BallSum = Val(Balls(1)) + Val(Balls(2)) + Val(Balls(3)) + Val(Balls(4)) + Val(Balls(5)) + Val(Balls(6))
End Function

Public Sub ReadGamelistToSavestring()
    ' converts the list box (" 01 02 03 04 05 06   21") back to 12-digit game lines
    Dim i As Integer, l As String
    SaveString = ""
    i = 0
    Do While GameList.List(i) <> ""
        l = GameList.List(i)
        If SaveString <> "" Then SaveString = SaveString & Chr$(10) & Chr$(13)
        SaveString = SaveString & Mid$(l, 2, 2) & Mid$(l, 5, 2) & Mid$(l, 8, 2) & _
                     Mid$(l, 11, 2) & Mid$(l, 14, 2) & Mid$(l, 17, 2)
        Games(i + 1) = Mid$(l, 2, 2) & Mid$(l, 5, 2) & Mid$(l, 8, 2) & Mid$(l, 11, 2) & Mid$(l, 14, 2) & Mid$(l, 17, 2)
        i = i + 1
    Loop
    GameCount = i
End Sub

Private Sub ClearGameList()
    ReadGamelistToSavestring
End Sub

Public Sub ManualCheck()
    If Len(Select1.Text) = 1 Then Select1.Text = "0" & Select1.Text
    If Len(Select2.Text) = 1 Then Select2.Text = "0" & Select2.Text
    If Len(Select3.Text) = 1 Then Select3.Text = "0" & Select3.Text
    If Len(Select4.Text) = 1 Then Select4.Text = "0" & Select4.Text
    If Len(Select5.Text) = 1 Then Select5.Text = "0" & Select5.Text
    If Len(Select6.Text) = 1 Then Select6.Text = "0" & Select6.Text
End Sub

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

Public Sub PREFER_CLEAR()
    ' puts Chosen(1..5) into the five PREFERRED boxes
    Dim k As Integer
    Prefer1.Text = "": Prefer2.Text = "": Prefer3.Text = "": Prefer4.Text = "": Prefer5.Text = ""
    For k = 1 To 5
        If Len(Chosen(k)) = 1 Then Chosen(k) = "0" & Chosen(k)
        If Prefer1.Text = "" Then
            Prefer1.Text = Chosen(k)
        ElseIf Prefer2.Text = "" Then
            Prefer2.Text = Chosen(k)
        ElseIf Prefer3.Text = "" Then
            Prefer3.Text = Chosen(k)
        ElseIf Prefer4.Text = "" Then
            Prefer4.Text = Chosen(k)
        ElseIf Prefer5.Text = "" Then
            Prefer5.Text = Chosen(k)
        End If
    Next k
End Sub

' counts how often each ball came up (all 8 balls) for the last "draws" draws
Private Sub CountSince(counts() As Integer, ByVal draws As Integer)
    Dim d As Integer, n As Integer, j As Integer
    For n = 1 To 45: counts(n) = 0: Next n
    DrawIdx = LastDraw
    For d = 1 To draws
        Update
        For n = 1 To 45
            For j = 1 To 8
                If n = Val(Drawn(j)) Then counts(n) = counts(n) + 1
            Next j
        Next n
        DrawIdx = DrawIdx - 1
    Next d
End Sub

'---------------------------------------------------------------------------
'  Form
'---------------------------------------------------------------------------
Private Sub LayoutForm()
    Board.Left = SysInfo1.WorkAreaLeft
    Board.Width = SysInfo1.WorkAreaWidth
    Board.Top = SysInfo1.WorkAreaTop
    Board.Height = SysInfo1.WorkAreaHeight
    GameList.Height = Board.Height - GameList.Top - 1200 - 375
    ClearListbox.Top = GameList.Top + GameList.Height
    MENU1.Width = SysInfo1.WorkAreaWidth / 5
    MENU2.Width = SysInfo1.WorkAreaWidth / 5
    MENU3.Width = SysInfo1.WorkAreaWidth / 5
    MENU4.Width = SysInfo1.WorkAreaWidth / 5
    MENU5.Width = SysInfo1.WorkAreaWidth / 5
    MENU2.Left = MENU1.Width
    MENU5.Left = MENU1.Width * 2
    MENU3.Left = MENU1.Width * 3
    MENU4.Left = MENU1.Width * 4
End Sub

Private Sub Form_Load()
    GamesWanted = 70
    ToGet.Caption = CStr(GamesWanted)
    LayoutForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Closedown
End Sub

Private Sub SysInfo1_DisplayChanged()
    LayoutForm
End Sub

'---------------------------------------------------------------------------
'  CREATE GAMES
'---------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim x As Integer, s As Integer
    For x = 1 To 1250
        Games(x) = ""
    Next x
    GameList.Clear
    For x = 1 To Val(ToGet.Caption)
        Do
            GetSixBalls
            Games(x) = Balls(1) & Balls(2) & Balls(3) & Balls(4) & Balls(5) & Balls(6) & "   " & CStr(BallSum())
            s = Val(Right$(Games(x), 3))
        Loop Until s >= SliderMin.Value And s <= SliderMax.Value
        GameList.AddItem " " & Balls(1) & " " & Balls(2) & " " & Balls(3) & " " & Balls(4) & " " & _
                         Balls(5) & " " & Balls(6) & "   " & CStr(BallSum())
    Next x
    GameCount = CLng(ToGet.Caption)
    SaveString = ""
    For x = 1 To Val(ToGet.Caption) - 1
        SaveString = SaveString & Left$(Games(x), 12) & Chr$(10) & Chr$(13)
    Next x
    SaveString = SaveString & Left$(Games(CLng(ToGet.Caption)), 12)
End Sub

Private Sub Saveas_Click()
    Dim f As Integer
    On Error GoTo SaveErr
    ReadGamelistToSavestring
    If SaveString = "" Then Exit Sub
    CommonDialog1.DialogTitle = "Save A File"
    CommonDialog1.Filter = "*.lot"
    CommonDialog1.FileName = "*.lot"
    CommonDialog1.Flags = 2                      ' cdlOFNOverwritePrompt
    CommonDialog1.ShowSave
    If SaveString <> "" Then
        If Right$(CommonDialog1.FileName, 4) <> ".lot" Then
            CommonDialog1.FileName = CommonDialog1.FileName & ".lot"
        End If
        f = FreeFile
        Open CommonDialog1.FileName For Output As #f
        Print #f, SaveString
        Close #f
    End If
SaveErr:
End Sub

'---------------------------------------------------------------------------
'  Manual entry boxes (Enter moves on, the 6th adds the game to the list)
'---------------------------------------------------------------------------
Private Sub Select1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Select2.SetFocus: ManualCheck
End Sub
Private Sub Select2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Select3.SetFocus: ManualCheck
End Sub
Private Sub Select3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Select4.SetFocus: ManualCheck
End Sub
Private Sub Select4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Select5.SetFocus: ManualCheck
End Sub
Private Sub Select5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Select6.SetFocus: ManualCheck
End Sub
Private Sub Select6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ManualCheck
        GameList.AddItem " " & Select1.Text & " " & Select2.Text & " " & Select3.Text & " " & _
            Select4.Text & " " & Select5.Text & " " & Select6.Text & "   " & _
            CStr(Val(Select1.Text) + Val(Select2.Text) + Val(Select3.Text) + Val(Select4.Text) + Val(Select5.Text) + Val(Select6.Text))
        Select1.SetFocus
        Select1.Text = ""
        Select2.Text = ""
        Select3.Text = ""
        Select4.Text = ""
        Select5.Text = ""
        Select6.Text = ""
        ReadGamelistToSavestring
    End If
End Sub

'---------------------------------------------------------------------------
'  PREFERRED numbers
'---------------------------------------------------------------------------
Private Sub Prefer1_Change()
    If Prefer1.Text = Prefer2.Text Then Prefer1.Text = ""
    If Prefer1.Text = Prefer3.Text Then Prefer1.Text = ""
    If Prefer1.Text = Prefer4.Text Then Prefer1.Text = ""
    If Prefer1.Text = Prefer5.Text Then Prefer1.Text = ""
End Sub
Private Sub Prefer2_Change()
    If Prefer2.Text = Prefer1.Text Then Prefer2.Text = ""
    If Prefer2.Text = Prefer3.Text Then Prefer2.Text = ""
    If Prefer2.Text = Prefer4.Text Then Prefer2.Text = ""
    If Prefer2.Text = Prefer5.Text Then Prefer2.Text = ""
End Sub
Private Sub Prefer3_Change()
    If Prefer3.Text = Prefer1.Text Then Prefer3.Text = ""
    If Prefer3.Text = Prefer2.Text Then Prefer3.Text = ""
    If Prefer3.Text = Prefer4.Text Then Prefer3.Text = ""
    If Prefer3.Text = Prefer5.Text Then Prefer3.Text = ""
End Sub
Private Sub Prefer4_Change()
    If Prefer4.Text = Prefer1.Text Then Prefer4.Text = ""
    If Prefer4.Text = Prefer2.Text Then Prefer4.Text = ""
    If Prefer4.Text = Prefer3.Text Then Prefer4.Text = ""
    If Prefer4.Text = Prefer5.Text Then Prefer4.Text = ""
End Sub
Private Sub Prefer5_Change()
    If Prefer5.Text = Prefer1.Text Then Prefer5.Text = ""
    If Prefer5.Text = Prefer2.Text Then Prefer5.Text = ""
    If Prefer5.Text = Prefer3.Text Then Prefer5.Text = ""
    If Prefer5.Text = Prefer4.Text Then Prefer5.Text = ""
End Sub

Private Sub Prefer1_LostFocus()
    If Len(Prefer1.Text) = 1 Then Prefer1.Text = "0" & Prefer1.Text
End Sub
Private Sub Prefer2_LostFocus()
    If Len(Prefer2.Text) = 1 Then Prefer2.Text = "0" & Prefer2.Text
End Sub
Private Sub Prefer3_LostFocus()
    If Len(Prefer3.Text) = 1 Then Prefer3.Text = "0" & Prefer3.Text
End Sub
Private Sub Prefer4_LostFocus()
    If Len(Prefer4.Text) = 1 Then Prefer4.Text = "0" & Prefer4.Text
End Sub
Private Sub Prefer5_LostFocus()
    If Len(Prefer5.Text) = 1 Then Prefer5.Text = "0" & Prefer5.Text
End Sub

' the small CLEAR buttons under each box
Private Sub Command2_Click()
    Prefer1.Text = ""
End Sub
Private Sub Command3_Click()
    Prefer2.Text = ""
End Sub
Private Sub Command4_Click()
    Prefer3.Text = ""
End Sub
Private Sub Command5_Click()
    Prefer4.Text = ""
End Sub
Private Sub Command6_Click()
    Prefer5.Text = ""
End Sub
Private Sub Command7_Click()
    Prefer1.Text = ""
    Prefer2.Text = ""
    Prefer3.Text = ""
    Prefer4.Text = ""
    Prefer5.Text = ""
End Sub

Private Sub Option1_Click()
    ' COLDEST since 45 numbers introduced
    Dim counts(45) As Integer
    Dim c As Integer, n As Integer, k As Integer
    CountSince counts, LastDraw - 659
    k = 1
    For c = 0 To 1000
        For n = 1 To 45
            If counts(n) = c And k <= 45 Then Chosen(k) = CStr(n): k = k + 1
        Next n
    Next c
    PREFER_CLEAR
    Option1.Value = False
End Sub

Private Sub Option2_Click()
    ' HOTTEST since 45 numbers introduced
    Dim counts(45) As Integer
    Dim c As Integer, n As Integer, k As Integer
    CountSince counts, LastDraw - 659
    k = 1
    For c = 1000 To 0 Step -1
        For n = 1 To 45
            If counts(n) = c And k <= 45 Then Chosen(k) = CStr(n): k = k + 1
        Next n
    Next c
    PREFER_CLEAR
    Option2.Value = False
End Sub

Private Sub Option3_Click()
    ' POWER/DUE: balls drawn less often recently than their long-term average
    Dim allTime(45) As Integer, since45(45) As Integer, power(45) As Integer
    Dim c As Integer, n As Integer, k As Integer
    CountSince allTime, LastDraw
    CountSince since45, LastDraw - 659
    For n = 1 To 45
        power(n) = CInt((allTime(n) / LastDraw) * 100 - (since45(n) / (LastDraw - 659)) * 100)
    Next n
    k = 1
    For c = 100 To -100 Step -1
        For n = 1 To 45
            If power(n) = c And k <= 45 Then Chosen(k) = CStr(n): k = k + 1
        Next n
    Next c
    PREFER_CLEAR
    Option3.Value = False
End Sub

'---------------------------------------------------------------------------
'  Number of games / sum range
'---------------------------------------------------------------------------
Private Sub SetGames(ByVal n As Integer)
    ToGet.Caption = CStr(n)
    Slider1.Value = CLng(ToGet.Caption)
End Sub
Private Sub fourteen_Click()
    SetGames 14
End Sub
Private Sub twentyeight_Click()
    SetGames 28
End Sub
Private Sub fortytwo_Click()
    SetGames 42
End Sub
Private Sub Seventy_Click()
    SetGames 70
End Sub
Private Sub hundred_Click()
    SetGames 100
End Sub
Private Sub twonineeight_Click()
    SetGames 298
End Sub
Private Sub seventhreethree_Click()
    SetGames 733
End Sub

Private Sub Slider1_Change()
    ToGet.Caption = CStr(Slider1.Value)
End Sub
Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToGet.Caption = CStr(Slider1.Value)
End Sub
Private Sub SliderMin_Change()
    SumMin.Caption = CStr(SliderMin.Value)
End Sub
Private Sub SliderMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SumMin.Caption = CStr(SliderMin.Value)
    If SliderMax.Value - 1 < SliderMin.Value Then SliderMax.Value = SliderMin.Value + 1
End Sub
Private Sub SliderMax_Change()
    SumMax.Caption = CStr(SliderMax.Value)
End Sub
Private Sub SliderMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SumMax.Caption = CStr(SliderMax.Value)
    If SliderMin.Value + 1 > SliderMax.Value Then SliderMin.Value = SliderMax.Value - 1
End Sub

Private Sub ClearListbox_Click()
    GameList.Clear
End Sub

'---------------------------------------------------------------------------
'  Navigation
'---------------------------------------------------------------------------
Private Sub MENU1_Click()
    adddraw.Show
    Board.Hide
End Sub
Private Sub MENU2_Click()
    Lotto.Show
    Board.Hide
End Sub
Private Sub MENU3_Click()
    Occurance.Show
    Board.Hide
End Sub
Private Sub MENU4_Click()
    Wheelsystem.Show
    Board.Hide
End Sub
Private Sub Quittomain_Click()
    Lotto.Show
    Board.Hide
End Sub
