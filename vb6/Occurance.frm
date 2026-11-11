VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form occurance 
   Caption         =   "Graphing"
   BackColor       =   &H00000080&
   BorderStyle     =   1
   Icon            =   "Occurance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   ClientLeft      =   2175
   ClientTop       =   1635
   ClientWidth     =   12075
   ClientHeight    =   7770
   PaletteMode     =   1
   ScaleHeight     =   7770
   ScaleWidth      =   12075
   Begin VB.OptionButton Option9 
      Caption         =   "Game wins (Must use check winnings screen)"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   4800
      Top             =   6720
      Width           =   4575
      Height          =   255
      TabIndex        =   15
      TabStop         =   0
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Sum of draws since 45 Numbers Intoduced Excl Supps"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   4800
      Top             =   6480
      Width           =   4575
      Height          =   255
      TabIndex        =   14
      TabStop         =   0
   End
   Begin VB.CommandButton MENU5 
      Caption         =   "GRAPHING"
      BackColor       =   &H00C0C0C0&
      Left            =   7200
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   13
      Appearance      =   0
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   4920
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton MENU4 
      Caption         =   "WHEELING SYSTEMS"
      Left            =   9600
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   12
      Appearance      =   0
   End
   Begin VB.CommandButton MENU3 
      Caption         =   "GAME DESIGNER"
      Left            =   4800
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   11
      Appearance      =   0
      ToolTipText     =   "Create your own game files"
   End
   Begin VB.CommandButton MENU2 
      Caption         =   "CHECK WINNINGS"
      Left            =   2400
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   10
      Appearance      =   0
   End
   Begin VB.CommandButton MENU1 
      Caption         =   "ADD DRAW"
      Left            =   0
      Top             =   0
      Width           =   2415
      Height          =   255
      TabIndex        =   9
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0
      Interval        =   2000
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Left            =   0
      Top             =   5400
      Width           =   9015
      Height          =   255
      TabIndex        =   8
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin MSChart20Lib.MSChart Chart1 
      Left            =   0
      Top             =   240
      Width           =   9375
      Height          =   4335
      TabIndex        =   7
      _ExtentX        =   16536
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Weeks Since Last Drawn"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   120
      Top             =   6480
      Width           =   4695
      Height          =   255
      TabIndex        =   0
      TabStop         =   0
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Drawn as a Supps  Since 45 Numbers Introduced"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   4800
      Top             =   6240
      Width           =   4575
      Height          =   255
      TabIndex        =   6
      TabStop         =   0
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Only Supplementary Numbers From All Drawn"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   120
      Top             =   6240
      Width           =   4695
      Height          =   255
      TabIndex        =   5
      TabStop         =   0
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Since 45 Numbers Introduced Excluding Supps"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   4800
      Top             =   6000
      Width           =   4575
      Height          =   255
      TabIndex        =   4
      TabStop         =   0
   End
   Begin VB.OptionButton Option3 
      Caption         =   "All Drawn Excluding Supplementary Numbers"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   4800
      Top             =   5760
      Width           =   4575
      Height          =   255
      TabIndex        =   3
      TabStop         =   0
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Since 45 Numbers Introduced Including Supps"
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   120
      Top             =   6000
      Width           =   4695
      Height          =   255
      TabIndex        =   2
      TabStop         =   0
   End
   Begin VB.OptionButton Option1 
      Caption         =   "All Drawn Including Supplementary "
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Left            =   120
      Top             =   5760
      Width           =   4695
      Height          =   255
      TabIndex        =   1
      TabStop         =   0
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Gameloadgraph 
         Caption         =   "Load Game"
         Enabled         =   0
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
         Enabled         =   0
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit To Main"
      End
   End
   Begin VB.Menu Type 
      Caption         =   "Graph Type"
      Begin VB.Menu autocycle 
         Caption         =   "Auto Cycle"
      End
      Begin VB.Menu d2graphs 
         Caption         =   "2D Graphs"
         Begin VB.Menu Area 
            Caption         =   "2D Area"
         End
         Begin VB.Menu twod 
            Caption         =   "2D Bar"
         End
         Begin VB.Menu line2d 
            Caption         =   "2D Line"
         End
         Begin VB.Menu Step2d 
            Caption         =   "2D Step"
         End
      End
      Begin VB.Menu d3graphs 
         Caption         =   "3D Graphs"
         Begin VB.Menu Area3d 
            Caption         =   "3D Area"
         End
         Begin VB.Menu threed 
            Caption         =   "3D Bar"
         End
         Begin VB.Menu Line3d 
            Caption         =   "3D Line"
         End
         Begin VB.Menu Step3d 
            Caption         =   "3D Step"
         End
      End
   End
End
Attribute VB_Name = "Occurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Occurance - "GRAPHING" screen: frequency / overdue / sum graphs drawn with MSChart

Private Const DRAW_45_INTRODUCED = "B0413"     ' first draw with 45 balls

Public Sub Update()
    ' identical to Lotto.Update: split DrawLine(DrawIdx) into Drawn(1..8)
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
    Occurance.Left = SysInfo1.WorkAreaLeft
    Occurance.Width = SysInfo1.WorkAreaWidth
    Occurance.Top = SysInfo1.WorkAreaTop
    Occurance.Height = SysInfo1.WorkAreaHeight
    Chart1.Width = SysInfo1.WorkAreaWidth - 2
    ProgressBar1.Width = SysInfo1.WorkAreaWidth - 2
    ProgressBar1.Value = 0
    ' left column of options
    Option1.Top = SysInfo1.WorkAreaHeight - 2420
    ProgressBar1.Top = Option1.Top - 270
    Chart1.Height = ProgressBar1.Top - 35
    Option2.Top = Option1.Top + 240
    Option5.Top = Option2.Top + 240
    Option7.Top = Option5.Top + 240
    ' right column
    Option3.Top = SysInfo1.WorkAreaHeight - 2420
    Option4.Top = Option3.Top + 240
    Option6.Top = Option4.Top + 240
    Option8.Top = Option6.Top + 240
    Option9.Top = Option8.Top + 240
    ' navigation buttons
    MENU1.Width = Occurance.Width / 5
    MENU2.Width = Occurance.Width / 5
    MENU3.Width = Occurance.Width / 5
    MENU4.Width = Occurance.Width / 5
    MENU5.Width = Occurance.Width / 5
    MENU2.Left = MENU1.Width
    MENU3.Left = MENU1.Width * 2
    MENU5.Left = MENU1.Width * 3
    MENU4.Left = MENU1.Width * 4
End Sub

' counts how often each ball 1-45 appeared, scanning draws from the newest back.
'   firstBall/lastBall select which of the 8 drawn balls count (1-6 main, 7-8 supps)
'   since45 = True stops at the draw where 45-ball games started
Private Sub CountBalls(ByVal firstBall As Integer, ByVal lastBall As Integer, ByVal since45 As Boolean, ByVal title As String)
    Dim counts(45) As Integer
    Dim d As Integer, n As Integer, j As Integer, stopAt As Integer
    Chart1.ShowLegend = False
    DrawIdx = LastDraw
    ProgressBar1.Min = 0
    stopAt = 1
    If since45 Then
        For d = 1 To LastDraw
            If Mid(DrawLine(d), 1, 5) = DRAW_45_INTRODUCED Then stopAt = d: ProgressBar1.Max = LastDraw - d
        Next d
    Else
        ProgressBar1.Max = LastDraw
    End If
    ProgressBar1.Value = 0
    For n = 1 To 45: counts(n) = 0: Next n
    For d = 1 To LastDraw
        Update
        For n = 1 To 45
            For j = firstBall To lastBall
                If n = Val(Drawn(j)) Then counts(n) = counts(n) + 1
            Next j
        Next n
        If since45 And Mid(DrawLine(DrawIdx), 1, 5) = DRAW_45_INTRODUCED Then Exit For
        DrawIdx = DrawIdx - 1
        If d <= ProgressBar1.Max Then ProgressBar1.Value = d
    Next d
    Chart1.TitleText = title
    For n = 1 To 45
        ChartData(n) = counts(n)
    Next n
    Chart1.ChartData = ChartData
End Sub

' "weeks since last drawn" for every ball
Private Sub WeeksSinceDrawn()
    Dim weeks(45) As Integer
    Dim n As Integer, j As Integer
    Chart1.ShowLegend = False
    DrawIdx = LastDraw
    ProgressBar1.Min = 0
    ProgressBar1.Max = 45
    ProgressBar1.Value = 0
    For n = 1 To 45: weeks(n) = -1: Next n
    Update
    For n = 1 To 45
        Do While weeks(n) = -1
            For j = 1 To 8
                If n = Val(Drawn(j)) Then
                    weeks(n) = LastDraw - DrawIdx
                    DrawIdx = LastDraw + 1
                End If
            Next j
            DrawIdx = DrawIdx - 1
            If DrawIdx < 1 Then weeks(n) = LastDraw: DrawIdx = LastDraw   ' guard (not in original)
            Update
        Loop
        ProgressBar1.Value = n
    Next n
    Chart1.TitleText = "NUMBER OF WEEKS SINCE DRAWN LAST"
    For n = 1 To 45
        ChartData(n) = weeks(n)
    Next n
    Chart1.ChartData = ChartData
End Sub

Private Sub Form_Load()
    Dim i As Integer
    DrawIdx = LastDraw
    LayoutForm
    For i = 1 To 45
        Chart1.Row = i
        Chart1.RowLabel = CStr(i)
    Next i
    WeeksSinceDrawn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Closedown
End Sub

Private Sub SysInfo1_DisplayChanged()
    LayoutForm
End Sub

'---------------------------------------------------------------------------
'  Graph selections
'---------------------------------------------------------------------------
Private Sub Option1_Click()
    CountBalls 1, 8, False, "NUMBER OF TIMES EACH BALL WAS DRAWN SINCE FIRST DRAW"
End Sub
Private Sub Option2_Click()
    CountBalls 1, 8, True, "NUMBER OF TIMES EACH BALL DRAWN SINCE 45 NUMBERS INTRODUCED"
End Sub
Private Sub Option3_Click()
    CountBalls 1, 6, False, "NUMBER OF TIMES EACH BALL DRAWN SINCE FIRST DRAW EXCLUDING SUPPLENTARY"
End Sub
Private Sub Option4_Click()
    CountBalls 1, 6, True, "NUMBER OF TIMES EACH BALL DRAWN SINCE 45 NUMBERS INTRODUCED EXCLUDING SUPPS"
End Sub
Private Sub Option5_Click()
    CountBalls 7, 8, False, "NUMBER OF TIMES EACH BALL DRAWN AS A SUPPLEMENTARY SINCE FIRST DRAW"
End Sub
Private Sub Option6_Click()
    CountBalls 7, 8, True, "NUMBER OF TIMES EACH BALL WAS DRAWN AS SUPPLEMENTARY SINCE 45 NUMBERS INTRODUCED"
End Sub
Private Sub Option7_Click()
    WeeksSinceDrawn
End Sub

Private Sub Option8_Click()
    ' sum of the six main balls for every draw since 45 numbers were introduced
    Dim sums(4000) As Integer
    Dim cd() As Variant
    Dim d As Integer, n As Integer
    Chart1.ShowLegend = False
    DrawIdx = LastDraw
    ProgressBar1.Min = 0
    ProgressBar1.Max = LastDraw - 659
    ProgressBar1.Value = 0
    ReDim cd(LastDraw - 659, 1)
    For n = 1 To 4000: sums(n) = 0: Next n
    For d = 1 To LastDraw - 659
        Update
        sums(d) = CInt(Val(Drawn(1)) + Val(Drawn(2)) + Val(Drawn(3)) + Val(Drawn(4)) + Val(Drawn(5)) + Val(Drawn(6)))
        DrawIdx = DrawIdx - 1
        ProgressBar1.Value = d
    Next d
    Chart1.TitleText = "SUM OF DRAWS EXCLUDING SUPPLEMENTARY NUMBERS SINCE 45 INTRODUCED"
    For d = 1 To LastDraw - 659
        cd(d - 1, 1) = sums(d)
    Next d
    Chart1.ChartType = 3                      ' 2D line
    Chart1.ChartData = cd
    Chart1.RowCount = (LastDraw - 659) - 1
    Chart1.ChartType = 3
    With Chart1.Plot.SeriesCollection(1).StatLine
        .VtColor.Set 12, 12, 255
        .Flag = 1 Or 4 Or 16                  ' minimum, mean, regression
        .Style(1) = 3
        .Style(4) = 1
        .Style(16) = 6
        .Width = 2
    End With
End Sub

Private Sub Option9_Click()
    ' division wins per draw for the game loaded on the CHECK WINNINGS screen
    If DrawsCompared = 0 Then Exit Sub
    ProgressBar1.Value = ProgressBar1.Min
    Chart1.TitleText = "WEEKS SINCE WIN GRAPH FOR CURRENT GAME OVER " & CStr(DrawsCompared) & " WEEKS"
    Chart1.ChartType = 1
    Chart1.ChartData = WinHistory
    Chart1.RowCount = DrawsCompared
    Chart1.ColumnCount = 5
    Chart1.ShowLegend = True
    Chart1.Column = 1: Chart1.ColumnLabel = "Division 1"
    Chart1.Column = 2: Chart1.ColumnLabel = "Division 2"
    Chart1.Column = 3: Chart1.ColumnLabel = "Division 3"
    Chart1.Column = 4: Chart1.ColumnLabel = "Division 4"
    Chart1.Column = 5: Chart1.ColumnLabel = "Division 5"
End Sub

'---------------------------------------------------------------------------
'  Graph type menu  (MSChart ChartType values)
'---------------------------------------------------------------------------
Private Sub threed_Click():  Chart1.ChartType = 0: End Sub   ' 3D bar
Private Sub twod_Click():    Chart1.ChartType = 1: End Sub   ' 2D bar
Private Sub Line3d_Click()
    Chart1.ChartType = 2
End Sub
Private Sub line2d_Click()
    Chart1.ChartType = 3
End Sub
Private Sub Area3d_Click()
    Chart1.ChartType = 4
End Sub
Private Sub Area_Click()
    Chart1.ChartType = 5
End Sub
Private Sub Step3d_Click()
    Chart1.ChartType = 6
End Sub
Private Sub Step2d_Click()
    Chart1.ChartType = 7
End Sub
Private Sub autocycle_Click()
    If autocycle.Checked = True Then
        autocycle.Checked = False
        Timer1.Enabled = False
    Else
        autocycle.Checked = True
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    ' steps through the 2D (1,3,5,7) or 3D (0,2,4,6) chart styles
    Select Case Chart1.ChartType
        Case 1: Chart1.ChartType = 3
        Case 3: Chart1.ChartType = 5
        Case 5: Chart1.ChartType = 7
        Case 7: Chart1.ChartType = 1
        Case 0: Chart1.ChartType = 2
        Case 2: Chart1.ChartType = 4
        Case 4: Chart1.ChartType = 6
        Case 6: Chart1.ChartType = 0
    End Select
End Sub

'---------------------------------------------------------------------------
'  Navigation
'---------------------------------------------------------------------------
Private Sub MENU1_Click()
    adddraw.Show
    Occurance.Hide
End Sub
Private Sub MENU2_Click()
    Lotto.Show
    Occurance.Hide
End Sub
Private Sub MENU3_Click()
    Board.Show
    Occurance.Hide
End Sub
Private Sub MENU4_Click()
    Wheelsystem.Show
    Occurance.Hide
End Sub
Private Sub Quit_Click()
    Lotto.Show
    Occurance.Hide
End Sub
