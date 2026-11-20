Attribute VB_Name = "Module1"
Option Explicit
' ---------------------------------------------------------------------------
'  Lottoluck '98 - shared data
'  Array sizes come from the bounds checks compiled into Xlotto.exe.
' ---------------------------------------------------------------------------

' Tattslotto history (tattslot.txt, one line per draw:  B1709,19971206,13,29,35,5,45,2,11,16)
Public DrawLine(5000) As String          ' was 1500 in 1997 - enlarged for the full history
Public DrawIdx As Integer               ' draw currently loaded into Drawn()
Public LastDraw As Integer              ' index of the most recent draw
Public Drawn(8) As String               ' balls of the current draw: 1-6 main, 7-8 supplementaries

' games being checked (loaded from a .lot file - 12 digits per line)
Public Const MAX_GAMES = 1250
Public Games(MAX_GAMES) As Variant
Public GameCount As Long

' CHECK WINNINGS screen
Public DrawsToCompare As Long           ' (number of draws to compare) + 1
Public DrawsCompared As Integer         ' draws used by the last Compareit (for the Game wins graph)
Public WinHistory(4999, 4) As Variant   ' division won per draw, used by the graphing screen
Public GamesWanted As Integer           ' Game Designer: games to create (also set to 1 at start-up)

' GRAPHING screen
Public ChartData(1 To 45) As Variant

' GAME DESIGNER screen
Public Balls(1 To 6) As String          ' the six balls of the game being generated
Public SaveString As String             ' games ready to be written to a .lot file (vbLf & vbCr separated)

' WHEELING SYSTEMS screen
Public WheelGame(1000) As String        ' wheel pattern, 12 digits = six positions into Chosen()
Public WheelGames As Integer            ' number of games in the wheel
Public WheelNumbers As Integer          ' field size (numbers to select)
Public WheelText As String              ' description of the wheel's guarantee
Public Wheeled(1000) As String          ' the expanded games
Public Chosen(45) As String             ' selected numbers
Public ClickedNum As Integer            ' last number clicked/picked on the grid
Public ClickedNum2 As Integer
Public SelectedCount As Integer

' ADD DRAW screen
Public AddDrawOpen As Integer
Public LastDrawNumber As Integer        ' number of the latest draw (from "B1709")

' tattslot.txt line 1 is a header, so the first draw is at index 1.  Draw B0413 (the
' first with 45 balls) is the 660th draw, so 659 draws came before it.
Public Const DRAWS_BEFORE_45 = 659

' Reads tattslot.txt into DrawLine() and returns the index of the newest draw.
' (FIX: the reconstructed loaders left LastDraw one short, so the newest draw was ignored
'  everywhere and ADD DRAW offered a draw number that already existed.)
Public Function LoadHistory() As Integer
    Dim f As Integer, i As Integer, l As String
    f = FreeFile
    Open App.Path & "\tattslot.txt" For Input As #f
    i = 0
    Do While Not EOF(f) And i <= UBound(DrawLine)
        Line Input #f, l
        If Len(Trim$(l)) > 0 Then
            DrawLine(i) = l
            i = i + 1
        End If
    Loop
    Close #f
    LoadHistory = i - 1
End Function

' Loads a .lot game file into Games(1..), skipping blank or short lines and stopping at
' MAX_GAMES (FIX: a file with more than 1250 games used to crash CHECK WINNINGS).
' Accepts the program's own LF+CR line breaks as well as normal CRLF.
Public Function LoadGameFile(ByVal path As String) As Long
    Dim f As Integer, l As String, n As Long
    f = FreeFile
    Open path For Input As #f
    n = 0
    Do While Not EOF(f) And n < MAX_GAMES
        Line Input #f, l
        l = Replace(Replace(l, vbLf, ""), " ", "")
        If Len(l) >= 12 Then
            n = n + 1
            Games(n) = Left$(l, 12)
        End If
    Loop
    Close #f
    LoadGameFile = n
End Function

' Two-digit text for a ball number ("7" -> "07")
Public Function Two(ByVal v As Variant) As String
    Two = Format$(Val(v), "00")
End Function
