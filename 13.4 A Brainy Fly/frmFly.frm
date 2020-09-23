VERSION 5.00
Begin VB.Form frmFly 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is a Visual Basic version of listing 13.4 of
'"Tricks of the Game-Programming Gurus" by Andre LeMothe, copyright 1994.
'Listing 13.4, 'Brainy Fly.'
Option Explicit
Private Const STATE_CHASE = 1
Private Const STATE_RANDOM = 2
Private Const STATE_EVADE = 3
Private Const STATE_PATTERN = 4

Private px As Long, py As Long, ex As Long, ey As Long
Private curr_xv As Long, curr_yv As Long
Private doing_pattern As Long
Private current_pattern As Long
Private pattern_element As Long
Private select_state As Long
Private Done As Long
Private GetCh As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private patterns_x(1 To 3, 1 To 20) As Long
Private patterns_y(1 To 3, 1 To 20) As Long
Private clicks As Long
Private fly_state As Long
Private distance As Double

Private Sub MainLoop()
    Dim NextTime As Long
    Do
        'Check for quit
        If Done = 1 Then
            Exit Do
        End If
        
        'Erase dots
        frmFly.ForeColor = vbBlack
        frmFly.PSet (px, py)
        frmFly.PSet (ex, ey)
        
        'Move player
        'Which way is player moving?
        If Len(GetCh) > 0 Then
            If GetCh = "u" Then
                py = py - 2
            ElseIf GetCh = "n" Then
                py = py + 2
            ElseIf GetCh = "j" Then
                px = px + 2
            ElseIf GetCh = "h" Then
                px = px - 2
            ElseIf GetCh = "q" Then
                Done = 1
            End If
            GetCh = ""
        End If
        
        'Move enemy
        'Begin brain
        'What state is brain in? Let FSM sort it out.
        frmFly.Cls
        frmFly.ForeColor = vbWhite
        frmFly.CurrentX = 0: frmFly.CurrentY = frmFly.ScaleHeight - 20
        Select Case fly_state
            Case STATE_CHASE:
                frmFly.Print "current state: chase. clicks: " & clicks
                'make the fly chase the player
                If px > ex Then ex = ex + 1
                If px < ex Then ex = ex - 1
                If py > ey Then ey = ey + 1
                If py < ey Then ey = ey - 1
                'time to go to another state
                clicks = clicks - 1: If clicks <= 0 Then select_state = 1
            Case STATE_RANDOM:
                frmFly.Print "current state: random. clicks: " & clicks
                'move fly in random direction
                ex = ex + curr_xv
                ey = ey + curr_yv
                'time to go to another state
                clicks = clicks - 1: If clicks <= 0 Then select_state = 1
            Case STATE_EVADE:
                frmFly.Print "current state: evade. clicks: " & clicks
                'make the fly run from the player
                If px > ex Then ex = ex - 1
                If px < ex Then ex = ex + 1
                If py > ey Then ey = ey - 1
                If py < ey Then ey = ey + 1
                'time to go to another state
                clicks = clicks - 1: If clicks <= 0 Then select_state = 1
            Case STATE_PATTERN:
                frmFly.Print "current state: pattern. clicks: " & clicks
                'Move the enemy using the next pattern element of the current pattern
                ex = ex + patterns_x(current_pattern, pattern_element)
                ey = ey + patterns_y(current_pattern, pattern_element)
                'Are we done doing pattern?
                pattern_element = pattern_element + 1
                If pattern_element >= 20 Then
                    pattern_element = 1
                    select_state = 1
                End If
        End Select
        
        'Does brain want another state?
        If select_state = 1 Then
            'Select a state based on the environment and on fuzzy logic.
            'Uses distance from player to select a new state.
            distance = Sqr(0.5 + Abs((px - ex) * (px - ex) + (py - ey) * (py - ey)))
            If distance > 5 And distance < 15 And Random(2) = 1 Then
                'get a new random pattern
                current_pattern = Random(3)
                'set brain into pattern state
                fly_state = STATE_PATTERN
                pattern_element = 1
            ElseIf distance < 10 Then 'too close let's run!
                clicks = 20
                fly_state = STATE_EVADE
            ElseIf distance > 25 And distance < 100 And Random(3) = 1 Then
                'let's chase the player
                clicks = 15
                fly_state = STATE_CHASE
            ElseIf distance > 30 And Random(2) = 1 Then
                'random
                clicks = 10
                fly_state = STATE_RANDOM
                curr_xv = -5 + Random(10)
                curr_yv = -5 + Random(10)
            Else
                'random
                clicks = 5
                fly_state = STATE_RANDOM
                curr_xv = -5 + Random(10)
                curr_yv = -5 + Random(10)
            End If
            select_state = 0
            
        End If
        
        'Make sure fly stays on screen
        If ex > frmFly.ScaleWidth Then ex = 0
        If ex < 0 Then ex = frmFly.ScaleWidth
        If ey > frmFly.ScaleHeight Then ey = 0
        If ey < 0 Then ey = frmFly.ScaleHeight
        
        'end brain
        
        'draw dots
        frmFly.ForeColor = vbGreen
        frmFly.PSet (px, py)
        frmFly.ForeColor = vbRed
        frmFly.PSet (ex, ey)
        
        DoEvents
        'Wait until the next timer
        Do While timeGetTime < NextTime
            DoEvents
        Loop
        NextTime = timeGetTime + 50
        
    Loop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Also let the player use the arrow keys. My addition.
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        'translate to one of the ugly key assignments
        If KeyCode = vbKeyUp Then
            Form_KeyPress Asc("u")
        ElseIf KeyCode = vbKeyDown Then
            Form_KeyPress Asc("n")
        ElseIf KeyCode = vbKeyLeft Then
            Form_KeyPress Asc("h")
        ElseIf KeyCode = vbKeyRight Then
            Form_KeyPress Asc("j")
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'There's no 'getch' function in VB, but this will do.
'Make sure to erase it after processing it.
    GetCh = Chr$(KeyAscii)
End Sub

Private Sub Form_Load()
    px = 160: py = 100: ex = 0: ey = 0
    Done = 0
    select_state = 0
    clicks = 20
    fly_state = STATE_CHASE
    frmFly.ScaleMode = vbPixels
    frmFly.Caption = "Brainy Fly - Q to Quit"
    frmFly.BackColor = vbBlack
    frmFly.DrawWidth = 3
    LoadArray
    frmFly.Show
    MainLoop
    End
End Sub

Private Sub LoadArray()
'Loads the array with data
    Dim FileNo As Integer, FileName As String
    Dim i As Long, j As Long
    FileName = App.Path & "\Data.txt"
    FileNo = FreeFile
    Open FileName For Input As #FileNo
    For i = 1 To 3
        For j = 1 To 20
            Input #FileNo, patterns_x(i, j)
        Next j
    Next i
    For i = 1 To 3
        For j = 1 To 20
            Input #FileNo, patterns_y(i, j)
        Next j
    Next i
    Close #FileNo
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'erase old dots
    frmFly.ForeColor = vbBlack
    frmFly.PSet (px, py)
    px = X: py = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Also let a form close quit the program. My addition.
    GetCh = "q"
End Sub

Private Function Random(ByVal Max As Long) As Long
'Returns a number from 1 to max
    Random = Int((Max * Rnd) + 1)
End Function

