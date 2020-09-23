VERSION 5.00
Begin VB.Form frmFly 
   Caption         =   "The Fly"
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
'This program is a Visual Basic version of listing 13.2 of
'"Tricks of the Game-Programming Gurus" by Andre LeMothe, copyright 1994.
'Listing 13.2, 'The Fly.'
Option Explicit
Private px As Long, py As Long, ex As Long, ey As Long
Private doing_pattern As Long
Private current_pattern As Long
Private pattern_element As Long
Private Done As Long
Private GetCh As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private patterns_x(1 To 3, 1 To 20) As Long
Private patterns_y(1 To 3, 1 To 20) As Long

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
        If doing_pattern = 0 Then
            If px > ex Then ex = ex + 1
            If px < ex Then ex = ex - 1
            If py > ey Then ey = ey + 1
            If py < ey Then ey = ey - 1
            
            'Check whether it's time to do a pattern
            '(that is, is enemy within 50 pixels of player?)
            If Sqr(0.1 + (px - ex) * (px - ex) + (py - ey) * (py - ey)) < 50 Then
                'never use SQR in a real game!
                'get a new random pattern
                current_pattern = Int((3 * Rnd) + 1)
                'set brain into pattern state
                doing_pattern = 1
                pattern_element = 1
            End If
        Else
            'Move the enemy using the next pattern element of the current pattern
            ex = ex + patterns_x(current_pattern, pattern_element)
            ey = ey + patterns_y(current_pattern, pattern_element)
            'Are we done doing pattern?
            pattern_element = pattern_element + 1
            If pattern_element > 20 Then
                pattern_element = 1
                doing_pattern = 0
            End If
        End If
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
        NextTime = timeGetTime + 15
        
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
    frmFly.ScaleMode = vbPixels
    frmFly.Caption = "The Fly - Q to Quit"
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

