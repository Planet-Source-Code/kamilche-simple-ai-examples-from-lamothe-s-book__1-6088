VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is a Visual Basic version of listing 13.1 of
'"Tricks of the Game-Programming Gurus" by Andre LeMothe, copyright 1994.
'It implements a simple 'chase' algorithm.
Option Explicit
Private px As Long, py As Long, ex As Long, ey As Long
Private Done As Long
Private GetCh As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub MainLoop()
    Dim NextTime As Long
    Do
        If Done = 1 Then
            Exit Do
        End If
        'Erase dots
        Form1.ForeColor = vbBlack
        Form1.PSet (px, py)
        Form1.PSet (ex, ey)
        'Move player
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
            'make it wrap around - my addition.
            If px >= Form1.ScaleWidth Then px = 0
            If py >= Form1.ScaleHeight Then py = 0
            If px < 0 Then px = Form1.ScaleWidth - 1
            If py < 0 Then py = Form1.ScaleHeight - 1
        End If
        
        'move enemy
        
        'begin brain
        If (px > ex) Then ex = ex + 1
        If (px < ex) Then ex = ex - 1
        If (py > ey) Then ey = ey + 1
        If (py < ey) Then ey = ey - 1
        'end brain
        
        'draw dots
        Form1.ForeColor = vbGreen
        Form1.PSet (px, py)
        Form1.ForeColor = vbRed
        Form1.PSet (ex, ey)
        
        'Wait until the next timer
        Do While timeGetTime < NextTime
            DoEvents
        Loop
        NextTime = timeGetTime + 100
        
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
    px = 0: py = 0: ex = 0: ey = 0
    Done = 0
    Form1.ScaleMode = vbPixels
    Form1.Caption = "The Terminator - Q to Quit"
    Form1.BackColor = vbBlack
    Form1.DrawWidth = 3
    Form1.Show
    MainLoop
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Also let a form close quit the program. My addition.
    GetCh = "q"
End Sub
