VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "A Dumb Fly"
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
'This program is a Visual Basic version of listing 13.3 of
'"Tricks of the Game-Programming Gurus" by Andre LeMothe, copyright 1994.
'Listing 13.3, 'A Dumb Fly.'
Option Explicit
Private ex As Long, ey As Long
Private curr_xv As Long, curr_yv As Long
Private clicks As Long
Private Done As Long
Private GetCh As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub MainLoop()
    Dim NextTime As Long
    Do
        'Check for quit
        If Len(GetCh) > 0 Then
            Done = 1
        End If
        If Done = 1 Then
            Exit Do
        End If
        
        'Erase dots
        Form1.ForeColor = vbBlack
        Form1.PSet (ex, ey)
        
        'Begin brain
        'Are we done with this direction?
        
        clicks = clicks + 1
        If clicks >= 20 Then
            curr_xv = -5 + Random(10)
            curr_yv = -5 + Random(10)
            clicks = 0
        End If
        
        'Move the fly
        ex = ex + curr_xv
        ey = ey + curr_yv
        
        'Make sure the fly stays on the screen.
        If ex > Form1.ScaleWidth Then ex = 0
        If ex < 0 Then ex = Form1.ScaleWidth
        If ey > Form1.ScaleHeight Then ey = 0
        If ey < 0 Then ey = Form1.ScaleHeight
        
        'draw fly
        Form1.ForeColor = vbRed
        Form1.PSet (ex, ey)
        
        DoEvents
        'Wait until the next timer
        Do While timeGetTime < NextTime
            DoEvents
        Loop
        NextTime = timeGetTime + 50
        
    Loop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'There's no 'getch' function in VB, but this will do.
'Make sure to erase it after processing it.
    GetCh = Chr$(KeyAscii)
End Sub

Private Sub Form_Load()
    ex = 160: ey = 100
    curr_xv = 1: curr_yv = 0
    Done = 0
    clicks = 0
    Form1.ScaleMode = vbPixels
    Form1.Caption = "The Dumb Fly - Any Key To Quit"
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

Private Function Random(ByVal Max As Long) As Long
'Returns a number from 1 to max
    Random = Int((Max * Rnd) + 1)
End Function
