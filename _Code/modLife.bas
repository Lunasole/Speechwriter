Attribute VB_Name = "modL"
Option Explicit
Const LifeX As Long = 30
Const LifeY As Long = 120
Public Const LifeMin = (LifeX * LifeY) / 20&
Public Const LifeMax = LifeMin * 2

Dim Life(LifeX, LifeY) As Byte

Dim x As Long
Dim y As Long

Public AliveCnt As Long

Private Sub cell_Kill()
    Life(x, y) = 0
    DrawCell
    AliveCnt = AliveCnt - 1
End Sub
Private Sub cell_Raise()
    Life(x, y) = 1
    DrawCell
    AliveCnt = AliveCnt + 1
End Sub
Public Sub LifeInsert1()
    x = RNDINT(LifeX)
    y = RNDINT(LifeY)
    If Not Life(x, y) = 1 Then
        cell_Raise
    End If
End Sub
Public Sub LifeLoop()
    Dim CAround As Byte
    
    Dim a As Long, b As Long
    Dim C As Long, D As Long
    Dim S1 As Long, S2 As Long
    
    If Rnd <= 0.5 Then
        a = 0
        b = LifeX
        S1 = 1
    Else
        a = LifeX
        b = 0
        S1 = -1
    End If
    If Rnd <= 0.5 Then
        C = 0
        D = LifeY
        S2 = 1
    Else
        C = LifeY
        D = 0
        S2 = -1
    End If
    
    For x = a To b Step S1
    For y = C To D Step S2
        CAround = MooreNeigh()

        If Not Life(x, y) = 1 Then
            If CAround = 3 Then
                'spawn new cell
                cell_Raise
            End If
        Else
            If CAround = 2 Or CAround = 3 Then
                If Rnd <= 0.5 And AliveCnt > LifeMax Then
                    'custom mod: die
                    cell_Kill
                Else
                    'continue life
                End If
            Else
                'die
                If AliveCnt > LifeMin Then
                    cell_Kill
                End If
            End If
        End If
    Next y
    Next x
End Sub
Private Function MooreNeigh() As Byte
Dim tX As Long, tY As Long
    tX = x - 1: tY = y - 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY And Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x: tY = y - 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x + 1: tY = y - 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY And Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x - 1: tY = y ': GoSub fixP
        If Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x + 1: tY = y ': GoSub fixP
        If Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x - 1: tY = y + 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY And Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x: tY = y + 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
    tX = x + 1: tY = y + 1 ': GoSub fixP
        If Not tY < 0 And Not tY > LifeY And Not tX < 0 And Not tX > LifeX Then
        MooreNeigh = MooreNeigh + Life(tX, tY)
        End If
Exit Function
fixP:
    If tX < 0 Then
        tX = LifeX
    ElseIf tX > LifeX Then
        tX = 0
    End If
    If tY < 0 Then
        tY = LifeY
    ElseIf tY > LifeY Then
        tY = 0
    End If
        Return
End Function
Private Sub DrawCell()
    If Life(x, y) = 1 Then
        frmMain.imgRandom.PSet (x, y) ', Clr
    Else
        frmMain.imgRandom.PSet (x, y), &H80000005   '&H8000000F
    End If
End Sub

Public Sub LifeSub()
    If AliveCnt <= LifeMin Then
        LifeInsert1
        LifeInsert1
        LifeInsert1
    Else
        LifeLoop
'        If AliveCnt <= LifeMin Then
'            LifeInsert1
'        End If
    End If
End Sub
