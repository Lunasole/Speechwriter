VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CRijndael
' FILENAME:     CRijndael.cls
' AUTHOR:       Phil Fresle
' CREATED:      16-Feb-2001
' COPYRIGHT:    Copyright 2001 Phil Fresle
' EMAIL:        phil@frez.co.uk
' WEB:          http://www.frez.co.uk
'
' DESCRIPTION:
' Implementation of the AES Rijndael Block Cipher. Inspired by Mike Scott's
' implementation in C. Permission for free direct or derivative use is granted
' subject to compliance with any conditions that the originators of the
' algorithm place on its exploitation.
'
' MODIFICATION HISTORY:
' 16-Feb-2001   Phil Fresle     Initial Version
' 03-Apr-2001   Phil Fresle     Added EncryptData and DecryptData functions to
'                               make it easier to use by VB developers for
'                               encrypting and decrypting strings. These procs
'                               take large byte arrays, the resultant encoded
'                               data includes the message length inserted on
'                               the front four bytes prior to encryption.
' 19-Apr-2001   Phil Fresle     Thanks to Paolo Migliaccio for finding a bug
'                               with 256 bit key. Problem was in the gkey
'                               function. Now properly matches NIST values.
'*******************************************************************************
Option Explicit

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private m_bytOnBits(7)  As Byte
Private m_byt2Power(7)  As Byte

Private m_InCo(3) As Byte

Private m_fbsub(255)    As Byte
Private m_rbsub(255)    As Byte
Private m_ptab(255)     As Byte
Private m_ltab(255)     As Byte
Private m_ftable(255)   As Long
Private m_rtable(255)   As Long
Private m_rco(29)       As Long

Private m_Nk        As Long
Private m_Nb        As Long
Private m_Nr        As Long
Private m_fi(23)    As Byte
Private m_ri(23)    As Byte
Private m_fkey(119) As Long
Private m_rkey(119) As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'    (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Class_Initialize()
    m_InCo(0) = &HB
    m_InCo(1) = &HD
    m_InCo(2) = &H9
    m_InCo(3) = &HE
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_bytOnBits(0) = 1          ' 00000001
    m_bytOnBits(1) = 3          ' 00000011
    m_bytOnBits(2) = 7          ' 00000111
    m_bytOnBits(3) = 15         ' 00001111
    m_bytOnBits(4) = 31         ' 00011111
    m_bytOnBits(5) = 63         ' 00111111
    m_bytOnBits(6) = 127        ' 01111111
    m_bytOnBits(7) = 255        ' 11111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_byt2Power(0) = 1          ' 00000001
    m_byt2Power(1) = 2          ' 00000010
    m_byt2Power(2) = 4          ' 00000100
    m_byt2Power(3) = 8          ' 00001000
    m_byt2Power(4) = 16         ' 00010000
    m_byt2Power(5) = 32         ' 00100000
    m_byt2Power(6) = 64         ' 01000000
    m_byt2Power(7) = 128        ' 10000000
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
End Sub

'*******************************************************************************
' LShift (FUNCTION)
'*******************************************************************************
Private Function LShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * _
            m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * _
            m_l2Power(iShiftBits))
    End If
End Function

'*******************************************************************************
' RShift (FUNCTION)
'*******************************************************************************
Private Function RShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

'*******************************************************************************
' LShiftByte (FUNCTION)
'*******************************************************************************
Private Function LShiftByte(ByVal bytValue As Byte, _
                            ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        LShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And 1 Then
            LShiftByte = &H80
        Else
            LShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * _
        m_byt2Power(bytShiftBits))
End Function

'*******************************************************************************
' RShiftByte (FUNCTION)
'*******************************************************************************
Private Function RShiftByte(ByVal bytValue As Byte, _
                            ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        RShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And &H80 Then
            RShiftByte = 1
        Else
            RShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
End Function

'*******************************************************************************
' RotateLeft (FUNCTION)
'*******************************************************************************
Private Function RotateLeft(ByVal lValue As Long, _
                            ByVal iShiftBits As Integer) As Long
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

''*******************************************************************************
'' RotateLeftByte (FUNCTION)
'*******************************************************************************
Private Function RotateLeftByte(ByVal bytValue As Byte, _
                                ByVal bytShiftBits As Byte) As Byte
    RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or _
        RShiftByte(bytValue, (8 - bytShiftBits))
End Function

'*******************************************************************************
' Pack (FUNCTION)
'*******************************************************************************
Private Function Pack(b() As Byte) As Long
    Dim lCount As Long
    Dim lTemp  As Long
    
    For lCount = 0 To 3
        lTemp = b(lCount)
        Pack = Pack Or LShift(lTemp, (lCount * 8))
    Next
End Function

'*******************************************************************************
' PackFrom (FUNCTION)
'*******************************************************************************
Private Function PackFrom(b() As Byte, _
                          ByVal K As Long) As Long
    Dim lCount As Long
    Dim lTemp  As Long
    
    For lCount = 0 To 3
        lTemp = b(lCount + K)
        PackFrom = PackFrom Or LShift(lTemp, (lCount * 8))
    Next
End Function

'*******************************************************************************
' Unpack (SUB)
'*******************************************************************************
Private Sub Unpack(ByVal a As Long, _
                   b() As Byte)
    b(0) = a And m_lOnBits(7)
    b(1) = RShift(a, 8) And m_lOnBits(7)
    b(2) = RShift(a, 16) And m_lOnBits(7)
    b(3) = RShift(a, 24) And m_lOnBits(7)
End Sub

'*******************************************************************************
' UnpackFrom (SUB)
'*******************************************************************************
Private Sub UnpackFrom(ByVal a As Long, _
                       b() As Byte, _
                       ByVal K As Long)
    b(0 + K) = a And m_lOnBits(7)
    b(1 + K) = RShift(a, 8) And m_lOnBits(7)
    b(2 + K) = RShift(a, 16) And m_lOnBits(7)
    b(3 + K) = RShift(a, 24) And m_lOnBits(7)
End Sub

'*******************************************************************************
' xtime (FUNCTION)
'*******************************************************************************
Private Function xtime(ByVal a As Byte) As Byte
    Dim b As Byte
    
    If (a And &H80) Then
        b = &H1B
    Else
        b = 0
    End If
    
    a = LShiftByte(a, 1)
    a = a Xor b
    
    xtime = a
End Function

'*******************************************************************************
' bmul (FUNCTION)
'*******************************************************************************
Private Function bmul(ByVal X As Byte, _
                      Y As Byte) As Byte
    If X <> 0 And Y <> 0 Then
        bmul = m_ptab((CLng(m_ltab(X)) + CLng(m_ltab(Y))) Mod 255)
    Else
        bmul = 0
    End If
End Function

'*******************************************************************************
' SubByte (FUNCTION)
'*******************************************************************************
Private Function SubByte(ByVal a As Long) As Long
    Dim b(3) As Byte
    
    Unpack a, b
    b(0) = m_fbsub(b(0))
    b(1) = m_fbsub(b(1))
    b(2) = m_fbsub(b(2))
    b(3) = m_fbsub(b(3))
    
    SubByte = Pack(b)
End Function

'*******************************************************************************
' product (FUNCTION)
'*******************************************************************************
Private Function product(ByVal X As Long, _
                         ByVal Y As Long) As Long
    Dim xb(3) As Byte
    Dim yb(3) As Byte
    
    Unpack X, xb
    Unpack Y, yb
    product = bmul(xb(0), yb(0)) Xor bmul(xb(1), yb(1)) Xor bmul(xb(2), yb(2)) _
        Xor bmul(xb(3), yb(3))
End Function

'*******************************************************************************
' InvMixCol (FUNCTION)
'*******************************************************************************
Private Function InvMixCol(ByVal X As Long) As Long
    Dim Y       As Long
    Dim m       As Long
    Dim b(3)    As Byte
    
    m = Pack(m_InCo)
    b(3) = product(m, X)
    m = RotateLeft(m, 24)
    b(2) = product(m, X)
    m = RotateLeft(m, 24)
    b(1) = product(m, X)
    m = RotateLeft(m, 24)
    b(0) = product(m, X)
    Y = Pack(b)
    
    InvMixCol = Y
End Function

'*******************************************************************************
' ByteSub (FUNCTION)
'*******************************************************************************
Private Function ByteSub(ByVal X As Byte) As Byte
    Dim Y As Byte
    
    Y = m_ptab(255 - m_ltab(X))
    X = Y
    X = RotateLeftByte(X, 1)
    Y = Y Xor X
    X = RotateLeftByte(X, 1)
    Y = Y Xor X
    X = RotateLeftByte(X, 1)
    Y = Y Xor X
    X = RotateLeftByte(X, 1)
    Y = Y Xor X
    Y = Y Xor &H63
    
    ByteSub = Y
End Function

'*******************************************************************************
' gentables (SUB)
'*******************************************************************************
Public Sub gentables()
    Dim i       As Long
    Dim Y       As Byte
    Dim b(3)    As Byte
    Dim ib      As Byte
    
    m_ltab(0) = 0
    m_ptab(0) = 1
    m_ltab(1) = 0
    m_ptab(1) = 3
    m_ltab(3) = 1
    
    For i = 2 To 255
        m_ptab(i) = m_ptab(i - 1) Xor xtime(m_ptab(i - 1))
        m_ltab(m_ptab(i)) = i
    Next
    
    m_fbsub(0) = &H63
    m_rbsub(&H63) = 0
    
    For i = 1 To 255
        ib = i
        Y = ByteSub(ib)
        m_fbsub(i) = Y
        m_rbsub(Y) = i
    Next
    
        Y = 1
    For i = 0 To 29
        m_rco(i) = Y
        Y = xtime(Y)
    Next
    
    For i = 0 To 255
        Y = m_fbsub(i)
        b(3) = Y Xor xtime(Y)
        b(2) = Y
        b(1) = Y
        b(0) = xtime(Y)
        m_ftable(i) = Pack(b)
        
        Y = m_rbsub(i)
        b(3) = bmul(m_InCo(0), Y)
        b(2) = bmul(m_InCo(1), Y)
        b(1) = bmul(m_InCo(2), Y)
        b(0) = bmul(m_InCo(3), Y)
        m_rtable(i) = Pack(b)
    Next
End Sub

'*******************************************************************************
' gkey (SUB)
'*******************************************************************************
Public Sub gkey(ByVal nb As Long, _
                ByVal nk As Long, _
                Key() As Byte)
                
    Dim i               As Long
    Dim j               As Long
    Dim K               As Long
    Dim m               As Long
    Dim n               As Long
    Dim C1              As Long
    Dim C2              As Long
    Dim C3              As Long
    Dim CipherKey(7)    As Long
    
    m_Nb = nb
    m_Nk = nk
    
    If m_Nb >= m_Nk Then
        m_Nr = 6 + m_Nb
    Else
        m_Nr = 6 + m_Nk
    End If
    
    C1 = 1
    If m_Nb < 8 Then
        C2 = 2
        C3 = 3
    Else
        C2 = 3
        C3 = 4
    End If
    
    For j = 0 To nb - 1
        m = j * 3
        
        m_fi(m) = (j + C1) Mod nb
        m_fi(m + 1) = (j + C2) Mod nb
        m_fi(m + 2) = (j + C3) Mod nb
        m_ri(m) = (nb + j - C1) Mod nb
        m_ri(m + 1) = (nb + j - C2) Mod nb
        m_ri(m + 2) = (nb + j - C3) Mod nb
    Next
    
    n = m_Nb * (m_Nr + 1)
    
    For i = 0 To m_Nk - 1
        j = i * 4
        CipherKey(i) = PackFrom(Key, j)
    Next
    
    For i = 0 To m_Nk - 1
        m_fkey(i) = CipherKey(i)
    Next
    
    j = m_Nk
    K = 0
    Do While j < n
        m_fkey(j) = m_fkey(j - m_Nk) Xor _
            SubByte(RotateLeft(m_fkey(j - 1), 24)) Xor m_rco(K)
        If m_Nk <= 6 Then
            i = 1
            Do While i < m_Nk And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
        Else
            ' Problem fixed here
            i = 1
            Do While i < 4 And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
            If j + 4 < n Then
                m_fkey(j + 4) = m_fkey(j + 4 - m_Nk) Xor _
                    SubByte(m_fkey(j + 3))
            End If
            i = 5
            Do While i < m_Nk And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
        End If
        
        j = j + m_Nk
        K = K + 1
    Loop
    
    For j = 0 To m_Nb - 1
        m_rkey(j + n - nb) = m_fkey(j)
    Next
    
    i = m_Nb
    Do While i < n - m_Nb
        K = n - m_Nb - i
        For j = 0 To m_Nb - 1
            m_rkey(K + j) = InvMixCol(m_fkey(i + j))
        Next
        i = i + m_Nb
    Loop
    
    j = n - m_Nb
    Do While j < n
        m_rkey(j - n + m_Nb) = m_fkey(j)
        j = j + 1
    Loop
End Sub

'*******************************************************************************
' encrypt (SUB)
'*******************************************************************************
#If Enc Then
Public Sub Encrypt(Buff() As Byte)
    Dim i       As Long
    Dim j       As Long
    Dim K       As Long
    Dim m       As Long
    Dim a(7)    As Long
    Dim b(7)    As Long
    Dim X()     As Long
    Dim Y()     As Long
    Dim t()     As Long
    
    For i = 0 To m_Nb - 1
        j = i * 4
        
        a(i) = PackFrom(Buff, j)
        a(i) = a(i) Xor m_fkey(i)
    Next
    
    K = m_Nb
    X = a
    Y = b
    
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            m = j * 3
            Y(j) = m_fkey(K) Xor m_ftable(X(j) And m_lOnBits(7)) Xor _
                RotateLeft(m_ftable(RShift(X(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
                RotateLeft(m_ftable(RShift(X(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
                RotateLeft(m_ftable(RShift(X(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
            K = K + 1
        Next
        t = X
        X = Y
        Y = t
    Next
    
    For j = 0 To m_Nb - 1
        m = j * 3
        Y(j) = m_fkey(K) Xor m_fbsub(X(j) And m_lOnBits(7)) Xor _
            RotateLeft(m_fbsub(RShift(X(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
            RotateLeft(m_fbsub(RShift(X(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
            RotateLeft(m_fbsub(RShift(X(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
        K = K + 1
    Next
    
    For i = 0 To m_Nb - 1
        j = i * 4
        UnpackFrom Y(i), Buff, j
        X(i) = 0
        Y(i) = 0
    Next
End Sub
#End If

'*******************************************************************************
' decrypt (SUB)
'*******************************************************************************
Public Sub Decrypt(Buff() As Byte)
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim m As Long
    Dim a(7) As Long
    Dim b(7) As Long
    Dim X() As Long
    Dim Y() As Long
    Dim t() As Long
    
    For i = 0 To m_Nb - 1
        j = i * 4
        a(i) = PackFrom(Buff, j)
        a(i) = a(i) Xor m_rkey(i)
    Next
    
    K = m_Nb
    X = a
    Y = b
    
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            m = j * 3
            Y(j) = m_rkey(K) Xor m_rtable(X(j) And m_lOnBits(7)) Xor _
                RotateLeft(m_rtable(RShift(X(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
                RotateLeft(m_rtable(RShift(X(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
                RotateLeft(m_rtable(RShift(X(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
            K = K + 1
        Next
        t = X
        X = Y
        Y = t
    Next
    
    For j = 0 To m_Nb - 1
        m = j * 3
        
        Y(j) = m_rkey(K) Xor m_rbsub(X(j) And m_lOnBits(7)) Xor _
            RotateLeft(m_rbsub(RShift(X(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
            RotateLeft(m_rbsub(RShift(X(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
            RotateLeft(m_rbsub(RShift(X(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
        K = K + 1
    Next
    
    For i = 0 To m_Nb - 1
        j = i * 4
        
        UnpackFrom Y(i), Buff, j
        X(i) = 0
        Y(i) = 0
    Next
End Sub

''*******************************************************************************
'' CopyBytesASP (SUB)
''
'' Slower non-API function you can use to copy array data
''*******************************************************************************
'Private Sub CopyBytesASP(bytDest() As Byte, _
'                         lDestStart As Long, _
'                         bytSource() As Byte, _
'                         lSourceStart As Long, _
'                         lLength As Long)
'    Dim lCount As Long
'
'    lCount = 0
'    Do
'        bytDest(lDestStart + lCount) = bytSource(lSourceStart + lCount)
'        lCount = lCount + 1
'    Loop Until lCount = lLength
'End Sub

'*******************************************************************************
' IsInitialized (FUNCTION)
'*******************************************************************************
'Private Function IsInitialized(ByRef vArray As Variant) As Boolean
'    On Error Resume Next
'
'    IsInitialized = IsNumeric(UBound(vArray))
'End Function
