Attribute VB_Name = "modA"
Option Explicit
Global byteKey(0 To 31) As Byte
Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Public Sub test()
'Dim NFF As String
'Dim TFF As String
'    TFF = Space$(444) '"123TestСукаТестБля"
'
'Dim ArrB() As Byte
'
'dataEncrypt TFF, ArrB
'Debug.Print UBound(ArrB)
'dataDecrypt ArrB, NFF
'Debug.Print Len(NFF)
'End Sub
'encrypt string and add zero-char to it's end
#If Enc Then
Sub dataEncrypt(ByRef StrToEncrypt As String, ByRef EmptyArray() As Byte)
    Dim AESobj As New clsA
    Dim TLNG&
    Dim ArrSize As Long
    Dim buff(0& To 31&) As Byte: Dim hBuff As Long: hBuff = VarPtr(buff(0))
    
    On Local Error Resume Next
'ABCDEFGHIJKLMNOPQRSTUVWXYZ
    'str to bytes
        EmptyArray() = StrConv(StrToEncrypt, vbFromUnicode)
    'add zero char
        ReDim Preserve EmptyArray(0& To UBound(EmptyArray) + 1) As Byte
        EmptyArray(UBound(EmptyArray)) = 0
    'round array bound to 256 bits
    ArrSize = UBound(EmptyArray): TLNG = ArrSize
        If Not TLNG Mod 32& = 31& Then
            Do While Not TLNG Mod 32& = 31&
                TLNG = TLNG + 1&
            Loop
            ReDim Preserve EmptyArray(0& To TLNG) As Byte
            'put some trash to the rest of array
            For TLNG = ArrSize + 1& To CLng(TLNG)
                ' EmptyArray(TLNG) = 1 + RNDINT(254)
                ' better leaving zeroes (maybe, here anyway doesn't matter but for simplicity)
                EmptyArray(TLNG) = 0
            Next TLNG
            ArrSize = UBound(EmptyArray)
        End If
    'encrypt
    With AESobj
        .gentables
        .gkey 8, 8, byteKey
        
        For TLNG = 0& To ArrSize Step 32
            CopyMemory ByVal hBuff, ByVal VarPtr(EmptyArray(TLNG)), 32&
            .Encrypt buff
            CopyMemory ByVal VarPtr(EmptyArray(TLNG)), ByVal hBuff, 32&
        Next TLNG
    End With
    Set AESobj = Nothing
        If Not Err.Number = 0& Then Erase EmptyArray
End Sub
#End If

Sub dataDecrypt(ByRef DataArray() As Byte, ByRef DecryptedStr As String)
    Dim AESobj As New clsA
    Dim TLNG As Long
    Dim ArrSize As Long
    Dim TByte As Byte
    Dim buff(0& To 31&) As Byte: Dim hBuff As Long: hBuff = VarPtr(buff(0))
    
    On Local Error Resume Next
    ArrSize = UBound(DataArray)

    'decrypt
    With AESobj
        .gentables
        .gkey 8, 8, byteKey
        
        For TLNG = 0& To ArrSize Step 32
            CopyMemory ByVal hBuff, ByVal VarPtr(DataArray(TLNG)), 32&
            .Decrypt buff
            CopyMemory ByVal VarPtr(DataArray(TLNG)), ByVal hBuff, 32&
            
            If TLNG Mod 32768 = 0& And Not TLNG = 0& Then DoEvents
        Next TLNG
    End With
    Set AESobj = Nothing
        If Not Err.Number = 0& Then
            Erase DataArray
        Else
            TLNG = UBound(DataArray): TByte = 0
            Do While TByte = 0
                TByte = DataArray(TLNG)
                TLNG = TLNG - 1&
            Loop
            ReDim Preserve DataArray(0& To TLNG + 1) As Byte
            DecryptedStr = StrConv(DataArray, vbUnicode, 0)
        End If
End Sub
Function RNDINT(ByVal rval As Long) As Long
    If rval < 0 Then
        RNDINT = Fix(Rnd * (rval - 1))
    Else
        RNDINT = Fix(Rnd * (rval + 1))
    End If
End Function
