Attribute VB_Name = "modB"
Option Explicit

Function GetHashByte(ByRef Data() As Byte, ModOffset) As Byte
    Dim TLNG&
    Data(LBound(Data) + 240 + ModOffset) = Data(UBound(Data) / 2) 'skip hash byte
        For TLNG = LBound(Data) To UBound(Data) Step 2
            GetHashByte = GetHashByte Xor Data(TLNG)
        Next TLNG
End Function
Function ExeHash(ByRef iFile As String, ByRef hWrite As Boolean, Optional ModOffset As Long) As Boolean
Dim CHash As Byte
Dim HashByte As Byte
Dim Bytes() As Byte
On Local Error GoTo lErr
Open iFile For Binary As 1
    ReDim Bytes(1& To LOF(1))
    Get 1, 1, Bytes


    If Not (Bytes(1) = 77) Or Not (Bytes(2) = 90) Or Not (Bytes(3) = 144) Or Not (Bytes(4) = 0) Then
        'is not an exe file
    Else
        CHash = Bytes(241 + ModOffset)  'hash byte offset
        HashByte = GetHashByte(Bytes, ModOffset)
        If hWrite Then
            Put 1, 241 + ModOffset, HashByte
            CHash = HashByte
        End If
        ExeHash = CHash = HashByte
    End If
lErr:
Close 1
End Function
