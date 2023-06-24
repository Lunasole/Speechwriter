Attribute VB_Name = "modRName"
Option Explicit
Function rName() As String
Dim nLen As Integer: nLen = 4 + (RNDINT(10))

If Rnd <= 0.07 Then
    rName = Trim$(Now)
Else
    For nLen = 1 To nLen
      If Rnd <= 0.8 Then
        rName = rName & Chr$(32 + RNDINT(223)) 'random
      ElseIf Rnd <= 0.5 Then
        rName = rName & Chr$(97 + RNDINT(25)) 'a-z
      Else
        rName = rName & Chr$(65 + RNDINT(25)) 'A-Z
      End If
    Next nLen
      rName = Replace$(rName, Chr$(34&), vbNullString, 1, -1, vbBinaryCompare)
End If

End Function

Function nName(ByRef nLen As Long, ByRef StartMap As Byte, ByRef UpCase As Boolean) As String
'"abcdefghijklmnopqrstuvwxyz"

'English mode not finished
'Table #0 and #1 maps
'If letter repeats, it chance to be used raising up
'Const cpTable0EMap As String = "aeiouy"
'Const cpTable1EMap As String = "bcdfghjklmnpqrstvwxz"


'a
'eiouy
'
'e
'aiou 'y
'
'i
'aeouy



'Table #0 and #1 maps
'If letter repeats, it chance to be used raising up
Const cpTable0Map As String = "ààååèèîîóóû"
Const cpTable1Map As String = "áâãäæçêëìíïðñòôõö÷øùé"



'Compactibility table #0
Const cp224 As String = "åèîóýþÿ"       'à
Const cp229 As String = "þÿ"            'å
Const cp232 As String = "åþÿ"           'è
Const cp238 As String = "åþÿ"           'î
Const cp243 As String = "åèþÿ"          'ó
Const cp251 As String = "åè"            'û
'Const cp253 As String = ""              'ý
'Const cp254 As String = ""              'þ
'Const cp255 As String = ""              'ÿ

'Compactibility table #1
Const cp225 As String = "äçëð"          'á
Const cp226 As String = "áæêëìíïðñöø"   'â
Const cp227 As String = "âäëíð"         'ã
Const cp228 As String = "âæíðø"         'ä
Const cp230 As String = "áäëìð"         'æ
Const cp231 As String = "âäëí"          'ç
Const cp234 As String = "âëìíð" '& ö    'ê
Const cp235 As String = "æêñ÷"    '& ä  'ë
Const cp236 As String = "áæäêëð"        'ì
Const cp237 As String = "ãäöí"          'í
Const cp239 As String = "ëìíðò÷ø"       'ï
Const cp240 As String = "âäæçêëìíò÷ù"   'ð
Const cp241 As String = "âäæêëìíïðòõöø" 'ñ
Const cp242 As String = "âêëìðø"        'ò
Const cp244 As String = cp225           'ô
Const cp245 As String = "âëíðò÷"        'õ
'Const cp246 As String = ""              'ö
Const cp247 As String = "âëïðò"         '÷
Const cp248 As String = "âêëìíðò"       'ø
Const cp249 As String = cp248           'ù
'----------------------------------------'

'Word 1st letter maps
Const cpStart0Map As String = "àáâãäåæçèêëìíîïðñòóôõö÷øùýþÿ"
Const cpStart1Map As String = "áâãäæçêëìíïðñòôõö÷øù"
Const cpStart2Map As String = "àåèîóýþÿ"

'loop var
Dim TLNG As Long

'Currently selected char code
Dim CChar As Byte
'Previously selected char code
Dim RChar As Byte
'If true, there will be switch between table #0 and #1 due current loop
Dim bSwap As Boolean
'Swap count
'Dim bSwCt As Integer
'If true, current char will be used again due next loop
Dim bReply As Boolean



If nLen < 2 Then Exit Function
    nName = Space$(nLen)
    
For TLNG = 1& To nLen
    If TLNG = 1& Then
        'init
        Select Case StartMap
            Case 1: CChar = Asc(Mid$(cpStart1Map, 1& + RNDINT(Len(cpStart1Map) - 1&), 1&))
            Case 2: CChar = Asc(Mid$(cpStart2Map, 1& + RNDINT(Len(cpStart2Map) - 1&), 1&))
            Case Else: CChar = Asc(Mid$(cpStart0Map, 1& + RNDINT(Len(cpStart0Map) - 1&), 1&))
        End Select
        GoTo jmpPut
    Else

        If Not bReply Then
            If CChar = 253 Or CChar = 254 Or CChar = 255 Or CChar = 246 Then
'            If CChar = 253 Or CChar = 254 Or CChar = 255 Or CChar = 246 Or (Tlng = nLen And bSwCt = 0) Then
                bSwap = True
            Else
                bSwap = Rnd <= 0.9
            End If
            
            If Not bSwap Then
                Select Case CChar
                    Case 224:
                        CChar = Asc(Mid$(cp224, 1& + RNDINT(Len(cp224) - 1&), 1&))
                    Case 229:
                        CChar = Asc(Mid$(cp229, 1& + RNDINT(Len(cp229) - 1&), 1&))
                    Case 232:
                        CChar = Asc(Mid$(cp232, 1& + RNDINT(Len(cp232) - 1&), 1&))
                    Case 238:
                        CChar = Asc(Mid$(cp238, 1& + RNDINT(Len(cp238) - 1&), 1&))
                    Case 243:
                        CChar = Asc(Mid$(cp243, 1& + RNDINT(Len(cp243) - 1&), 1&))
                    Case 251:
                        CChar = Asc(Mid$(cp251, 1& + RNDINT(Len(cp251) - 1&), 1&))
                    Case 225:
                        CChar = Asc(Mid$(cp225, 1& + RNDINT(Len(cp225) - 1&), 1&))
                    Case 226:
                        CChar = Asc(Mid$(cp226, 1& + RNDINT(Len(cp226) - 1&), 1&))
                    Case 227:
                        CChar = Asc(Mid$(cp227, 1& + RNDINT(Len(cp227) - 1&), 1&))
                    Case 228:
                        CChar = Asc(Mid$(cp228, 1& + RNDINT(Len(cp228) - 1&), 1&))
                    Case 230:
                        CChar = Asc(Mid$(cp230, 1& + RNDINT(Len(cp230) - 1&), 1&))
                    Case 231:
                        CChar = Asc(Mid$(cp231, 1& + RNDINT(Len(cp231) - 1&), 1&))
                    Case 234:
                        CChar = Asc(Mid$(cp234, 1& + RNDINT(Len(cp234) - 1&), 1&))
                    Case 235:
                        CChar = Asc(Mid$(cp235, 1& + RNDINT(Len(cp235) - 1&), 1&))
                    Case 236:
                        CChar = Asc(Mid$(cp236, 1& + RNDINT(Len(cp236) - 1&), 1&))
                    Case 237:
                        CChar = Asc(Mid$(cp237, 1& + RNDINT(Len(cp237) - 1&), 1&))
                    Case 239:
                        CChar = Asc(Mid$(cp239, 1& + RNDINT(Len(cp239) - 1&), 1&))
                    Case 240:
                        CChar = Asc(Mid$(cp240, 1& + RNDINT(Len(cp240) - 1&), 1&))
                    Case 241:
                        CChar = Asc(Mid$(cp241, 1& + RNDINT(Len(cp241) - 1&), 1&))
                    Case 242:
                        CChar = Asc(Mid$(cp242, 1& + RNDINT(Len(cp242) - 1&), 1&))
                    Case 244:
                        CChar = Asc(Mid$(cp244, 1& + RNDINT(Len(cp244) - 1&), 1&))
                    Case 245:
                        CChar = Asc(Mid$(cp245, 1& + RNDINT(Len(cp245) - 1&), 1&))
                    Case 247:
                        CChar = Asc(Mid$(cp247, 1& + RNDINT(Len(cp247) - 1&), 1&))
                    Case 248:
                        CChar = Asc(Mid$(cp248, 1& + RNDINT(Len(cp248) - 1&), 1&))
                    Case 249:
                        CChar = Asc(Mid$(cp249, 1& + RNDINT(Len(cp249) - 1&), 1&))
                End Select
                
            Else
'                bSwCt = bSwCt + 1
                
                'this is for table #0 chars that do not have "compactibility" constant
                If CChar = 224 Or CChar = 229 Or CChar = 232 Or CChar = 238 _
                     Or CChar = 243 Or CChar = 251 Or CChar = 253 Or CChar = 254 _
                      Or CChar = 255 Then
                      
                        If TLNG = nLen Then
                            CChar = Asc(Mid$(cpTable1Map, 1& + RNDINT(Len(cpTable1Map) - 1&), 1&))
                        Else
                            CChar = Asc(Mid$(cpTable1Map, 1& + RNDINT(Len(cpTable1Map) - 2&), 1&)) 'without é
                        End If
                Else 'normal chars
                        If TLNG = nLen Then
                            CChar = Asc(Mid$(cpTable0Map, 1& + RNDINT(Len(cpTable0Map) - 1&), 1&))
                        Else
                            CChar = Asc(Mid$(cpTable0Map, 1& + RNDINT(Len(cpTable0Map) - 2&), 1&)) 'without û
                        End If
'                    CChar = Asc(Mid$(cpTable0Map, 1& + RNDINT(Len(cpTable0Map) - 1&), 1&))
                End If
                
            End If
        End If
    
    If Not CChar = RChar Or bReply Then
jmpPut:
        RChar = CChar
        If Not bReply Then
            If CChar = 224 Or CChar = 237 Then 'double à & í
                bReply = Rnd <= 0.04
            ElseIf CChar = 232 Or CChar = 238 Then 'double è & î
                bReply = Rnd <= 0.02
            End If
        Else
            bReply = False
        End If
            Mid$(nName, TLNG, 1&) = Chr$(CChar)
    Else
        TLNG = TLNG - 1&
    End If
    
    End If
Next TLNG

    If UpCase Then Mid$(nName, 1&, 1&) = UCase$(Mid$(nName, 1&, 1&))
End Function

'Public Sub rRewrite(ByRef nName As String)
''If letter repeats, it chance to be used raising up
'Const cpTable0Map As String = "ààååèèîîóóû"
'Const cpTable1Map As String = "áâãäæçêëìíïðñòôõö÷øùé"
'
''Word 1st letter maps
'Const cpStart0Map As String = "àáâãäåæçèêëìíîïðñòóôõö÷øùýþÿ"
'Const cpStart1Map As String = "áâãäæçêëìíïðñòôõö÷øù"
'Const cpStart2Map As String = "àåèîóýþÿ"
'
''Currently selected char code
'Dim CChar As Byte
'Dim Tlng&
'
'
'For Tlng = 1& To Len(nName)
'    CChar = Asc(Mid$(nName, Tlng, 1&))
'
'        Select Case CChar
'            Case 224:
'                CChar = 0
'            Case 229:
'                CChar = 0
'            Case 232:
'                CChar = 0
'            Case 238:
'                CChar = 0
'            Case 243:
'                CChar = 0
'            Case 251:
'                CChar = 0
'            Case 253:
'                CChar = 0
'            Case 254:
'                CChar = 0
'            Case 255:
'                CChar = 0
'
'            Case 225:
'                CChar = 1
'            Case 226:
'                CChar = 1
'            Case 227:
'                CChar = 1
'            Case 228:
'                CChar = 1
'            Case 230:
'                CChar = 1
'            Case 231:
'                CChar = 1
'            Case 234:
'                CChar = 1
'            Case 235:
'                CChar = 1
'            Case 236:
'                CChar = 1
'            Case 237:
'                CChar = 1
'            Case 239:
'                CChar = 1
'            Case 240:
'                CChar = 1
'            Case 241:
'                CChar = 1
'            Case 242:
'                CChar = 1
'            Case 244:
'                CChar = 1
'            Case 245:
'                CChar = 1
'            Case 246:
'                CChar = 1
'            Case 247:
'                CChar = 1
'            Case 248:
'                CChar = 1
'            Case 249:
'                CChar = 1
'        End Select
'
'        If CChar = 0 Then
'            If Tlng = 1& Then
'                CChar = Asc(Mid$(cpStart2Map, 1& + RNDINT(Len(cpStart2Map) - 1&), 1&))
'            Else
'                CChar = Asc(Mid$(cpTable0Map, 1& + RNDINT(Len(cpTable0Map) - 1&), 1&))
'            End If
'        ElseIf CChar = 1 Then
'            If Tlng = 1& Then
'                CChar = Asc(Mid$(cpStart1Map, 1& + RNDINT(Len(cpStart1Map) - 1&), 1&))
'            ElseIf Tlng = Len(nName) Then
'                CChar = Asc(Mid$(cpTable1Map, 1& + RNDINT(Len(cpTable1Map) - 1&), 1&))
'            Else
'                CChar = Asc(Mid$(cpTable1Map, 1& + RNDINT(Len(cpTable1Map) - 2&), 1&)) 'without é
'            End If
'        Else
'
''        Stop
'        End If
'
'        Mid$(nName, Tlng, 1&) = Chr$(CChar)
'Next Tlng
'End Sub
