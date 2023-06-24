VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// At top level of a module, always include to be sure that all
'variables have the right type
Option Explicit

Dim Crc32Table(255) As Long
Dim IsTableFull As Boolean

'// fill CRC32 table
Private Sub InitCrc32()
    Crc32Table(0) = &H0
    Crc32Table(1) = &H77073096
    Crc32Table(2) = &HEE0E612C
    Crc32Table(3) = &H990951BA
    Crc32Table(4) = &H76DC419
    Crc32Table(5) = &H706AF48F
    Crc32Table(6) = &HE963A535
    Crc32Table(7) = &H9E6495A3
    Crc32Table(8) = &HEDB8832
    Crc32Table(9) = &H79DCB8A4
    Crc32Table(10) = &HE0D5E91E
    Crc32Table(11) = &H97D2D988
    Crc32Table(12) = &H9B64C2B
    Crc32Table(13) = &H7EB17CBD
    Crc32Table(14) = &HE7B82D07
    Crc32Table(15) = &H90BF1D91
    Crc32Table(16) = &H1DB71064
    Crc32Table(17) = &H6AB020F2
    Crc32Table(18) = &HF3B97148
    Crc32Table(19) = &H84BE41DE
    Crc32Table(20) = &H1ADAD47D
    Crc32Table(21) = &H6DDDE4EB
    Crc32Table(22) = &HF4D4B551
    Crc32Table(23) = &H83D385C7
    Crc32Table(24) = &H136C9856
    Crc32Table(25) = &H646BA8C0
    Crc32Table(26) = &HFD62F97A
    Crc32Table(27) = &H8A65C9EC
    Crc32Table(28) = &H14015C4F
    Crc32Table(29) = &H63066CD9
    Crc32Table(30) = &HFA0F3D63
    Crc32Table(31) = &H8D080DF5
    Crc32Table(32) = &H3B6E20C8
    Crc32Table(33) = &H4C69105E
    Crc32Table(34) = &HD56041E4
    Crc32Table(35) = &HA2677172
    Crc32Table(36) = &H3C03E4D1
    Crc32Table(37) = &H4B04D447
    Crc32Table(38) = &HD20D85FD
    Crc32Table(39) = &HA50AB56B
    Crc32Table(40) = &H35B5A8FA
    Crc32Table(41) = &H42B2986C
    Crc32Table(42) = &HDBBBC9D6
    Crc32Table(43) = &HACBCF940
    Crc32Table(44) = &H32D86CE3
    Crc32Table(45) = &H45DF5C75
    Crc32Table(46) = &HDCD60DCF
    Crc32Table(47) = &HABD13D59
    Crc32Table(48) = &H26D930AC
    Crc32Table(49) = &H51DE003A
    Crc32Table(50) = &HC8D75180
    Crc32Table(51) = &HBFD06116
    Crc32Table(52) = &H21B4F4B5
    Crc32Table(53) = &H56B3C423
    Crc32Table(54) = &HCFBA9599
    Crc32Table(55) = &HB8BDA50F
    Crc32Table(56) = &H2802B89E
    Crc32Table(57) = &H5F058808
    Crc32Table(58) = &HC60CD9B2
    Crc32Table(59) = &HB10BE924
    Crc32Table(60) = &H2F6F7C87
    Crc32Table(61) = &H58684C11
    Crc32Table(62) = &HC1611DAB
    Crc32Table(63) = &HB6662D3D
    Crc32Table(64) = &H76DC4190
    Crc32Table(65) = &H1DB7106
    Crc32Table(66) = &H98D220BC
    Crc32Table(67) = &HEFD5102A
    Crc32Table(68) = &H71B18589
    Crc32Table(69) = &H6B6B51F
    Crc32Table(70) = &H9FBFE4A5
    Crc32Table(71) = &HE8B8D433
    Crc32Table(72) = &H7807C9A2
    Crc32Table(73) = &HF00F934
    Crc32Table(74) = &H9609A88E
    Crc32Table(75) = &HE10E9818
    Crc32Table(76) = &H7F6A0DBB
    Crc32Table(77) = &H86D3D2D
    Crc32Table(78) = &H91646C97
    Crc32Table(79) = &HE6635C01
    Crc32Table(80) = &H6B6B51F4
    Crc32Table(81) = &H1C6C6162
    Crc32Table(82) = &H856530D8
    Crc32Table(83) = &HF262004E
    Crc32Table(84) = &H6C0695ED
    Crc32Table(85) = &H1B01A57B
    Crc32Table(86) = &H8208F4C1
    Crc32Table(87) = &HF50FC457
    Crc32Table(88) = &H65B0D9C6
    Crc32Table(89) = &H12B7E950
    Crc32Table(90) = &H8BBEB8EA
    Crc32Table(91) = &HFCB9887C
    Crc32Table(92) = &H62DD1DDF
    Crc32Table(93) = &H15DA2D49
    Crc32Table(94) = &H8CD37CF3
    Crc32Table(95) = &HFBD44C65
    Crc32Table(96) = &H4DB26158
    Crc32Table(97) = &H3AB551CE
    Crc32Table(98) = &HA3BC0074
    Crc32Table(99) = &HD4BB30E2
    Crc32Table(100) = &H4ADFA541
    Crc32Table(101) = &H3DD895D7
    Crc32Table(102) = &HA4D1C46D
    Crc32Table(103) = &HD3D6F4FB
    Crc32Table(104) = &H4369E96A
    Crc32Table(105) = &H346ED9FC
    Crc32Table(106) = &HAD678846
    Crc32Table(107) = &HDA60B8D0
    Crc32Table(108) = &H44042D73
    Crc32Table(109) = &H33031DE5
    Crc32Table(110) = &HAA0A4C5F
    Crc32Table(111) = &HDD0D7CC9
    Crc32Table(112) = &H5005713C
    Crc32Table(113) = &H270241AA
    Crc32Table(114) = &HBE0B1010
    Crc32Table(115) = &HC90C2086
    Crc32Table(116) = &H5768B525
    Crc32Table(117) = &H206F85B3
    Crc32Table(118) = &HB966D409
    Crc32Table(119) = &HCE61E49F
    Crc32Table(120) = &H5EDEF90E
    Crc32Table(121) = &H29D9C998
    Crc32Table(122) = &HB0D09822
    Crc32Table(123) = &HC7D7A8B4
    Crc32Table(124) = &H59B33D17
    Crc32Table(125) = &H2EB40D81
    Crc32Table(126) = &HB7BD5C3B
    Crc32Table(127) = &HC0BA6CAD
    Crc32Table(128) = &HEDB88320
    Crc32Table(129) = &H9ABFB3B6
    Crc32Table(130) = &H3B6E20C
    Crc32Table(131) = &H74B1D29A
    Crc32Table(132) = &HEAD54739
    Crc32Table(133) = &H9DD277AF
    Crc32Table(134) = &H4DB2615
    Crc32Table(135) = &H73DC1683
    Crc32Table(136) = &HE3630B12
    Crc32Table(137) = &H94643B84
    Crc32Table(138) = &HD6D6A3E
    Crc32Table(139) = &H7A6A5AA8
    Crc32Table(140) = &HE40ECF0B
    Crc32Table(141) = &H9309FF9D
    Crc32Table(142) = &HA00AE27
    Crc32Table(143) = &H7D079EB1
    Crc32Table(144) = &HF00F9344
    Crc32Table(145) = &H8708A3D2
    Crc32Table(146) = &H1E01F268
    Crc32Table(147) = &H6906C2FE
    Crc32Table(148) = &HF762575D
    Crc32Table(149) = &H806567CB
    Crc32Table(150) = &H196C3671
    Crc32Table(151) = &H6E6B06E7
    Crc32Table(152) = &HFED41B76
    Crc32Table(153) = &H89D32BE0
    Crc32Table(154) = &H10DA7A5A
    Crc32Table(155) = &H67DD4ACC
    Crc32Table(156) = &HF9B9DF6F
    Crc32Table(157) = &H8EBEEFF9
    Crc32Table(158) = &H17B7BE43
    Crc32Table(159) = &H60B08ED5
    Crc32Table(160) = &HD6D6A3E8
    Crc32Table(161) = &HA1D1937E
    Crc32Table(162) = &H38D8C2C4
    Crc32Table(163) = &H4FDFF252
    Crc32Table(164) = &HD1BB67F1
    Crc32Table(165) = &HA6BC5767
    Crc32Table(166) = &H3FB506DD
    Crc32Table(167) = &H48B2364B
    Crc32Table(168) = &HD80D2BDA
    Crc32Table(169) = &HAF0A1B4C
    Crc32Table(170) = &H36034AF6
    Crc32Table(171) = &H41047A60
    Crc32Table(172) = &HDF60EFC3
    Crc32Table(173) = &HA867DF55
    Crc32Table(174) = &H316E8EEF
    Crc32Table(175) = &H4669BE79
    Crc32Table(176) = &HCB61B38C
    Crc32Table(177) = &HBC66831A
    Crc32Table(178) = &H256FD2A0
    Crc32Table(179) = &H5268E236
    Crc32Table(180) = &HCC0C7795
    Crc32Table(181) = &HBB0B4703
    Crc32Table(182) = &H220216B9
    Crc32Table(183) = &H5505262F
    Crc32Table(184) = &HC5BA3BBE
    Crc32Table(185) = &HB2BD0B28
    Crc32Table(186) = &H2BB45A92
    Crc32Table(187) = &H5CB36A04
    Crc32Table(188) = &HC2D7FFA7
    Crc32Table(189) = &HB5D0CF31
    Crc32Table(190) = &H2CD99E8B
    Crc32Table(191) = &H5BDEAE1D
    Crc32Table(192) = &H9B64C2B0
    Crc32Table(193) = &HEC63F226
    Crc32Table(194) = &H756AA39C
    Crc32Table(195) = &H26D930A
    Crc32Table(196) = &H9C0906A9
    Crc32Table(197) = &HEB0E363F
    Crc32Table(198) = &H72076785
    Crc32Table(199) = &H5005713
    Crc32Table(200) = &H95BF4A82
    Crc32Table(201) = &HE2B87A14
    Crc32Table(202) = &H7BB12BAE
    Crc32Table(203) = &HCB61B38
    Crc32Table(204) = &H92D28E9B
    Crc32Table(205) = &HE5D5BE0D
    Crc32Table(206) = &H7CDCEFB7
    Crc32Table(207) = &HBDBDF21
    Crc32Table(208) = &H86D3D2D4
    Crc32Table(209) = &HF1D4E242
    Crc32Table(210) = &H68DDB3F8
    Crc32Table(211) = &H1FDA836E
    Crc32Table(212) = &H81BE16CD
    Crc32Table(213) = &HF6B9265B
    Crc32Table(214) = &H6FB077E1
    Crc32Table(215) = &H18B74777
    Crc32Table(216) = &H88085AE6
    Crc32Table(217) = &HFF0F6A70
    Crc32Table(218) = &H66063BCA
    Crc32Table(219) = &H11010B5C
    Crc32Table(220) = &H8F659EFF
    Crc32Table(221) = &HF862AE69
    Crc32Table(222) = &H616BFFD3
    Crc32Table(223) = &H166CCF45
    Crc32Table(224) = &HA00AE278
    Crc32Table(225) = &HD70DD2EE
    Crc32Table(226) = &H4E048354
    Crc32Table(227) = &H3903B3C2
    Crc32Table(228) = &HA7672661
    Crc32Table(229) = &HD06016F7
    Crc32Table(230) = &H4969474D
    Crc32Table(231) = &H3E6E77DB
    Crc32Table(232) = &HAED16A4A
    Crc32Table(233) = &HD9D65ADC
    Crc32Table(234) = &H40DF0B66
    Crc32Table(235) = &H37D83BF0
    Crc32Table(236) = &HA9BCAE53
    Crc32Table(237) = &HDEBB9EC5
    Crc32Table(238) = &H47B2CF7F
    Crc32Table(239) = &H30B5FFE9
    Crc32Table(240) = &HBDBDF21C
    Crc32Table(241) = &HCABAC28A
    Crc32Table(242) = &H53B39330
    Crc32Table(243) = &H24B4A3A6
    Crc32Table(244) = &HBAD03605
    Crc32Table(245) = &HCDD70693
    Crc32Table(246) = &H54DE5729
    Crc32Table(247) = &H23D967BF
    Crc32Table(248) = &HB3667A2E
    Crc32Table(249) = &HC4614AB8
    Crc32Table(250) = &H5D681B02
    Crc32Table(251) = &H2A6F2B94
    Crc32Table(252) = &HB40BBE37
    Crc32Table(253) = &HC30C8EA1
    Crc32Table(254) = &H5A05DF1B
    Crc32Table(255) = &H2D02EF8D
End Sub

'C language A >> B operator analog
Private Function RShiftLong(ByVal Value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    If (Value And &H80000000) Then hi = &H40000000
  
    RShiftLong = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
End Function



'Check : 0xCBF43926 ("123456789")
Public Function CRC32(ByRef incMsg1 As String) As Long
    Dim Tlng As Long    'loop counter
    
    'fill our table using common polynome 0x04C11DB7  (x^32 + x^26 + x^23 + x^22 + x^16 + x^12 + x^11 + x^10 + x^8 + x^7 + x^5 + x^4 + x^2 + x + 1)
    If Not IsTableFull Then
        IsTableFull = True
            Call InitCrc32
    End If
    
    'set initial value
    CRC32 = &HFFFFFFFF

    'Get CRC
    For Tlng = 1 To Len(incMsg1)
        '// Right shift 8 bits
        CRC32 = RShiftLong(CRC32, 8) Xor Crc32Table((CRC32 Xor (Asc(Mid$(incMsg1, Tlng, 1&)))) And &HFF)
    Next Tlng
    CRC32 = CRC32 Xor &HFFFFFFFF

'    Debug.Print Hex(CRC32)
End Function

