VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8880
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3150
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1185
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Опечатки"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   3150
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Не деай этого, ссука"
      Top             =   4350
      Width           =   1185
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Получить рифмующиеся слова"
      Top             =   3960
      Width           =   405
   End
   Begin VB.Timer TmrAUX 
      Interval        =   40
      Left            =   30
      Top             =   30
   End
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   7815
   End
   Begin VB.TextBox txtCmd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   30
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      ToolTipText     =   "Типа консолька, можно даже пердолиться."
      Top             =   3960
      Width           =   6555
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   7440
      TabIndex        =   13
      ToolTipText     =   "Генерация по заданному сиду"
      Top             =   3960
      Width           =   405
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "L"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7020
      TabIndex        =   12
      ToolTipText     =   "Получить рифмующиеся строки"
      Top             =   3960
      Width           =   405
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   2340
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Превращение некоторых слов в другие"
      Top             =   4350
      Width           =   705
   End
   Begin VB.PictureBox imgRandom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   7890
      ScaleHeight     =   120
      ScaleMode       =   0  'User
      ScaleWidth      =   30
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "UTF-8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7920
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Копировать в UTF-8. Если ОС виста/спермерка или выше, вроде как обязательно."
      Top             =   360
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Кэш"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7920
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Избегает недавно использованных фраз, если уникальных строк достаточно"
      Top             =   720
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "CP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7920
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Автоматически копировать"
      Top             =   0
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "а < a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   7920
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Замена некоторых символов кириллицы на латиницу, иногда полезно"
      Top             =   1050
      Width           =   1035
   End
   Begin VB.TextBox txtOutLine 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   30
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      ToolTipText     =   "Эпилог,  \n = новая строка"
      Top             =   4650
      Width           =   645
   End
   Begin VB.TextBox txtInLine 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   30
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      ToolTipText     =   "Пролог,  \n = новая строка"
      Top             =   4305
      Width           =   645
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   124
      Left            =   6990
      TabIndex        =   9
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   6120
      TabIndex        =   8
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   5250
      TabIndex        =   7
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   4380
      TabIndex        =   6
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   3510
      TabIndex        =   5
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2640
      TabIndex        =   4
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1770
      TabIndex        =   3
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   900
      TabIndex        =   2
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   3540
      Width           =   855
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Энтропия"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   6660
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Вечное сияние чистого рандома"
      Top             =   4710
      Width           =   1275
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Aa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   780
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Отключить CAPSLOCK"
      Top             =   4710
      Value           =   1  'Checked
      Width           =   585
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Куклотеги"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5340
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Классические теги: **%%"
      Top             =   4710
      Width           =   1215
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Хаос"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   4440
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Загрузка..."
      Top             =   4710
      Width           =   795
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "IQ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   2340
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Большие проблемы с речью..."
      Top             =   4710
      Width           =   705
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "§"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   780
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Больше единого разрыва"
      Top             =   4350
      Value           =   1  'Checked
      Width           =   585
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   ". , ?!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1470
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Пунктуация"
      Top             =   4350
      Value           =   1  'Checked
      Width           =   765
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Омск"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Viva la Omsk!"
      Top             =   4350
      Width           =   795
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Разметка"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5340
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Теги вида [b][i][u]"
      Top             =   4350
      Width           =   1215
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "Лирика"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   6660
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Ехал быдло через быдло - быдло быдло быдло быдло"
      Top             =   4350
      Width           =   1275
   End
   Begin VB.CheckBox opt 
      BackColor       =   &H80000005&
      Caption         =   "аАа"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   1470
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "СлУЧайнЫй рЕГистР"
      Top             =   4710
      Width           =   765
   End
   Begin VB.Label lblBack 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   2
      Left            =   7890
      TabIndex        =   18
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblBack 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   -210
      Width           =   7965
   End
   Begin VB.Menu mnuMain 
      Caption         =   ".."
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Открыть "
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Перезагрузить"
      End
      Begin VB.Menu mnuClean 
         Caption         =   "Форматировать файл"
      End
      Begin VB.Menu mnuDiv0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeSource 
         Caption         =   "Сменить базу"
         Begin VB.Menu mnuFileEn 
            Caption         =   "[SW: стандартный]"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileEn 
            Caption         =   "[SW: треш]"
            Index           =   2
         End
         Begin VB.Menu mnuFileEn 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFile 
            Caption         =   ""
            Index           =   20
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Создать базу"
      End
      Begin VB.Menu mnuDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Сохранить в ""+out.txt"""
      End
      Begin VB.Menu mnuOut 
         Caption         =   "Открыть ""+out.txt"""
      End
      Begin VB.Menu mnuDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   ">"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Развернуть"
      End
      Begin VB.Menu mnuTrayDiv00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayGenA 
         Caption         =   "[ Авто ]"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrayGen 
         Caption         =   "Сгенерировать"
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuTrayGenT 
            Caption         =   ""
            Index           =   8
         End
      End
      Begin VB.Menu mnuTrayDiv01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Выход"
      End
      Begin VB.Menu mnuTrayNothing 
         Caption         =   "Ничего..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Speech
    MArr() As String        'usual strings
    MArrRyphm() As String   'lyrics strings
End Type

Dim Scriptkiddie As Long

Dim SWArr As Speech
Dim SWCnt As Long
Dim LDate As String 'base file time
Dim InnerFile As Byte 'value <> 0 means external file, 0 = base loaded from .res

'Cache
Dim sUsed As String
Dim sCnt As Long
Dim MaxUsed As Long 'Cache size

'8 bits masks
Private Const BIT1 As Byte = &H1: Private Const BIT2 As Byte = &H2: Private Const BIT3 As Byte = &H4: Private Const BIT4 As Byte = &H8: Private Const BIT5 As Byte = &H10: Private Const BIT6 As Byte = &H20: Private Const BIT7 As Byte = &H40: Private Const BIT8 As Byte = &H80
'config structure
Private Type Config
    FLAG1 As Byte
    FLAG2 As Byte
    FLAG3 As Byte
    FLAG4 As Byte

    sInlineText As String
    sOulineText As String
    sSourceName As String
    sCMD As String
End Type

'Temp config scope
Dim opOnline As Boolean 'if enabled, speechwriter acts as tcp server

Dim opCopyPaste As Boolean
Dim opCopyPasteUTF8 As Boolean
Dim opPunctuation As Boolean
Dim opLowerCase As Boolean
Dim opCache As Boolean
Dim opOmsk As Boolean
Dim opEmptyLines As Boolean
Dim opLetterFix As Boolean
Dim opTags As Boolean
Dim opTagsDoll As Boolean
Dim opLyrics As Boolean
Dim opChaos As Boolean
Dim opFullRandom As Boolean
Dim opBaar As Boolean
Dim opRewrite As Boolean
Dim opRandomCase As Boolean
Dim opMistakes As Boolean

Private sConfig As Config
Private Const sCfgName = "speechwriter.cfg"

Private SW_OpState As Byte 'Operation state: 1 - autogen, 2 - file operations
Public AutoGen As Integer

Dim AppSPath As String 'app.path & \

'''''''''''''''''''
Const appMajor As Byte = 1
Const appMinor As Byte = 0
Const appBuild As Byte = 72

Dim hIcon As Long ' handle to an app icon extracted with WINAPI

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long


Private Sub cmdC_Click(Index As Integer)
    Select Case Index
        Case 0: Call cmdRyphm(txtCmd, False)
        Case 1: Call cmdRyphm(txtCmd, True)
        Case 2: Call cmdShitByKey(txtCmd)
    End Select
    cmdC(Index).Default = True
    txtCmd.Tag = Index
End Sub

Private Sub Form_Initialize()
InitXPstyler
If App.PrevInstance Then End
Randomize

On Error Resume Next
    AppSPath = App.Path & "\"
    hIcon = ExtractIcon(0, AppSPath & App.EXEName & ".exe", 0)
    
    If Not PathFileExists(AppSPath & sCfgName) = 0& Then
        Open AppSPath & sCfgName For Binary As 1
            Get 1, 1, sConfig
        Close 1
        With sConfig
            opt(0).Value = orCheck(.FLAG1, BIT1)
            opt(1).Value = orCheck(.FLAG1, BIT2)
            opt(2).Value = orCheck(.FLAG1, BIT3)
            opt(3).Value = orCheck(.FLAG1, BIT4)
            opt(4).Value = orCheck(.FLAG1, BIT5)
            opt(5).Value = orCheck(.FLAG1, BIT6)
            opt(6).Value = orCheck(.FLAG1, BIT7)
            opt(7).Value = orCheck(.FLAG1, BIT8)
            
            opt(8).Value = orCheck(.FLAG2, BIT1)
            opt(9).Value = orCheck(.FLAG2, BIT2)

            opt(10).Value = orCheck(.FLAG2, BIT3)

            opt(11).Value = orCheck(.FLAG2, BIT4)
            opt(12).Value = orCheck(.FLAG2, BIT5)
            opt(13).Value = orCheck(.FLAG2, BIT6)
            opt(14).Value = orCheck(.FLAG2, BIT7)
                        If orCheck(.FLAG2, BIT8) = 1 Then imgRandom_DblClick
            opt(15).Value = orCheck(.FLAG3, BIT1)
            opt(16).Value = orCheck(.FLAG3, BIT2)

            txtInLine.Text = .sInlineText
            txtOutLine.Text = .sOulineText
            txtCmd = .sCMD
        End With
    End If
        
        'do some anti-scriptkiddie checks (enabled when not commandline = x'\, so use that cmdline to prevent crashes when working at IDE)
        'don't forget apply HashExeByte.exe to compiled file!
        
        If Not StrComp(Command$, "x'\") = 0& Then
            If Not IsDebuggerPresent = 0& Then End
            If Not ExeHash(AppSPath & App.EXEName & ".exe", False, 48) Then
                Dim strDeath As String
                Scriptkiddie = Timer
                Call CopyMemory(ByVal StrPtr(strDeath), 0, 4&) 'obj data nullified, so good night
                While 1 = 1
                    Call CopyMemory(ByVal StrPtr(strDeath) + Rnd() * 2000000, 0, 4&) 'obj data nullified, so good night
                    strDeath = strDeath & "1"
                    DoEvents
                Wend
            Else
                Call hook(Me.hwnd)
            End If
        End If
        
        Call zcs_UpdateTempCfg
        Call TmrAUX_Timer

        'check form
        If Not orCheck(sConfig.FLAG2, BIT8) = 1 Then
            If Not frmMain.Visible Then frmMain.Show:            DoEvents
        End If
        
        'set random caption for IQ
        opt(13).Caption = "IQ" & RNDINT(3)
        
        Call a_loadFile
End Sub

'simple str encryption/decryption
Private Function zcf_nStr$(ByRef pNStr$)
zcf_nStr = pNStr: Dim sEn&: Dim tXr As Byte: tXr = (30 And Len(zcf_nStr)) + 1
  For sEn = 1& To Len(zcf_nStr)
    Mid$(zcf_nStr, sEn, 1&) = Chr(Asc(Mid$(zcf_nStr, sEn, 1&)) Xor tXr)
  Next
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If InStr(1&, "txtCmdtxtInLinetxtOutLine", ActiveControl.Name, vbBinaryCompare) = 0& Then
        Call cmdS_KeyUp(0, KeyCode, 0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
    If Not MainTray.UserID = 0& Then MainTray.Remove
        Call unhook(Me.hwnd)
    DestroyIcon (hIcon)
    
    With sConfig
        orSet .FLAG1, BIT1, opt(0).Value
        orSet .FLAG1, BIT2, opt(1).Value
        orSet .FLAG1, BIT3, opt(2).Value
        orSet .FLAG1, BIT4, opt(3).Value
        orSet .FLAG1, BIT5, opt(4).Value
        orSet .FLAG1, BIT6, opt(5).Value
        orSet .FLAG1, BIT7, opt(6).Value
        orSet .FLAG1, BIT8, opt(7).Value
        
        orSet .FLAG2, BIT1, opt(8).Value
        orSet .FLAG2, BIT2, opt(9).Value
        orSet .FLAG2, BIT3, opt(10).Value
        orSet .FLAG2, BIT4, opt(11).Value
        orSet .FLAG2, BIT5, opt(12).Value
        orSet .FLAG2, BIT6, opt(13).Value
        orSet .FLAG2, BIT7, opt(14).Value
        orSet .FLAG2, BIT8, IIf(frmMain.Visible, 0, 1)
        
        orSet .FLAG3, BIT1, opt(15).Value
        orSet .FLAG3, BIT2, opt(16).Value
        
        .sInlineText = txtInLine.Text
        .sOulineText = txtOutLine.Text
        .sCMD = txtCmd
    End With
    
    frmMain.Visible = False
    
    Call DeleteFile(AppSPath & sCfgName)
        Open AppSPath & sCfgName For Binary As 1
            Put 1, 1, sConfig
        Close 1
End Sub

Private Sub imgRandom_DblClick()
    If Not MainTray.Add(Me.hwnd, hIcon, App.Title & IIf(AutoGen = 0, vbNullString, ": АКТИВИРОВАН"), 1&) = 0& Then
        Me.Visible = False
    End If
End Sub


Private Sub imgRandom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) And (SW_OpState = 0&) Then
        Call ShowMenu
    End If
End Sub


Private Sub mnuAbout_Click()
    Call aux_showAbout
End Sub

Private Sub mnuClean_Click()
    Call SHIT_CLEANUP
End Sub


Private Sub mnuFile_Click(Index As Integer)
    If Not Index > FileCount And Not Index = 0 Then
        sConfig.sSourceName = FileList(Index)
            Call a_loadFile
    End If
        Erase FileList
        FileCount = 0&
End Sub


Private Sub mnuFileEn_Click(Index As Integer)
    If Not Index > 2 And Not Index <= 0 Then
        sConfig.sSourceName = mnuFileEn(Index).Caption
            Call a_loadFile
    End If
End Sub

Private Sub mnuImport_Click()
    MsgBox "Перетащите текстовый файл с проводника в окно программы.", vbInformation
End Sub


Private Sub mnuOpen_Click()
    Call ShellExecute(0&, vbNullString, AppSPath & sConfig.sSourceName, vbNullString, App.Path, vbNormalFocus)
End Sub

Private Sub mnuOut_Click()
    Call ShellExecute(0&, vbNullString, App.Path & "\+out.txt", vbNullString, App.Path, vbNormalFocus)
End Sub

Private Sub mnuReload_Click()
    Call a_loadFile
End Sub

Private Sub cmdShitByKey(ByRef Key As String)
If Not SW_OpState = 0 Then Exit Sub
If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)

Dim USEED As Long
Dim HCrc As New clsC
    If Not Len(Key) = 0& Then
        USEED = HCrc.CRC32(Key)
            Call SHIT(0&, USEED)
    End If
Set HCrc = Nothing
End Sub

Private Sub cmdRyphm(ByRef IWord As String, ByRef RyphmString As Boolean)

If Not SW_OpState = 0 Then Exit Sub
If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)

Dim NRyphm As String
Dim Result As String
Dim ResultB As Byte
Dim ResultMaxB As Byte: ResultMaxB = 32 + RNDINT(16)

Dim SeekResultsL(0& To 47&) As Long
Dim SeekResults() As String
Dim SeekCount As Long
Dim TLNG As Long
    If Not Len(IWord) = 0& Then
        IWord = fc_GetLastWord(IWord, True, False, 0&)
        IWord = UCase$(IWord)
        
        
        If Not RyphmString Then 'получить рифмующиеся к слову слова
            SeekCount = fc_RyphmWord(IWord, 3&, SeekResults, True, 6000)
                If Not SeekCount = -1& Then
                    For TLNG = 0& To 32&
                        NRyphm = IWord & " — " & SeekResults(RNDINT(SeekCount)) & vbNewLine
                        
                        If InStr(1&, Result, NRyphm, vbBinaryCompare) = 0& Then
                                ResultB = ResultB + 1&
                                Result = Result & NRyphm
                                If ResultB > ResultMaxB Then
                                    Exit For
                                End If
                        End If
                    Next TLNG
                End If
            If Not ResultB > ResultMaxB Then
                SeekCount = fc_RyphmWord(IWord, 2&, SeekResults, True, 6000)
                    If Not SeekCount = -1& Then
                        For TLNG = 0& To 64&
                            NRyphm = IWord & " — " & SeekResults(RNDINT(SeekCount)) & vbNewLine
                            
                            If InStr(1&, Result, NRyphm, vbBinaryCompare) = 0& Then
                                ResultB = ResultB + 1&
                                Result = Result & NRyphm
                                If ResultB > ResultMaxB Then
                                    Exit For
                                End If
                            End If
                        Next TLNG
                    End If
            End If
        Else 'получить рифмующиеся к слову фразы
            ''Get all available ryphms by 3 sym
            SeekCount = fc_RyphmLine(IWord, 3&, SeekResultsL(), 31&)
                If Not SeekCount = -1& Then
                        For TLNG = 0& To SeekCount
                            NRyphm = SWArr.MArr(SeekResultsL(TLNG)) & vbNewLine
                            If InStr(1&, Result, NRyphm, vbBinaryCompare) = 0& Then
                                ResultB = ResultB + 1&
                                Result = Result & NRyphm
                                If ResultB > ResultMaxB Then
                                    Exit For
                                End If
                            End If
                        Next TLNG
                End If
            ''Then try with 2 if 1st attempt failed
            SeekCount = fc_RyphmLine(IWord, 2&, SeekResultsL(), 31&)
                If Not SeekCount = -1& Then
                        For TLNG = 0& To SeekCount
                            NRyphm = SWArr.MArr(SeekResultsL(TLNG)) & vbNewLine
                            If InStr(1&, Result, NRyphm, vbBinaryCompare) = 0& Then
                                ResultB = ResultB + 1&
                                Result = Result & NRyphm
                                If ResultB > ResultMaxB Then
                                    Exit For
                                End If
                            End If
                        Next TLNG
                End If
            
        End If

        'show results
        If Not Len(Result) = 0& Then
            MainOut Result, vbLeftJustify
        Else
            MainOut "Очень жаль, но не нашлось СЛОВА, брат.", vbLeftJustify
        End If
        MainOut IWord & vbNewLine & String(IIf(Len(IWord) <= 10, Len(IWord), 10), "—") & vbNewLine & txtMain.Text, vbLeftJustify
    End If
End Sub

Private Sub mnuSave_Click()
Dim FName As String
Dim FText As String
    On Local Error Resume Next
    FName = App.Path & "\+out.txt"
    FText = "# " & Right$("0000" & Year(Date), 4&) & "/" & Right$("00" & Month(Date), 2&) & "/" & Right$("00" & Day(Now), 2) & " " & Right$("00" & Hour(Time), 2) & ":" & Right$("00" & Minute(Time), 2) & ":" & Right$("00" & Second(Time), 2)
    FText = String$(16&, "—") & FText & " #" & String$(16&, "—") & vbNewLine & txtMain.Text & vbNewLine & vbNewLine
    Open FName For Binary As 1
        Put 1, LOF(1) + 1&, FText
    Close 1
    If Not Err.Number = 0 Then
        MainOut "Ошибка сохранения: " & vbNewLine & Err.Description, vbCenter
    End If
End Sub

Private Sub mnuTrayExit_Click()
    Unload frmMain
End Sub


Private Sub mnuTrayGenT_Click(Index As Integer)
    Call SHIT_Autogen(Val(mnuTrayGenT(Index).Tag), vbLeftButton)
End Sub

Private Sub mnuTrayShow_Click()
    MainTray.Remove
    frmMain.Visible = True
End Sub

Private Sub opt_Click(Index As Integer)
    Call zcs_UpdateTempCfg
End Sub


Private Sub cmdS_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode >= vbKey1 And KeyCode <= vbKey9 Then
        Select Case KeyCode
          Case vbKey1: Index = 1
          Case vbKey2: Index = 5
          Case vbKey3: Index = 10
          Case vbKey4: Index = 15
          Case vbKey5: Index = 20
          Case vbKey6: Index = 28
          Case vbKey7: Index = 44
          Case vbKey8: Index = 60
          Case vbKey9: Index = 124
        End Select
            Call SHIT_Autogen(Index, IIf(Shift = 0, vbLeftButton, vbRightButton))
                cmdS(Index).SetFocus
    End If
End Sub

Private Sub cmdS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Or X > cmdS(Index).Width Then Exit Sub
If Y < 0 Or Y > cmdS(Index).Height Then Exit Sub
    Call SHIT_Autogen(Index, Button)
End Sub


Private Sub txtCmd_Change()
    cmdC(2).Enabled = Not Len(txtCmd) = 0&
    cmdC(0).Enabled = cmdC(2).Enabled
    cmdC(1).Enabled = cmdC(2).Enabled
End Sub

Private Sub txtCmd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdC_Click(Val(txtCmd.Tag))
    End If
End Sub

Private Sub txtInLine_GotFocus()
    txtInLine.Width = txtMain.Width
    opt(2).Visible = False
    opt(8).Visible = False
    opt(6).Visible = False
    opt(10).Visible = False
    opt(5).Visible = False
    opt(14).Visible = False
    opt(16).Visible = False
End Sub

Private Sub txtInLine_LostFocus()
    txtInLine.Width = 645
    opt(2).Visible = True
    opt(8).Visible = True
    opt(6).Visible = True
    opt(10).Visible = True
    opt(5).Visible = True
    opt(14).Visible = True
    opt(16).Visible = True
End Sub

Private Sub txtMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Not Data.GetFormat(15) Then Exit Sub
If Not SW_OpState = 0 Then Exit Sub
If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)
SW_OpState = 2 'set operation state


Dim TFile&
Dim OFile$
Dim iFile$

Dim lPosA&
Dim lPosR&
Dim tChrB As Byte

Dim Tstr$

    If Data.Files.Count = 0 Then Exit Sub
    If Not MsgBox("Импортировать полученные файлы?" & " (" & Data.Files.Count & ")", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then Exit Sub

    For TFile = 1& To Data.Files.Count
        MainOut "Импорт файла " & TFile & " / " & Data.Files.Count & "...", vbCenter
        'in file
        iFile = Data.Files(TFile)
        
        'resulting file
        OFile = LCase$(fc_GetNameFromPath(iFile))
        OFile = App.Path & "\source_" & OFile
            If Not StrComp(Right$(OFile, 4&), ".txt", vbTextCompare) = 0& Then
                OFile = OFile & ".txt"
            End If
                
        If Not PathFileExists(iFile) = 0& Then
        If PathFileExists(OFile) = 0& Then
            Err.Clear
            'read source file
            Open iFile For Binary Access Read As #1
                If Not LOF(1) > 10485760 Then
                    Tstr = Space$(LOF(1&))
                Else
                    Tstr = Space$(10485760)
                End If
                    Get #1, 1&, Tstr
            Close #1
            
            If Err.Number = 0 Then
                lPosR = 1&
                For lPosA = 1& To Len(Tstr)
                    tChrB = Asc(Mid$(Tstr, lPosA, 1&))
    
                    If lPosA - lPosR >= 20 + (RNDINT(10)) Then
                        If tChrB = 32 Or tChrB = 10 Or tChrB = 13 Then
                            Mid$(Tstr, lPosA, 1&) = Chr$(1&)
                                lPosR = lPosA
                        End If
                    ElseIf lPosA - lPosR >= 3 Then
                        If tChrB = 10 Or tChrB = 13 Then
                            Mid$(Tstr, lPosA, 1&) = Chr$(1&)
                                lPosR = lPosA
                        End If
                    End If
'                    DoEvents
                Next lPosA

                Tstr = Replace$(Tstr, Chr$(10&), vbNullString, 1&, -1&, vbBinaryCompare)
                Tstr = Replace$(Tstr, Chr$(13&), vbNullString, 1&, -1&, vbBinaryCompare)
                Tstr = Replace$(Tstr, Chr$(1&), vbNewLine, 1&, -1&, vbBinaryCompare)
                Tstr = UCase$(Tstr)
                
                Dim FixRet&
                
                DoEvents

                FixRet = SHIT_CLEANUP_sbFilter(Tstr)

                Tstr = Replace$(Tstr, Chr$(10&), vbNewLine, 1&, -1&, vbBinaryCompare)
                
                'resulting file write
                    Open OFile For Binary Access Write As #1
                        Put #1, 1&, Tstr
                    Close #1
            End If
        End If
        End If
    Next TFile
    MainOut "Импорт файлов завершен", vbCenter
    SW_OpState = 0
End Sub

Private Sub txtOutLine_GotFocus()
    txtOutLine.Width = txtMain.Width
    opt(3).Visible = False
    opt(9).Visible = False
    opt(11).Visible = False
    opt(13).Visible = False
    opt(12).Visible = False
    opt(15).Visible = False
    opt(17).Visible = False
End Sub

Private Sub txtOutLine_LostFocus()
    txtOutLine.Width = 645
    opt(3).Visible = True
    opt(9).Visible = True
    opt(11).Visible = True
    opt(13).Visible = True
    opt(12).Visible = True
    opt(15).Visible = True
    opt(17).Visible = True
End Sub
Private Sub txtMain_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
    If KeyAscii = 3 Then 'Ctrl+C
        If Not txtMain.SelLength = 0& Then
            Call zcs_SetClipboard(txtMain.SelText)
'        Else
'            Call zcs_SetClipboard(txtMain.Text)
        End If
    ElseIf KeyAscii = 1 Then 'Ctrl+A
        txtMain.SelStart = 0
        txtMain.SelLength = Len(txtMain.Text)
    End If
    KeyAscii = 0
End Sub



Private Sub TmrAUX_Timer()
    Dim FullRandom As Boolean

    'реализация задержки с использованием GetTickCount. State = True когда срабатывает задержка.
        Dim Tick As Long: Tick = GetTickCount()
        Dim State As Boolean
        Static Trigger As Long
        Const Delay As Long = 120

            If Tick >= 0 Then
                If Not (Trigger < 0) Then
                    If Tick >= Trigger Then
                        Trigger = Tick + Delay
                        State = True
                    End If
                Else 'switch Trigger to +, w/o state set
                    Trigger = Tick + Delay
                End If
            Else
                If Trigger < 0 Then
                    If Tick >= Trigger Then
                        Trigger = Tick + Delay
                        State = True
                    End If
                Else 'switch Trigger to -, w/o state set
                    Trigger = Tick + Delay
                End If
            End If
    
    If frmMain.Visible Then
        LifeSub
    End If

    If Not State Then Exit Sub
    
    FullRandom = opt(11).Value = vbChecked
    opt(4).Enabled = Not opt(12).Value = vbChecked
    opt(2).Enabled = Not FullRandom
    opt(3).Enabled = Not FullRandom
    opt(5).Enabled = Not FullRandom And Not opt(12).Value = vbChecked
    opt(6).Enabled = Not FullRandom
    opt(8).Enabled = Not FullRandom
    opt(10).Enabled = Not FullRandom
    opt(12).Enabled = Not FullRandom
    opt(13).Enabled = Not FullRandom
    opt(14).Enabled = Not FullRandom And Not opt(12).Value = vbChecked
    opt(9).Enabled = opt(8).Value = vbChecked And Not FullRandom
    opt(15).Enabled = Not FullRandom And Not opt(3).Value = vbChecked
    opt(16).Enabled = Not FullRandom
    
    opt(1).Enabled = opt(0).Value = vbChecked

    If Rnd <= 0.1 Then
        opt(12).ToolTipText = RTrim$(nName(2& + RNDINT(1), 1 + RNDINT(1), True) & "-" & nName(2 + RNDINT(5), 1 + RNDINT(1), Rnd <= 0.5) & fc_RandomSign(True, True, 100))
    End If
    
    If Not AutoGen = 0 And Not KeyPressed(vbKeyControl) Then
        Call SHIT(AutoGen)
    End If
    
    If Rnd <= 0.03 Then
        frmMain.Caption = zcf_nStr(frmMain.LinkTopic)
    End If
End Sub



'процедура генерации текста
Public Sub SHIT(ByVal Cnt As Long, Optional ByRef USEED As Long)
On Error Resume Next
If Not SW_OpState = 0 Then Exit Sub

Dim bSeedMode As Boolean
    bSeedMode = Not USEED = 0&

If bSeedMode Then
    opFullRandom = True 'set some settings to allow proper generation by seed
    opCache = False
    opTags = False
    opChaos = False
    
    Call Rnd(-1) 'reset generator seed
        Randomize (USEED) 'init

    Cnt = 1& + RNDINT(31) 'set rnd text size
End If

'Randomize options
If opFullRandom Then
    opPunctuation = Rnd <= 0.5
    opLowerCase = Rnd <= 0.8
    opRandomCase = Rnd <= 0.5
    opMistakes = Rnd <= 0.5
    opOmsk = Rnd <= 0.15
    opEmptyLines = Rnd <= 0.5
    If Not bSeedMode Then
'        opTags = Rnd <= 0.5
        opChaos = Rnd <= 0.15
    End If

    opLyrics = Rnd <= 0.3

    opBaar = Rnd <= 0.15
    opRewrite = Rnd <= 0.15
End If

If SWCnt = -1& And Not opChaos Then Call MainOut("Я ОБКАКАЛСЯ. ПОЧЕМУ?", vbCenter): Exit Sub

Dim wCnt& 'Result strings count
Dim TLNG&

Dim rArr$()
Dim RCnt As Long

Dim rStr$

Dim CSign$

Dim CapitalizeNext As Boolean

wCnt = Cnt + RNDINT(4)

    If Not opChaos Then
        If opLyrics Then ' Additional strings to recover broken lyrics strings
             wCnt = wCnt + RNDINT((Cnt / 10))
        End If
        
        If opCache Then  ' Cache fix
            If (SWCnt - MaxUsed) <= wCnt Then
                Mid$(sUsed, 1&, MaxUsed * 10&) = Space$(MaxUsed * 10&)
            End If
        End If
    End If

'Part 1, form raw strings array
    ReDim rArr(1& To wCnt) As String
    RCnt = SHIT_Part1(rArr, wCnt)

    'empty result, try again
        If RCnt = 0& Then
            Call SHIT(ByVal Cnt)
            If Err.Number = 28& Then
                MainOut "Не взлетело. Слишком мало строк?", vbCenter
            End If
                Exit Sub
        ElseIf Not UBound(rArr) = RCnt Then
            ReDim Preserve rArr(1& To RCnt) As String
        End If
'- - - - - - - - - '

'Here we have rArr() array of strings, aka result
'Further code is used to parse it
'- - - - - - - - - '
    
    
'Part 2, lyrics code
    If opLyrics And Not opChaos Then
        Call SHIT_Part2(RCnt, rArr())
    End If


'Part 2.5, Rewrite
If opRewrite And Not opChaos Then
    Call sb_Rewrite(rArr, RCnt, opLyrics)
End If

'Part 2.6, Mistakes
If opMistakes Then
    Call sb_Mistakes(rArr, RCnt, opLyrics)
End If

CSign = " "
For TLNG = 1& To RCnt
    'lower case
    If opLowerCase Or opRandomCase Then
        If Not Len(rArr(TLNG)) = 0& Then
            If opLowerCase Then
                zcs_SwapStrings rArr(TLNG), LCase$(rArr(TLNG))
            Else
                sb_NyaFilter rArr(TLNG)
            End If

            If TLNG = 1& Or opLyrics Then
                Mid$(rArr(TLNG), 1&, 1&) = UCase$(Mid$(rArr(TLNG), 1&, 1&))
            ElseIf opPunctuation Or CapitalizeNext Then
                If Right$(CSign, 2&) = ". " Or Right$(CSign, 2&) = "? " Or Right$(CSign, 2&) = "! " Then
                    Mid$(rArr(TLNG), 1&, 1&) = UCase$(Mid$(rArr(TLNG), 1&, 1&))
                End If
                    CapitalizeNext = False
            End If
        End If
    End If
    
    'Punctuation (random char, also used for newlines, etc)
    If Not Len(fc_GetLastWord(rArr(TLNG), True, False, 0&)) <= 2& Then
        CSign = fc_RandomSign(True, False, 64)
    Else
        CSign = " "
    End If
    
'    'Punctuation inside every string
'    If opPunctuation Then
'        Call sb_RandomSigns(rArr(Tlng), 0.0066, 0.5)
'    End If
    
    'Punctuation
    If opPunctuation Then
        If Not TLNG = RCnt Then
            If opLyrics Then 'Lyrics outside punctuation fix
                If Not TLNG Mod 2 = 0& Then
                    If Right$(CSign, 2&) = ". " Or Right$(CSign, 2&) = "? " Or Right$(CSign, 2&) = "! " Then
                        CSign = " "
                    End If
                End If
            End If
            rArr(TLNG) = rArr(TLNG) & CSign 'outside punctuation
        Else
        'Punctuation post-fix
            CSign = fc_RandomSign(True, True, 100)
            rArr(TLNG) = RTrim$(rArr(TLNG) & CSign)
        End If
    ElseIf Not TLNG = RCnt Then
        rArr(TLNG) = rArr(TLNG) & " "
    End If
    'Tags
    If opTags Then
        If Rnd <= 0.1 Then
            rArr(TLNG) = RTrim$(rArr(TLNG))
                Call sb_RandomTag(rArr(TLNG))
            If Not TLNG = RCnt Then rArr(TLNG) = rArr(TLNG) & " "
        End If
    End If
    'force lyrics 1st spaceholder
        If opLyrics Then
            If Not TLNG = RCnt Then rArr(TLNG) = RTrim$(rArr(TLNG)) & "\n"
        End If
    
    'New lines
    If opEmptyLines And Not TLNG = RCnt Then
        If opLyrics Then 'force lyrics second spaceholder
            If TLNG Mod 4 = 0 Then
                If Right$(CSign, 2&) = " " Or Right$(CSign, 2&) = ". " Or Right$(CSign, 2&) = "! " Or Right$(CSign, 2&) = "? " Then
                    If Rnd <= 0.8 Then rArr(TLNG) = rArr(TLNG) & "\n"
                End If
            End If
        ElseIf Rnd <= 0.22 Then 'New lines
            If Right$(CSign, 2&) = ". " Or Right$(CSign, 2&) = "! " Or Right$(CSign, 2&) = "? " Then
                'fix: capitalize 1st letter of next line
                If opLowerCase And Not TLNG + 1 > RCnt Then
                    CapitalizeNext = True
                End If
                
                If Rnd <= 0.42 Then
                    rArr(TLNG) = RTrim$(rArr(TLNG)) & "\n"
                Else
                    rArr(TLNG) = RTrim$(rArr(TLNG)) & "\n\n"
                End If
            End If
        End If
    End If

Next TLNG

'Array to a single string

rStr = Join(rArr, vbNullString)

'Apply Medved filter
If opBaar Then Call sb_MedvedFilterEx(rStr)

'letter-fix
If opLetterFix Then Call sb_CyrFilterEx(rStr)

'restore settings of seed mode
If bSeedMode Then
    zcs_UpdateTempCfg 'restore temp config
    
    Call Rnd(-1) 'reset generator seed to system time
        Randomize 'init
    
'    rStr = rStr & vbNewLine & vbNewLine & "Seed: " & txtCmd & "@x" & UBound(SWArr.MArr) + 1&
Else
    'join all together, don't do that in seed mode
        rStr = txtInLine & rStr & txtOutLine
End If

'set newline placeholders
rStr = Replace$(rStr, "\n", vbNewLine, 1&, -1&, vbBinaryCompare)



'show result
MainOut rStr, vbLeftJustify

If opCopyPaste Or Not frmMain.Visible Then Call zcs_SetClipboard(rStr)

End Sub
'Расставляет слова в строке случайным образом
'Если KeepLastWord, тогда последнее слово останется на своем месте
Private Function fc_StrWordsResort(ByRef ISTR As String, KeepLastWord As Boolean) As String
Dim iUsed As Long
Dim iCurrent As Long
Dim sRUsed As String
Dim Tarr() As String
Dim TLNG As Long

If Len(ISTR) = 0& Then Exit Function

Tarr = Split(ISTR, " ", -1&, vbBinaryCompare)


'omsk words splitter
Dim ATemp As Long
Dim TLast As Long
Dim TTemp As Long
Dim TLMax As Long
Dim TSCnt As Long
Dim SNew As String

For ATemp = 0& To UBound(Tarr)
    If Rnd <= 0.5 Then
        TLMax = Len(Tarr(ATemp))
        
        If TLMax >= 4& Then
         SNew = Tarr(ATemp)
         Tarr(ATemp) = vbNullString
             For TTemp = 1& To TLMax
                  Tarr(ATemp) = Tarr(ATemp) & Mid$(SNew, TTemp, 1&)
                  If TSCnt < 1 Then
                  If (TTemp - TLast >= 2&) And (TLMax - TTemp >= 2) Then
                    If Rnd <= 0.1 Then
                      Tarr(ATemp) = Tarr(ATemp) & "-"
                      TLast = TTemp
                        If Rnd <= 0.5 Then TSCnt = TSCnt + 1&
                    End If
                  End If
                  End If
             Next TTemp
         TLast = 0&
         TSCnt = 0&
        End If
    End If
Next ATemp
''''''''''''''



Dim MaxCycle As Long
    If KeepLastWord And Not UBound(Tarr) = 0& Then
        MaxCycle = UBound(Tarr) - 1&
    Else
        MaxCycle = UBound(Tarr)
    End If
    
Recycle:
iCurrent = RNDINT(MaxCycle)
If InStr(1&, sRUsed, vbNullChar & iCurrent & vbNullChar, vbBinaryCompare) = 0& Then
   sRUsed = sRUsed & vbNullChar & iCurrent & vbNullChar
   
   iUsed = 0&
   For TLNG = 0& To MaxCycle
    If Not InStr(1&, sRUsed, vbNullChar & TLNG & vbNullChar, vbBinaryCompare) = 0& Then
        iUsed = iUsed + 1
    End If
   Next TLNG

    fc_StrWordsResort = fc_StrWordsResort & Tarr(iCurrent) & " "

   If iUsed = (MaxCycle + 1) Then
        If KeepLastWord And Not MaxCycle = UBound(Tarr) Then
            fc_StrWordsResort = fc_StrWordsResort & Tarr(UBound(Tarr))
        End If

    fc_StrWordsResort = RTrim$(fc_StrWordsResort)
        Exit Function
   End If
   
End If
GoTo Recycle

End Function

'Очистка массива от сторонних символов и строк-дубликатов
Private Sub SHIT_CLEANUP()
On Error Resume Next

Dim GTFound&

Dim NText As String
Dim SymFixes As Long
                
If Not SW_OpState = 0 Then Exit Sub
    Call a_loadFile 'reload source file

If SWCnt = -1& Then Call MainOut(sConfig.sSourceName & " пуст. Нет пути форматировать.", vbCenter):  Exit Sub
SW_OpState = 2 'set operation state


'fix symbols
MainOut "Фикс символов...", vbCenter:      DoEvents
NText = Join(SWArr.MArr, Chr$(10&))

'fix
SymFixes = SHIT_CLEANUP_sbFilter(NText)

SWArr.MArr() = Split(NText, Chr$(10&), -1&, vbBinaryCompare)
MainOut "Фикс символов... OK", vbCenter:     DoEvents

'remove string duplicates
SWArr.MArr() = SHIT_CLEANUP_fcFixDoubleString(SWArr.MArr(), GTFound)
MainOut "Поиск дубликатов: OK", vbCenter:     DoEvents

Dim OutFile As String: OutFile = AppSPath & sConfig.sSourceName
Dim BckFile As String: BckFile = AppSPath & "\back_" & sConfig.sSourceName

If Not GTFound = 0& Or Not SymFixes = 0& Then
    MainOut "* Очистка завершена:", vbCenter
    
    If Not GTFound = 0& Then
        MainOut txtMain.Text & vbNewLine & " -" & GTFound & " " & zcf_GetStringsWord(GTFound) & ".", vbCenter
    End If
    If Not SymFixes = 0& Then
        MainOut txtMain.Text & vbNewLine & " -" & SymFixes & " символов.", vbCenter
    End If

    If MsgBox("Перезаписать " & sConfig.sSourceName & "?", vbInformation + vbYesNo) = vbYes Then
        'write file
        Err.Clear
            Call MoveToRecycleBin(OutFile, True)
                Open OutFile For Binary Access Write As 1
                    Put 1, 1, Join(SWArr.MArr, vbNewLine)
                Close 1
            If Err.Number = 0& Then
                MainOut txtMain.Text & vbNewLine & sConfig.sSourceName & " обновлен.", vbCenter
            Else
                MainOut txtMain.Text & vbNewLine & "! " & Err.Description, vbCenter
            End If
    End If
    LDate = vbNullString:   Call a_loadFile 'force reset file date and reload file
Else
    MainOut "* Очистка завершена. Ничего лишнего не найдено.", vbCenter
End If
SW_OpState = 0
End Sub
Public Function MoveToRecycleBin(FileSpec As String, Optional _
NoConfirm As Boolean = False) As Boolean

'returns true if succesful, false if not

'If NoConfirm is set to true,
'Windows Confirmation dialog is suppressed

'FileSpec can be a file name
'e.g., "C:\myfile.txt"
'or a directory/wildcard combination
'e.g., C:\*.txt

Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lFlags As Long

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION
With WinType_SFO
    .wFunc = FO_DELETE
    .pFrom = FileSpec
    .fFlags = lFlags
End With
lRet = SHFileOperation(WinType_SFO)
MoveToRecycleBin = (lRet = 0)
End Function


'Quickly removes duplicates of strings array
'By speechwriter@mail.ru.
Private Function SHIT_CLEANUP_fcFixDoubleString(ByRef iRawArr() As String, ByRef DoublesCount As Long) As Variant
''On Local Error Resume Next
Dim lTemp&
Dim sMatch$

Dim hRC As New clsC

Dim sRes As String: sRes = Space$(1024&) ': Mid$(sRes, 1&, 2&) = vbNewLine
Dim lLen As Long: lLen = 1&
Dim TLen As Long

Dim NArr() As String: ReDim NArr(LBound(iRawArr) To UBound(iRawArr)) As String
Dim NCnt As Long

    sMatch = Space$(10&)
    TLen = Len(sMatch)

If UBound(iRawArr) = -1& Then DoublesCount = -1&: Exit Function
    For lTemp = 0& To UBound(iRawArr)
        Mid$(sMatch, 1&, 1&) = "+"
        Mid$(sMatch, 2&, 8&) = Right$("00000000" & Hex$(hRC.CRC32(iRawArr(lTemp))), 8&)
        Mid$(sMatch, 10&, 1&) = "-"
            If InStr(1&, sRes, sMatch, vbBinaryCompare) = 0& Then
                If lLen + TLen > Len(sRes) Then
                    sRes = sRes & Space$(500500)
                End If
                Mid$(sRes, lLen, TLen) = sMatch
                lLen = lLen + TLen
                
                    zcs_SwapStrings NArr(NCnt), iRawArr(lTemp)
                    NCnt = NCnt + 1&
                    
            Else
                DoublesCount = DoublesCount + 1&
            End If
            If lTemp Mod 1024& = 0& Or lTemp = UBound(iRawArr) Then
                'code to display state
                MainOut "Поиск дубликатов: " & lTemp + 1& & " из " & UBound(iRawArr) + 1&, vbCenter
                DoEvents
            End If
    Next lTemp

ReDim Preserve NArr(LBound(iRawArr) To NCnt - 1&) As String
    
SHIT_CLEANUP_fcFixDoubleString = NArr
Set hRC = Nothing
End Function

Private Sub zcs_SetClipboard(ByRef StringToCopy As String)
Dim hMem As Long, pMem As Long
    If Len(StringToCopy) = 0& Then Exit Sub
    If opCopyPasteUTF8 Then

        Call OpenClipboard(Me.hwnd)
        Call EmptyClipboard
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 2& + LenB(StringToCopy))
        pMem = GlobalLock(hMem)
        Call CopyMemory(ByVal pMem, ByVal StrPtr(StringToCopy), ByVal LenB(StringToCopy))
        Call GlobalUnlock(hMem)
        Call SetClipboardData(CF_UNICODETEXT, hMem)
        Call CloseClipboard
    Else
        Clipboard.Clear
        Clipboard.SetText StringToCopy
    End If
End Sub
'Возвращает случайную строку массива у которой последнее слово рифмуется с iRyphm
'iLen - количество символов от конца строки, по которым производится сравнение
'iRyphm - строка, к которой ищется рифма. Из нее будет выделено последнее слово.
'Результат - количество найденных строк в массиве или -1, если не удалось найти ни 1
'В ArrOfMatches записываются индексы найденных строк, он должен быть заранее сформирован

Private Function fc_RyphmLine(ByRef iRyphm As String, ByRef iLen As Long, ByRef ArrOfMatches() As Long, ByRef MaxMatches As Long) As Long
Dim RCnt As Long
Dim TCnt As Long

Dim OldRyphm$
    OldRyphm = fc_GetLastWord(iRyphm, True, True, 0&)
Dim OldRyphmPart$
    OldRyphmPart = Right$(OldRyphm, iLen)
    
Dim NewRyphm$
Dim NewRyphmPart$
    NewRyphmPart = Space$(iLen)

Dim OKRyphm As Boolean

RCnt = -1&

Dim TLNG As Long
For TLNG = 1& To 1024&
TCnt = RNDINT(SWCnt)
'For TCnt = 0& To SWCnt
    OKRyphm = False
    If Not Len(SWArr.MArrRyphm(TCnt)) - iLen + 1& <= 0& Then
    If StrComp(OldRyphmPart, Mid$(SWArr.MArrRyphm(TCnt), Len(SWArr.MArrRyphm(TCnt)) - iLen + 1&, iLen), vbBinaryCompare) = 0& Then
        zcs_SwapStrings NewRyphm, fc_GetLastWord(SWArr.MArrRyphm(TCnt), True, False, Len(OldRyphm))
        OKRyphm = InStr(1&, OldRyphm, NewRyphm, vbBinaryCompare) = 0&
    End If
    
    'Add +1 to results
    If OKRyphm Then
        RCnt = RCnt + 1&
        ArrOfMatches(RCnt) = TCnt
        'Cap max results
            If RCnt = MaxMatches& Then Exit For
    End If
    End If
'Next TCnt
Next TLNG

fc_RyphmLine = RCnt
End Function


'Возвращает случайное слово из массива, которое рифмуется с iRyphm
'iLen - количество символов от конца строки, по которым производится сравнение
'iRyphm - строка, к которой ищется рифма. Из нее будет выделено последнее слово.
'Результат - найденное слово или пустая строка
'В ArrOfMatches записываются найденные слова
Private Function fc_RyphmWord(ByRef iRyphm As String, ByRef iLen As Long, ByRef ArrOfMatches() As String, ByVal bUnique As Boolean, Optional RandomSeek As Long) As Long
Dim RCnt As Long
Dim TCnt As Long
Dim RndC As Long

Dim OldRyphm$
    OldRyphm = fc_GetLastWord(iRyphm, True, True, 0&)
Dim OldRyphmPart$
    OldRyphmPart = Right$(OldRyphm, iLen)
    
Dim NewRyphm$

Dim OKRyphm As Boolean
Dim OKRyphmL As Long
Dim WordDispos As Long
Dim Words$()

Erase ArrOfMatches
RCnt = -1&

If Not bUnique Then OKRyphm = True
For RndC = 0& To SWCnt
    If Not RandomSeek = 0& Then
        RandomSeek = RandomSeek - 1&
            If RandomSeek = 0 Then
                Exit For
            End If
        TCnt = RNDINT(SWCnt)
    Else
        TCnt = RndC
    End If
    
    If Not Len(SWArr.MArrRyphm(TCnt)) - iLen + 1& <= 0& Then
    If StrComp(OldRyphmPart, Mid$(SWArr.MArrRyphm(TCnt), Len(SWArr.MArrRyphm(TCnt)) - iLen + 1&, iLen), vbBinaryCompare) = 0& Then
        zcs_SwapStrings NewRyphm, fc_GetLastWord(SWArr.MArrRyphm(TCnt), True, False, Len(OldRyphm))
        
        If bUnique Then
            OKRyphm = InStr(1&, OldRyphm, NewRyphm, vbBinaryCompare) = 0&
        End If
            'Add +1 to results
            If OKRyphm Then
                RCnt = RCnt + 1&
                    ReDim Preserve ArrOfMatches(0& To RCnt) As String
                        ArrOfMatches(RCnt) = fc_GetLastWord(SWArr.MArr(TCnt), True, False, 0&)
            End If
    
    Else
        OKRyphmL = InStr(1&, SWArr.MArrRyphm(TCnt), OldRyphmPart & " ", vbBinaryCompare)
            If Not OKRyphmL = 0& Then
                Words = Split(SWArr.MArr(TCnt), " ", -1&, vbBinaryCompare)
                
                For WordDispos = 0& To UBound(Words)
                    NewRyphm = Right$(Words(WordDispos), iLen)
                    sb_MedvedFilter NewRyphm
                    
                    If StrComp(OldRyphmPart, NewRyphm) = 0& Then
                        NewRyphm = Right$(Words(WordDispos), Len(OldRyphm))
                        sb_MedvedFilter NewRyphm
                        If bUnique Then
                            OKRyphm = InStr(1&, OldRyphm, NewRyphm, vbBinaryCompare) = 0&
                        End If
                        
                        'Add +1 to results
                        If OKRyphm Then
                            RCnt = RCnt + 1&
                                ReDim Preserve ArrOfMatches(0& To RCnt) As String
                                    ArrOfMatches(RCnt) = Words(WordDispos)
                        End If
                    End If
                Next WordDispos

            End If
    End If
    End If
Next RndC

fc_RyphmWord = RCnt
End Function


'Temrorary disabled
'Like a fc_RyphmWord, but uses exact word matches
'Private Function fc_RyphmWordEx(ByRef iRyphm As String, ByRef iLen As Long, ByRef ArrOfMatches() As String, ByVal bUnique As Boolean) As Long
'Dim RCnt As Long
'Dim TCnt As Long
'
'Dim OldRyphm$
'    OldRyphm = fc_GetLastWord(iRyphm, True, False, 0&)
'Dim OldRyphmPart$
'    OldRyphmPart = Right$(OldRyphm, iLen)
'
'Dim NewRyphm$
'
'Dim OKRyphm As Boolean
'Dim OKRyphmL As Long
'Dim WordDispos As Long
'Dim Words$()
'
'Erase ArrOfMatches
'RCnt = -1&
'
'If Not bUnique Then OKRyphm = True
'For TCnt = 0& To SWCnt
'
'    If Not Len(SWArr.MArr(TCnt)) - iLen + 1& <= 0& Then
'    If StrComp(OldRyphmPart, Mid$(SWArr.MArr(TCnt), Len(SWArr.MArr(TCnt)) - iLen + 1&, iLen), vbBinaryCompare) = 0& Then
'        zcs_SwapStrings NewRyphm, fc_GetLastWord(SWArr.MArr(TCnt), True, False, Len(OldRyphm))
'
'        If bUnique Then
'            OKRyphm = InStr(1&, OldRyphm, NewRyphm, vbBinaryCompare) = 0&
'        End If
'            'Add +1 to results
'            If OKRyphm Then
'                RCnt = RCnt + 1&
'                    ReDim Preserve ArrOfMatches(0& To RCnt) As String
'                        ArrOfMatches(RCnt) = fc_GetLastWord(SWArr.MArr(TCnt), True, False, 0&)
'            End If
'
'    Else
'        OKRyphmL = InStr(1&, SWArr.MArr(TCnt), OldRyphmPart & " ", vbBinaryCompare)
'            If Not OKRyphmL = 0& Then
'                Words = Split(SWArr.MArr(TCnt), " ", -1&, vbBinaryCompare)
'
'                For WordDispos = 0& To UBound(Words)
'                    NewRyphm = Right$(Words(WordDispos), iLen)
'                    sb_MedvedFilter NewRyphm
'
'                    If StrComp(OldRyphmPart, NewRyphm) = 0& Then
'                        NewRyphm = Right$(Words(WordDispos), Len(OldRyphm))
'                        sb_MedvedFilter NewRyphm
'                        If bUnique Then
'                            OKRyphm = InStr(1&, OldRyphm, NewRyphm, vbBinaryCompare) = 0&
'                        End If
'
'                        'Add +1 to results
'                        If OKRyphm Then
'                            RCnt = RCnt + 1&
'                                ReDim Preserve ArrOfMatches(0& To RCnt) As String
'                                    ArrOfMatches(RCnt) = Words(WordDispos)
'                        End If
'                    End If
'                Next WordDispos
'
'            End If
'    End If
'    End If
'Next TCnt
'
'fc_RyphmWordEx = RCnt
'End Function
Private Function fc_RandomSign(ByRef outroSigns As Boolean, ByRef forceOutro As Boolean, ByVal MaxChance As Byte) As String
If outroSigns And forceOutro Then GoTo jmpOutro

    Select Case CByte(RNDINT(MaxChance))
        Case Is <= 23: '30%
            If Rnd <= 0.86 Then
                fc_RandomSign = ", "
            Else
                fc_RandomSign = " — "
            End If
            
'        Case Is <= 24: If Rnd <= 0.88 Then fc_RandomSign = ": " Else fc_RandomSign = "; " '2x 2%

        Case Is <= 44 And outroSigns = True:  '16%
jmpOutro:
            If Rnd <= 0.6 Then
                fc_RandomSign = ". "
            ElseIf Rnd <= 0.75 Then
                fc_RandomSign = String$(2& + RNDINT(4&), ".") & " " '...
            ElseIf Rnd <= 0.88 Then
                If Rnd <= 0.5 Then
                    fc_RandomSign = String$(1& + RNDINT(2&), "?") & " " '???
                Else
                    fc_RandomSign = String$(1& + RNDINT(4&), "!") & " " '!!!
                End If
            Else
                If Rnd <= 0.5 Then
                    fc_RandomSign = String$(1& + RNDINT(2&), "?") & String$(1& + RNDINT(2&), "!") & " " '??!!'
                Else
                    fc_RandomSign = String$(1& + RNDINT(2&), "!") & String$(1& + RNDINT(2&), "?") & " "   '!!??
                End If
            End If
        Case Else: fc_RandomSign = " "
    End Select
End Function

'Cache
Private Function fc_Cache(ByRef iNewItem As Long) As Boolean
  Dim CacheItem As String
      If iNewItem < 0& Then Exit Function
      
      CacheItem = vbNullChar & Right$("00000000" & iNewItem, 8&) & vbNullChar
    If InStr(1&, sUsed, CacheItem, vbBinaryCompare) = 0& Then
       sCnt = sCnt + 1&
       If sCnt > MaxUsed Then sCnt = 1& 'Reset cache position
          Mid$(sUsed, (sCnt * 10) - 9&, 10&) = CacheItem
       fc_Cache = True
    Else
       fc_Cache = False
    End If
    
End Function
'заключает строку в случайный тег
Private Sub sb_RandomTag(ByRef rStr As String)
    Dim TagR As Byte: TagR = RNDINT(3)
    Dim Tags(0 To 7) As String
        If Not opTagsDoll Then
            '2ch.hk tags
            Tags(0) = "[b]"
            Tags(4) = "[/b]"
        
            Tags(1) = "[spoiler]"
            Tags(5) = "[/spoiler]"
            
            Tags(2) = "[u]"
            Tags(6) = "[/u]"
            
            Tags(3) = "[s]"
            Tags(7) = "[/s]"
        Else 'dollchan tags
            Tags(0) = "**"
            Tags(4) = "**"
        
            Tags(1) = "%%"
            Tags(5) = "%%"
            
            Tags(2) = "__"
            Tags(6) = "__"
            
            Tags(3) = "*"
    '        sTagO = "^H" 'example: test^H^H^H^H значит что test будет зачеркнуто, по 1 ^H на каждый символ
            Tags(7) = "*" 'Мне лень с этим ебаться, так что просто замена зачеркивания на наклонный
        End If
    rStr = Tags(TagR) & rStr & Tags(TagR + 4)
End Sub

'Заменяет пробелы на случайные символы

Private Sub sb_RandomSigns(ByRef rStr As String, ByRef chanseQuotes As Double, ByRef chanceSign As Double)
Dim Tarr$()
Dim TLNG&
Dim CSign$
Dim SLimit As Byte
Dim SStart As Long
Dim SEndng As Long

Dim QLimit As Byte


    Tarr = Split(rStr, " ", -1&, vbBinaryCompare)
    SLimit = CLng(1& + Len(rStr) \ 32)
    SStart = UBound(Tarr) * 0.3
    SEndng = UBound(Tarr) * 0.6
    
    QLimit = SLimit
    
    For TLNG = 0& To UBound(Tarr)
'    For Tlng = SStart To SEndng
        'add quotes
        If Rnd <= chanseQuotes And Not QLimit = 0 And Len(Tarr(TLNG)) > 2& Then
            sb_RandomQuotes Tarr(TLNG)
            QLimit = QLimit - 1
        End If

        'generate sign
        If Rnd <= chanceSign And Not TLNG = UBound(Tarr) And Not SLimit = 0 And Len(Tarr(TLNG)) > 2& And TLNG >= SStart And TLNG <= SEndng Then
            CSign = fc_RandomSign(False, False, 100)
            SLimit = SLimit - 1
        Else
            CSign = " "
        End If
        
        'add sign
        If Not TLNG = UBound(Tarr) Then
            Tarr(TLNG) = Tarr(TLNG) & CSign
        End If
        
    Next TLNG
        rStr = Join(Tarr, vbNullString)


End Sub


'Returns last word of the string
'If IncludeDefice then words are divided by "-" and " ", else only by " " symbol
Private Function fc_GetLastWord(ByRef rStr As String, ByRef IncludeDefice As Boolean, ByRef ApplyMedvedFilter As Boolean, ByRef MaxLen As Long) As String
    Dim TLNG&: TLNG = Len(rStr)
    Dim TByte As Byte
    
    
    For TLNG = Len(rStr) To 1& Step -1&
        TByte = Asc(Mid$(rStr, TLNG, 1&))
            If TByte = 45 And IncludeDefice Then
                Exit For
            ElseIf TByte = 32 Then
                Exit For
            End If
    Next TLNG
    
    TLNG = TLNG + 1&
        
    If Not MaxLen = 0& Then
        If Len(rStr) - (TLNG - 1&) > MaxLen Then
            TLNG = (Len(rStr) - MaxLen) + 1&
        End If
    End If
    
    fc_GetLastWord = Mid$(rStr, TLNG)
    
    If ApplyMedvedFilter Then Call sb_MedvedFilter(fc_GetLastWord)

End Function

Private Function SHIT_CLEANUP_sbFilter(ByRef rStr As String) As Long
  'фильтр левЫх символов
  'Диапазон допустимых символов, > - знак замены
  '10
  '13
  '45
  '48-57
  '32
  '65-90
  '97-122
  '150-151 > 45
  '160 > 32
  '165
  '168
  '170
  '173 > 45
  '175
  '178-180
  '184
  '186
  '191-255
  '168 > 197  Ё>Е
  Dim CHR00 As String * 1, CHR32 As String * 1
  Dim TLNG&, TMAX&
  Dim CCount As Long
  Dim TByte As Byte
  Dim tByteI As Byte
  Dim tByteO As Byte
    CHR00 = Chr$(0&): CHR32 = Chr$(32&)
    TMAX = Len(rStr)

XCycle:
    For TLNG = 1& To TMAX
        If TLNG Mod 500500 = 0 Then DoEvents
      TByte = Asc(Mid$(rStr, TLNG, 1&))

      If TByte = 0 _
        Or (TByte >= 65 And TByte <= 90) Or (TByte >= 97 And TByte <= 122) _
        Or TByte = 165 Or TByte = 170 Or TByte = 175 _
        Or (TByte >= 178 And TByte <= 180) Or TByte = 184 Or TByte = 186 _
        Or (TByte >= 191 And TByte <= 255) Then
            'leave it as is
      Else
        'get bytes at position +1 and -1
        If Not TLNG = TMAX Then
            tByteO = Asc(Mid$(rStr, TLNG + 1&, 1&))
                If tByteO = 0 Then tByteO = 255
        Else
            tByteO = 0
        End If
        If Not TLNG = 1& Then
            tByteI = Asc(Mid$(rStr, TLNG - 1&, 1&))
                If tByteI = 0 Then tByteI = 255
        Else
            tByteI = 0
        End If
        
        If (TByte >= 48 And TByte <= 57) Then
            If tByteO = 10 Or tByteO = 13 Or tByteO = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
        ElseIf TByte = 10 Or TByte = 13 Then
            If TByte = 13 Then
                Mid$(rStr, TLNG, 1&) = Chr$(10&)
                    CCount = CCount + 1&
            End If

            If tByteO = 10 Or tByteO = 13 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            If tByteO = 32 Then
                Mid$(rStr, TLNG + 1, 1&) = CHR00
                    CCount = CCount + 1&
            End If
        
        
            If tByteI = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            If tByteO = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If

            
            If tByteI = 45 Then
                Mid$(rStr, TLNG - 1&, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            If tByteO = 45 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
        
        ElseIf TByte = 32 Then
            If tByteO = 45 Or tByteO = 173 Then
                Mid$(rStr, TLNG + 1&, 1&) = CHR32
            End If
            If tByteI = 45 Then
                Mid$(rStr, TLNG - 1&, 1&) = CHR00
                    CCount = CCount + 1&
            End If
        
            If tByteO = 32 Or tByteO = 160 Or tByteO = 13 Or tByteO = 10 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            If tByteI = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            
        ElseIf TByte = 45 Then
            If tByteI = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
            If tByteO = 45 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            End If
        ElseIf TByte = 168 Then
            Mid$(rStr, TLNG, 1&) = Chr$(197&)
                CCount = CCount + 1&
                
        ElseIf (TByte >= 150 And TByte <= 151) Then
            Mid$(rStr, TLNG, 1&) = Chr$(45&)
                CCount = CCount + 1&
        ElseIf TByte = 160 Then
            If tByteO = 32 Or tByteO = 160 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            ElseIf tByteO = 0 Then
                Mid$(rStr, TLNG, 1&) = CHR00
                    CCount = CCount + 1&
            Else
                Mid$(rStr, TLNG, 1&) = CHR32
                    CCount = CCount + 1&
            End If

        
        ElseIf TByte = 173 Then
            Mid$(rStr, TLNG, 1&) = Chr$(45&)
                CCount = CCount + 1&
        Else
            Mid$(rStr, TLNG, 1&) = CHR00
                CCount = CCount + 1&
        End If
      End If
    Next TLNG
    
  If Not CCount = 0& Then
    MainOut "Фикс символов (-" & CCount & ")", vbCenter
    DoEvents
    SHIT_CLEANUP_sbFilter = SHIT_CLEANUP_sbFilter + CCount
        CCount = 0&
            GoTo XCycle
  End If
  If Not SHIT_CLEANUP_sbFilter = 0& Then
        rStr = Replace$(rStr, CHR00, vbNullString, 1&, -1&, vbBinaryCompare)
  End If
End Function

'заключает строку в случайные кавычки
Private Sub sb_RandomQuotes(ByRef rStr As String)
'“ ”, ‘ ’, « », ‹ ›, " "

'If Rnd <= 0.03 Then
'    rStr = "“" & rStr & "”"
'ElseIf Rnd <= 0.05 Then
'    rStr = "‘" & rStr & "’"
'ElseIf Rnd <= 0.07 Then
'    rStr = "‹" & rStr & "›"
'Else

If Rnd <= 0.2 Then
    rStr = "«" & rStr & "»"
Else
    rStr = """" & rStr & """"
End If

End Sub

'Set temporary config
Private Sub zcs_UpdateTempCfg()
    
    'for online mode
    If opOnline Then

    Else
        opFullRandom = opt(11).Value = vbChecked
        
        If opFullRandom Then
    ' These options will be randomly defined in SHIT sub:
    '        opPunctuation
    '        opLowerCase
    '        opRandomCase
    '        opMistakes
    '        opOmsk
    '        opEmptyLines
    '        opLyrics
    '        opBaar
    '        opRewrite
        Else
            opRewrite = opt(14).Value = vbChecked
            opChaos = opt(12).Value = vbChecked
            opPunctuation = opt(2).Value = vbChecked
            opLowerCase = opt(3).Value = vbChecked
            opRandomCase = opt(15).Value = vbChecked
            opMistakes = opt(16).Value = vbChecked
            opOmsk = opt(5).Value = vbChecked
            opEmptyLines = opt(6).Value = vbChecked
    
            opLyrics = opt(10).Value = vbChecked
    
            opBaar = opt(13).Value = vbChecked
        End If
            'strategical options, always static
            opTags = opt(8).Value = vbChecked
            opCopyPaste = opt(0).Value = vbChecked
            opCopyPasteUTF8 = opt(1).Value = vbChecked
            opCache = opt(4).Value = vbChecked
            opLetterFix = opt(7).Value = vbChecked
            opTagsDoll = opt(9).Value = vbChecked
    End If
End Sub
Public Sub zcs_SwapStrings(ByRef Str1 As String, ByRef Str2 As String)
    Dim lpStr1 As Long, lP1 As Long
    Dim lpStr2 As Long, lP2 As Long
    
        'get real pointers
        lpStr1 = VarPtr(Str1)
        lpStr2 = VarPtr(Str2)
        
        'get vb6 pointers
        CopyMemory lP1, ByVal lpStr1, 4&
        CopyMemory lP2, ByVal lpStr2, 4&
        
        'swap vb6 pointers
        CopyMemory ByVal lpStr1, lP2, 4&
        CopyMemory ByVal lpStr2, lP1, 4&
End Sub
'common key sub
Public Function KeyPressed(ByRef vkKey As Long) As Boolean
    Dim TKeyPressed As Integer: TKeyPressed = GetAsyncKeyState(vkKey)
    KeyPressed = IIf(TKeyPressed = -32767 Or TKeyPressed = -32768, True, False)
End Function

'Замена кириллицы на схожую по виду латиницу
Private Sub sb_CyrFilterEx(ByRef sStr$)
    Dim TLNG&
    For TLNG = 1& To Len(sStr)
        Select Case Asc(Mid$(sStr, TLNG, 1&))
            Case 224: 'а , a
               Mid$(sStr, TLNG, 1&) = "a"
            Case 229: 'е , e
               Mid$(sStr, TLNG, 1&) = "e"
            Case 238: 'о , o
               Mid$(sStr, TLNG, 1&) = "o"
            Case 240: 'р , p
               Mid$(sStr, TLNG, 1&) = "p"
            Case 241: 'с , c
               Mid$(sStr, TLNG, 1&) = "c"
            Case 243: 'у , y
               Mid$(sStr, TLNG, 1&) = "y"
            Case 245: 'х , x
               Mid$(sStr, TLNG, 1&) = "x"
'            Case 228: 'д , g
'               Mid$(sStr, Tlng, 1&) = "g"
'            Case 239: 'п , n
'               Mid$(sStr, Tlng, 1&) = "n"
            Case 179: 'і , i
                Mid$(sStr, TLNG, 1&) = "i"
               
            Case 192: 'А , A
               Mid$(sStr, TLNG, 1&) = "A"
            Case 197: 'Е , E
               Mid$(sStr, TLNG, 1&) = "E"
            Case 202: 'К , K
               Mid$(sStr, TLNG, 1&) = "K"
            Case 206: 'О , O
               Mid$(sStr, TLNG, 1&) = "O"
            Case 208: 'Р , P
               Mid$(sStr, TLNG, 1&) = "P"
            Case 209: 'С , C
               Mid$(sStr, TLNG, 1&) = "C"
            Case 213: 'Х , X
               Mid$(sStr, TLNG, 1&) = "X"
            Case 210: 'Т , T
               Mid$(sStr, TLNG, 1&) = "T"
            Case 205: 'Н , H
               Mid$(sStr, TLNG, 1&) = "H"
            Case 178: 'І , I
               Mid$(sStr, TLNG, 1&) = "I"
        End Select
    Next TLNG
End Sub
Private Sub sb_NyaFilter(ByRef sStr$)
    Dim TLNG&
    For TLNG = 1& To Len(sStr)
        If Rnd <= 0.5 Then
            Mid$(sStr, TLNG, 1&) = UCase$(Mid$(sStr, TLNG, 1&))
        Else
            Mid$(sStr, TLNG, 1&) = LCase$(Mid$(sStr, TLNG, 1&))
        End If
    Next TLNG
End Sub
'Сервисная версия, обрабатывает только кириллические символы ВЕРХНЕГО РЕГИСТРА
Private Sub sb_MedvedFilter(ByRef sStr$)
'Полный список активированных преобразований:
'Э , И
'І , И
'Є , И
'Ї , И
'Ґ , Х
'Ф , Х
'Г , Х
'П , Х
'Б , Х
'В , Х
'З , С
'Ц , С
'Д , Т
'Я , А
'Ю , У
'М , Н
'Ч , Ш
'Ж , Ш
'Щ , Ш

'Сомнительно:
'Е , И

    Dim TLNG&
    
    For TLNG = 1& To Len(sStr)
        Select Case Asc(Mid$(sStr, TLNG, 1&))
            'Хуй знает, убрать Е-И или оставить. Пусть решает рандом :3
            'Case 197: 'Е , И
                'Mid$(sStr, TLNG, 1&) = "И"

            Case 221: 'Э , И
               Mid$(sStr, TLNG, 1&) = "И"
            Case 178: 'І , И
               Mid$(sStr, TLNG, 1&) = "И"
            Case 170: 'Є , И
               Mid$(sStr, TLNG, 1&) = "И"
            Case 175: 'Ї , И
               Mid$(sStr, TLNG, 1&) = "И"
               
            'Х
            Case 165: 'Ґ , Х
               Mid$(sStr, TLNG, 1&) = "Х"
            Case 212: 'Ф , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            Case 195: 'Г , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            Case 207: 'П , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            Case 193: 'Б , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            Case 194: 'В , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            
            'С
            Case 199: 'З , С
                Mid$(sStr, TLNG, 1&) = "С"
            Case 214: 'Ц , С
                Mid$(sStr, TLNG, 1&) = "С"

            Case 196: 'Д, Т
                Mid$(sStr, TLNG, 1&) = "Т"
                
            Case 223: 'Я , А
                 Mid$(sStr, TLNG, 1&) = "А"
            Case 222: 'Ю , У
                Mid$(sStr, TLNG, 1&) = "У"
            
            Case 204: 'М , Н
                Mid$(sStr, TLNG, 1&) = "Н"
     
            Case 215: 'Ч , Ш
                Mid$(sStr, TLNG, 1&) = "Ш"
            Case 198: 'Ж , Ш
                Mid$(sStr, TLNG, 1&) = "Ш"
            Case 217: 'Щ , Ш
                Mid$(sStr, TLNG, 1&) = "Ш"
            
        End Select
    Next TLNG
End Sub
'Обычная версия, включена обработка обоих регистров кириллицы
Private Sub sb_MedvedFilterEx(ByRef sStr$)
    Dim TLNG&
    For TLNG = 1& To Len(sStr)
        Select Case Asc(Mid$(sStr, TLNG, 1&))
            Case 193: 'Б , В-Ф
                If Rnd <= 0.5 Then
                    If Rnd <= 0.8 Then
                        Mid$(sStr, TLNG, 1&) = "В"
                    Else
                        Mid$(sStr, TLNG, 1&) = "Ф"
                    End If
                End If
                
            'ukr
            Case 165: 'Ґ , Г-Х
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "Г"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "Х"
                End If
            Case 178: 'І , И
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "И"
                End If
            Case 170: 'Є , И-Ї-Е
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "И"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "Ї"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "Е"
                End If
            Case 175: 'Ї , І
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "І"
                End If
            '  '
            
            Case 195: 'Г , Х
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "Х"
                End If
            Case 213: 'Х , Г
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "Г"
                End If
            Case 193: 'Д , Т
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "Т"
                End If
            Case 210: 'Т , Д
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "Д"
                End If
            Case 197: 'Е , И-Э
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "И"
                Else
                    Mid$(sStr, TLNG, 1&) = "Э"
                End If
            Case 168: 'Ё , И-Э
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "И"
                Else
                    Mid$(sStr, TLNG, 1&) = "Э"
                End If
                
            Case 198: 'Ж , Ш-С
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "Ш"
                Else
                    Mid$(sStr, TLNG, 1&) = "С"
                End If
                
            Case 199: 'З , С
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "С"
                End If
            Case 212: 'Ф , Х
                Mid$(sStr, TLNG, 1&) = "Х"
            Case 214: 'Ц , С
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "С"
                End If
            Case 215: 'Ч , Ш
                Mid$(sStr, TLNG, 1&) = "Ш"
            Case 221: 'Э , И
                Mid$(sStr, TLNG, 1&) = "И"
            Case 222: 'Ю , У
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "У"
                End If
            Case 223: 'Я , А
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "А"
                End If
'**********************************************8
            Case 225: 'б , в-ф
                If Rnd <= 0.5 Then
                    If Rnd <= 0.8 Then
                        Mid$(sStr, TLNG, 1&) = "в"
                    Else
                        Mid$(sStr, TLNG, 1&) = "ф"
                    End If
                End If

            'ukr
            Case 180: 'ґ , г-х
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "г"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "х"
                End If
            Case 179: 'і , и
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "и"
                End If
            Case 186: 'є , И-Ї-Е
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "є"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "ї"
                ElseIf Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "е"
                End If
            Case 191: 'ї , і
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "і"
                End If
            '  '


            Case 227: 'г , х
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "х"
                End If
            Case 245: 'х , г
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "г"
                End If
            Case 228: 'д , т
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "т"
                End If
            Case 242: 'т , д
                If Rnd <= 0.7 Then
                    Mid$(sStr, TLNG, 1&) = "д"
                End If
            Case 229: 'е , и-э
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "и"
                Else
                    Mid$(sStr, TLNG, 1&) = "э"
                End If
            Case 184: 'ё , и-э
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "и"
                Else
                    Mid$(sStr, TLNG, 1&) = "э"
                End If
                
            Case 230: 'ж , ш-с
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "ш"
                Else
                    Mid$(sStr, TLNG, 1&) = "с"
                End If
                
            Case 231: 'з , с
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "с"
                End If
            Case 244: 'ф , х
                Mid$(sStr, TLNG, 1&) = "х"
            Case 246: 'ц , с
                If Rnd <= 0.5 Then
                    Mid$(sStr, TLNG, 1&) = "с"
                End If
            Case 247: 'ч , ш
                Mid$(sStr, TLNG, 1&) = "ш"
            Case 253: 'э , и
                Mid$(sStr, TLNG, 1&) = "и"
            Case 254: 'ю , у
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "у"
                End If
            Case 255: 'я , а
                If Rnd <= 0.8 Then
                    Mid$(sStr, TLNG, 1&) = "а"
                End If
        End Select
    Next TLNG
End Sub
Private Function zcf_GetStringsWord(ByRef iLng As Long)
    Dim StringsS As String: StringsS = Right$(CStr(iLng), 2&)
    Dim StringsL As Byte:   StringsL = CByte(StringsS)

    If (StringsL Mod 10 = 0) Or (StringsL >= 11 And StringsL <= 19) Then
        zcf_GetStringsWord = "строк" '0, 10, 20, 30 ..., 11-12-13...19
    ElseIf Val(Right$(StringsS, 1&)) >= 5 Then
        zcf_GetStringsWord = "строк" '5-9
    ElseIf Val(Right$(StringsS, 1&)) >= 2 Then
        zcf_GetStringsWord = "строки" '2-4
    Else
        zcf_GetStringsWord = "строка" '1
    End If
End Function

Private Function SHIT_Part1(ByRef rArr() As String, ByRef wCnt As Long)
'Just build array of raw strings

Dim TLNG As Long
Dim Selected As Long 'id of string that will be included into result array
Dim RetryCount As Long 'number of tries to bypass cache
Dim UfoWords As Long
Dim UfoCLen As Long

For wCnt = 1& To CLng(wCnt)
    'UFO
    If opChaos Then
        SHIT_Part1 = SHIT_Part1 + 1&
        'Add string to result
        Select Case Rnd
            Case Is <= 0.07:     UfoWords = 1&
            Case Is <= 0.22:    UfoWords = 2&
            Case Is <= 0.45:     UfoWords = 3&
            Case Is <= 0.65:     UfoWords = 4&
            Case Is <= 0.75:     UfoWords = 5&
            Case Is <= 0.85:    UfoWords = 6&
            Case Is <= 0.95:     UfoWords = 7&
            Case Else:          UfoWords = 8&
        End Select
        
        For TLNG = 1& To UfoWords
            Select Case Rnd
                Case Is <= 0.15:     UfoCLen = 0&
                Case Is <= 0.3:      UfoCLen = 2&
                Case Is <= 0.5:      UfoCLen = 3&
                Case Is <= 0.74:     UfoCLen = 4&
                Case Is <= 0.82:     UfoCLen = 5&
                Case Is <= 0.9:      UfoCLen = 6&
                Case Is <= 0.95:     UfoCLen = 7&
                Case Else:           UfoCLen = 8&
            End Select
            
            If Not TLNG = UfoWords Then
                rArr(SHIT_Part1) = rArr(SHIT_Part1) & nName(2& + UfoCLen, 1 + RNDINT(1), False) & " "
            End If
            
        Next TLNG
        rArr(SHIT_Part1) = UCase$(Trim$(rArr(SHIT_Part1)))

    Else
        SHIT_Part1 = SHIT_Part1 + 1&
        If wCnt > SWCnt + 1& Then Exit For
        'Random string to append
        Selected = RNDINT(SWCnt)
        
        'Cache
        If opCache Then
            RetryCount = 0
            Selected = -1&
                Do While fc_Cache(Selected) = False
                    If RetryCount > 32& Then
                        Selected = -1&
                            Exit Do
                    Else
                        Selected = RNDINT(SWCnt)
                        RetryCount = RetryCount + 1
                    End If
                Loop
        End If
        
        'Add string to result after it passed cache
        If Not Selected = -1& Then
            If opOmsk Then  'plain text chaos
                rArr(SHIT_Part1) = UCase$(fc_StrWordsResort(Trim$(SWArr.MArr(Selected)), False))
            Else 'plain text or lyrics
                rArr(SHIT_Part1) = UCase$(Trim$(SWArr.MArr(Selected)))
            End If
        End If
    End If
    
    If Len(rArr(SHIT_Part1)) = 0& Then
        If Not SHIT_Part1 = 0& Then SHIT_Part1 = SHIT_Part1 - 1&
    End If
Next wCnt
End Function


Private Function SHIT_Part2(ByRef RCnt As Long, ByRef rArr() As String)
    Dim TLNG As Long
    Dim lCurrent As Long 'Current ryphm, lyrics mode
    Dim lRecent As Long
    Dim Matches(0& To 31&) As Long
    Dim MatchesL As Long
    Dim RetryCount As Long
    Dim lCnt As Long
    Dim lArr() As String
    
    TLNG = (RCnt \ 2&) * 2&
        If Not TLNG = 0& Then
            RCnt = TLNG
            If Not UBound(rArr) = RCnt Then
                ReDim Preserve rArr(1& To RCnt) As String
            End If
        End If
        
    ReDim lArr(1& To RCnt&) As String
    
    lCurrent = -1&
    For TLNG = 1& To RCnt
        If TLNG Mod 2 = 0 Then
            lRecent = TLNG - 1&
            
            'Get all available ryphms by 3 sym
            MatchesL = fc_RyphmLine(rArr(lRecent), 3&, Matches(), 31&)
'            If MatchesL = 31 Then Stop
            
            'Then try with 2 if 1st attempt failed
            If MatchesL = -1& Then
                MatchesL = fc_RyphmLine(rArr(lRecent), 2&, Matches(), 31&)
            End If
            
            'If still fail, there is no ryphm
            If MatchesL = -1& Then
                lCurrent = -1&
                
            Else 'Chose one of results received
                'Lyrics with cache
                If opCache Then
                    RetryCount = 0
                    lCurrent = -1&
                    'To bypass cache, use randomized bruteforce first. Retry count should be limited.
                    Do While fc_Cache(lCurrent) = False
                        If RetryCount > MatchesL Then lCurrent = -1&: Exit Do
                            lCurrent = Matches(RNDINT(MatchesL))
                            RetryCount = RetryCount + 1&
                    Loop
    
                    'There is only one way, if you have failed
                    If lCurrent = -1& Then
                        For RetryCount = 0& To MatchesL
                            lCurrent = Matches(RetryCount) 'The grotesque and the linear way
                                If fc_Cache(lCurrent) Then
                                    Exit For
                                Else
                                    lCurrent = -1&
                                End If
                        Next RetryCount
                    End If
                Else 'Lyrics W/O cache
                    lCurrent = Matches(RNDINT(MatchesL))
                End If
            End If

            'lyrics text complete
            If Not lCurrent = -1& Then
                If opOmsk Then
                    rArr(TLNG) = UCase$(fc_StrWordsResort(Trim$(SWArr.MArr(lCurrent)), True))
                Else
                    rArr(TLNG) = UCase$(Trim$(SWArr.MArr(lCurrent)))
                End If
                
                lCnt = lCnt + 2&
                
                zcs_SwapStrings lArr(lCnt), rArr(TLNG)
                zcs_SwapStrings lArr(lCnt - 1&), rArr(TLNG - 1&)
                
            ElseIf Not lCnt = 0& Then
                rArr(TLNG) = vbNullString
                rArr(TLNG - 1&) = vbNullString
            End If
        End If
    Next TLNG

'Remove broken strings
'Debug.Print "No ryphm: " & rCnt - lCnt
If Not lCnt = 0& Then
    If Not UBound(lArr) = lCnt Then
        ReDim Preserve lArr(1& To lCnt) As String
    End If
    
    rArr = lArr
    RCnt = lCnt
End If

'Another kind of ryphm [swap each 2nd and 3rd strings]
    If Rnd <= 0.2 Then
        For TLNG = 4& To RCnt Step 4&
            zcs_SwapStrings rArr(TLNG - 1&), rArr(TLNG - 2&)
        Next TLNG
    End If
End Function


'заменяет некоторые слова на другие с похожими окончаниями
'с определенной вероятностью
Private Sub sb_Rewrite(ByRef iArr$(), ByRef iCount As Long, ByRef fixLastWord As Boolean)
    Dim WordsArr() As String
    Dim sArr() As String
    Dim iTmp As Long
    Dim wTmp As Long

    Const extStr As String = "НЫЙ НАЯ НОЙ НЫЕ ЕЛА АЛА ТАЛ БАЛ СИТ ВАЛ ДЕТ АТЬ НЫМ РИТ ЙТЕ ИТЕ ИТЬ ДАЛ ЬСЯ ШУТ ШАТ ПИТ РИЛ ДИТ ТСЯ ИЖУ АЛИ АЮТ УШИ ВЕМ ИШЬ ИСЬ КАЯ ЩАТ НИТ ЗЕТ УЕТ НЕТ ЬМИ ЛСЯ ИЛА МАЛ ДЯТ НЯЛ ЕШЬ ШЕН МЕМ АЕТ НАЛ РАЛ ПЕЛ ОВИ АСЬ ЯЮТ ТЯТ ЬЕТ ОРЮ ЧЕТ ЬЮТ ЯТЬ ТРИ СИЛ ЛЯЛ ДЕЛ УЛО ЧИТ ТЕТ ВИТ АШУ САЛ ЯЕМ ЗИЛ СЛА СЛО УТЬ ПАЛ СТЬ ЕБУ НИЛ ВИЛ ЗАЛ ЬТЕ ВАЯ ИЛИ ЧИЛ ГАЛ ИЛО СУТ ЯЕТ ТИТ КИЕ ЛАЯ ТАЯ ОГО АВЬ НЯЯ УСЬ УЛИ ОСИ КИЙ ЛАЛ НУЮ БЬЮ ЖУТ КОЙ ЧЕЕ ПЫМ ТЫМ ЗЛИ НУЛ ЖЕТ ЖИЛ ДОХ ШЛО ЗИЙ ТЫЙ ВЫЕ ЫЛИ ДАЮ ЖАТ НОЕ ОСЬ ЛЫЕ ВОМ РХУ ЖАЛ ТЫХ ХАЛ ИБИ ШЕЛ АЛО МОЙ АЕМ ОХО ГЛИ ИВУ ШАЛ ЕТЬ УЛА ЩЕТ ОВО РЕТ ХЛИ СЕТ АЖИ ЕЛИ ЬШЕ ЗТЬ НЫХ КАЛ КИХ ОМУ РАЯ НИЯ ЬНО ЕДУ ДУТ РВУ ВЫЙ ВУТ ТУЮ ШИЙ ИВО ДИЛ НИЕ НУТ ЯЛИ КАЙ ННО ЫМИ ТЫЕ АНА ЕНА ЙСЯ ЛУЮ КЛО НИЙ ЕЖУ ЧИЕ ЙСЯ КОЙ АТЬ ХЛИ АЛИ БЕТ АЮТ ЩИЙ НЕЕ ЕТЕ ГАЯ ЩАЛ ЖНО МАЮ МЕЛ СЕН"
    
    Dim extLen As Long
        extLen = 3&

    For iTmp = 1& To iCount
        WordsArr = Split(iArr(iTmp), " ", -1&, vbBinaryCompare)
            
        'exept last words
        'если fixLastWord -- последнее слово фразы не будет трансформироваться, нужно для режима лирика чтобы не выпадало одинаковой рифмы
        For wTmp = 0& To IIf(fixLastWord, UBound(WordsArr) - 1&, UBound(WordsArr))
            If Len(WordsArr(wTmp)) > extLen Then
                

                'можно ограничить список разрешенных для замены окончаний чтобы речь была более осмысленной
                If Rnd <= 0.5 Then
                If Not InStr(1&, extStr, Right$(WordsArr(wTmp), extLen), vbBinaryCompare) = 0& Then
                    'замена производится по последним символам слова
                    'fc_RyphmWord or fc_RyphmWordEx? who knows
                        If Not fc_RyphmWord(Right$(WordsArr(wTmp), extLen), extLen, sArr, False, 2048) = -1& Then
                            WordsArr(wTmp) = sArr(RNDINT(UBound(sArr)))
                        End If
                End If
                End If

            End If
        Next wTmp
        iArr(iTmp) = Join(WordsArr, " ")
    Next iTmp
End Sub
'делает некоторые опечатки
'в планах - добавить замену букв на схожие. Т.е гласные на гласные/ пригласные на пригласные
'так же перестановку двух букв местами
Private Sub sb_Mistakes(ByRef iArr$(), ByRef iCount As Long, ByRef ReserveSymbols As Boolean)
    Dim iTmp As Long
    Dim aTmp As Long
    Dim mCnt As Long
    Dim sByt As Long
    Const SymMask As String = "БВГДЖЗКЛМНПРСТФХЦЧШЩЪЬЭЫЮЯАЕИЙОУ"
    For iTmp = 1& To iCount
        For aTmp = 2& To Len(iArr(iTmp)) - IIf(ReserveSymbols, 3&, 1&)
            If Rnd <= 0.065 Then 'main chance
                If Not Asc(Mid$(iArr(iTmp), aTmp - 1&, 1&)) = 32 Then
                If Not Asc(Mid$(iArr(iTmp), aTmp + 1&, 1&)) = 32 Then
                    sByt = InStr(1&, SymMask, Mid$(iArr(iTmp), aTmp, 1&), vbBinaryCompare)
                    If sByt >= 23& Then 'Э и дальше
                        If Rnd <= 0.4 Then
                            If Rnd <= 0.5 Then
                                Mid$(iArr(iTmp), aTmp, 1&) = Mid$(SymMask, 1& + RNDINT(Len(SymMask) - 1&), 1&)
                            Else
                                Mid$(iArr(iTmp), aTmp, 1&) = Chr$(0&)
                                    mCnt = mCnt + 1&
                            End If
                        End If
                    ElseIf sByt >= 1& Then
                            If Rnd <= 0.5 Then
                                Mid$(iArr(iTmp), aTmp, 1&) = Mid$(SymMask, 1& + RNDINT(Len(SymMask) - 1&), 1&)
                            Else
                                Mid$(iArr(iTmp), aTmp, 1&) = Chr$(0&)
                                    mCnt = mCnt + 1&
                            End If
                    End If
                End If
                End If
            End If
        Next aTmp
        
        If Not mCnt = 0& Then iArr(iTmp) = Replace$(iArr(iTmp), Chr$(0&), vbNullString, 1&, -1&, vbBinaryCompare)
        mCnt = 0&
    Next iTmp
End Sub

'bit functions
Private Function orCheck(ByRef iByte As Byte, ByVal iBIT As Byte) As Byte
  orCheck = IIf((iByte And iBIT) = iBIT, 1, 0)
End Function

Private Sub orSet(ByRef iByte As Byte, ByVal iBIT As Byte, ByVal iState As Byte)
    If Not iState = 0 Then
        If orCheck(iByte, iBIT) = 0 Then
            iByte = iByte Or iBIT
        End If
    Else
        If orCheck(iByte, iBIT) = 1 Then
            iByte = iByte Xor iBIT
        End If
    End If
End Sub



Private Sub a_loadInnerFile(ByRef oStr As String)
Dim Bytes() As Byte
Dim fID As Long
On Local Error Resume Next
    Select Case InnerFile
        Case 1: fID = 242 'стандартный словарь
        Case 2: fID = 232 'стандартный треш-словарь
        Case Else:
            Exit Sub
    End Select
 
    Bytes = LoadResData(fID, 42) 'swfile
    
    'read file
'        Open AppSPath & "_DICTENCRYPT_out.bin" For Binary As 1
'            ReDim Bytes(0 To LOF(1) - 1)
'                Get 1, 1&, Bytes
'        Close 1
    
    If Err.Number = 0& Then
        'установить ключ
        Call CopyMemory(ByVal VarPtr(byteKey(0)), 242&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(1)), 124&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(2)), 8&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(3)), 45&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(4)), 224&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(5)), 88&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(6)), 111&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(7)), 201&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(8)), 124&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(9)), 168&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(10)), 198&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(11)), 244&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(12)), 54&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(13)), 86&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(14)), 102&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(15)), 151&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(16)), 12&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(17)), 208&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(18)), 143&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(19)), 125&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(20)), 199&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(21)), 0&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(22)), 18&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(23)), 23&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(24)), 2&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(25)), 61&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(26)), 95&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(27)), 146&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(28)), 81&, 4&)
        Call CopyMemory(ByVal VarPtr(byteKey(29)), 0&, 3&)
        Call CopyMemory(ByVal VarPtr(byteKey(30)), 129&, 2&)
        Call CopyMemory(ByVal VarPtr(byteKey(31)), 121&, 1&)
            
        'попытка расшифровки
        Call dataDecrypt(Bytes, oStr)
        Erase Bytes
        'стереть ключь
            For fID = 0& To 31&
                Call CopyMemory(ByVal VarPtr(byteKey(fID)), fID, 1&)
            Next fID
    End If
End Sub

Private Sub a_loadFile()
On Error Resume Next
    Static Started As Boolean
    Dim NDate As String
    
    Static LFile As String
    Dim Mstr As String
    Dim RawSize As Long

    Dim LoadTitle As String
    
    Dim Init1 As Long
    Dim Init2 As Long
    Dim Init3 As Long
    
    If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)

        'fix file namex
        If Len(sConfig.sSourceName) = 0& Then
            sConfig.sSourceName = mnuFileEn(2).Caption '"source_sw.txt"
        End If
        'set inner file flag
        If StrComp(sConfig.sSourceName, mnuFileEn(1).Caption, vbTextCompare) = 0& Then
            InnerFile = 1
        ElseIf StrComp(sConfig.sSourceName, mnuFileEn(2).Caption, vbTextCompare) = 0& Then
            InnerFile = 2
        Else
            InnerFile = 0
        End If
        
        'yoba-protection v2
        If Scriptkiddie <> 0 Then
            InnerFile = 0
        End If
        
        If SW_OpState = 0 Then
            LoadTitle = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "Загрузка..." & vbNewLine & vbNewLine & sConfig.sSourceName
        End If
            
        'load base
        If Not (InnerFile = 0) Then
            'do not reload same file
            If (StrComp(LFile, sConfig.sSourceName, vbBinaryCompare) = 0&) Then Exit Sub
                LFile = sConfig.sSourceName
                LDate = vbNullString
                SWCnt = -1&
                    If Not Len(LoadTitle) = 0& Then Call MainOut(LoadTitle, vbCenter)
                Call a_loadInnerFile(Mstr)
                RawSize = Len(Mstr)
                
'                Open AppSPath & "\_DICTEN.txt" For Binary Access Write As 1
'                    Put 1, 1, Mstr
'                Close 1
'                Mstr = ""
'                 Open AppSPath & "\_DICTEN.txt" For Binary As 1
'                    RawSize = LOF(1)
'                    Mstr = Space$(RawSize)
'                        Get 1, 1&, Mstr
'                Close 1
                'Call DeleteFile(AppSPath & "\_DICTEN.txt")
                
                    
        Else
            If Not PathFileExists(AppSPath & sConfig.sSourceName) = 0& Then
            'get current file date
            NDate = FileDate(AppSPath & sConfig.sSourceName)
            'reload file only if changed
            If (StrComp(NDate, LDate, vbBinaryCompare) = 0&) And (StrComp(LFile, sConfig.sSourceName, vbBinaryCompare) = 0&) Then Exit Sub
                LFile = sConfig.sSourceName
                LDate = NDate
                SWCnt = -1&
                    If Not Len(LoadTitle) = 0& Then Call MainOut(LoadTitle, vbCenter)
            'read file
                Open AppSPath & sConfig.sSourceName For Binary As 1
                    RawSize = LOF(1)
                    Mstr = Space$(RawSize)
                        Get 1, 1&, Mstr
                Close 1
            End If
        End If
        frmMain.Caption = zcf_nStr("\jjlgx}f{j}/T") & zcf_nStr("<^")

'    Dim UsedSize As Long
'    Dim UsedOffs As Long
'        ограничить размер загружаемого текста
'        Dim Limited As Boolean
'        UsedSize = 7000 + RNDINT(1000)
'        Limited = UsedSize < RawSize
'            If UsedSize > RawSize Then UsedSize = RawSize: Limited = False
'        UsedOffs = RawSize - UsedSize
'        UsedOffs = 1 + RNDINT(UsedOffs)
'
'        If Limited Then
'            Mstr = Mid$(Mstr, UsedOffs, UsedSize)
'        End If
    
        'load base
        Init1 = GetTickCount
        Mstr = Replace$(Mstr, Chr$(10&) & Chr$(13&), vbNewLine, 1&, -1&)
        
'        ограничить размер загружаемого текста
'            'fix loaded file part
'                Dim L1&
'                L1 = InStr(1&, Mstr, vbNewLine, vbBinaryCompare)
'                    If Not L1 = 0& Then
'                        Mid$(Mstr, 1&, L1 + 1&) = String$(L1 + 1&, " ")
'                    End If
'                L1 = InStrRev(Mstr, vbNewLine, -1&, vbBinaryCompare)
'                    If (Not L1 = 0&) And (Not L1 = Len(Mstr) - 1&) Then
'                        Mid$(Mstr, L1, 1 + Len(Mstr) - L1) = String$(1 + Len(Mstr) - L1, " ")
'                    End If
'                Mstr = Trim$(Mstr)

        Mstr = UCase$(Mstr)
        Init1 = GetTickCount - Init1
        
        'Define array
        Init2 = GetTickCount
        SWArr.MArr = Split(Mstr, vbNewLine, -1, vbBinaryCompare)
        SWCnt = UBound(SWArr.MArr)
        Init2 = GetTickCount - Init2
        
        'Define lyrics array
        Init3 = GetTickCount
        sb_MedvedFilter Mstr
        SWArr.MArrRyphm = Split(Mstr, vbNewLine, -1, vbBinaryCompare)
        Init3 = GetTickCount - Init3
        
        'Set cache max size
        MaxUsed = ((SWCnt / 100) * 15) + 1
            If MaxUsed > 256& Then MaxUsed = 256&
        'Cache flush
        sUsed = Space$(MaxUsed * 10&)
        sCnt = 0

        'update caption
        frmMain.Caption = Replace$(frmMain.Caption, "?", SWCnt + 1&, 1&, 1&, vbBinaryCompare)

        frmMain.LinkTopic = zcf_nStr(frmMain.Caption)
        frmMain.Caption = zcf_nStr(frmMain.LinkTopic)
        
        'STOP if there are other operations
        If Not SW_OpState = 0 Then Exit Sub
        If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)
    
        'show about info at program start sometimes
        If Not Started Then
            Started = True
                If Rnd <= 0.2 Then aux_showAbout: Exit Sub 'show about window instead of file info
        End If
        
        'show file info
        MainOut vbNewLine & "Файл:" & vbNewLine & sConfig.sSourceName & vbNewLine & vbNewLine & SWCnt + 1& & vbNewLine & Replace$(FormatNumber(RawSize / 1024, 2&), Chr$(0&), Chr(32), 1&, -1&, vbBinaryCompare) & " кб" & vbNewLine & vbNewLine & "I1: " & Init1 & vbNewLine & "I2: " & Init2 & vbNewLine & "I3: " & Init3 & vbNewLine & vbNewLine & "DONE", vbCenter
End Sub

Private Sub aux_showAbout()
    Dim AboutMsg(10 To 23) As String
'    Dim RndWord As String
    Static RndPhrase As String
    
    If Not SW_OpState = 0 Then Exit Sub
    If Not AutoGen = 0 Then Call SHIT_Autogen(AutoGen, vbRightButton)
        'phrase
        If Not SWCnt = -1& Then
            If Len(RndPhrase) = 0 Or Rnd <= 0.2 Then
                RndPhrase = SWArr.MArr(RNDINT(SWCnt))
                Call sb_RandomSigns(RndPhrase, 0, 1)
            End If
        Else
            RndPhrase = "АБАСРАЦА"
        End If
            
            RndPhrase = LCase$(RndPhrase)
                Mid$(RndPhrase, 1&, 1&) = UCase$(Mid$(RndPhrase, 1&, 1&))
        'word
'        If Not SWCnt = -1& And Rnd <= 0.6 Then
'            RndWord = fc_GetLastWord(IIf(Rnd <= 0.5, SWArr.MArr(RNDINT(SWCnt)), SWArr.MArrRyphm(RNDINT(SWCnt))), True, False, 0&)
'        Else
'            RndWord = nName(2& + RNDINT(5), 1 + RNDINT(1), False)
'        End If
'            RndWord = LCase$(RndWord)
'                Mid$(RndWord, 1&, 1&) = UCase$(Mid$(RndWord, 1&, 1&))

    'about text
    AboutMsg(10) = "v" & appMajor & "." & appMinor & "." & appBuild
    '— Сайт —
    AboutMsg(12) = zcf_nStr("Ћ9Чнн9ищрл9с9ыиь9лщучь9Ћ")
    AboutMsg(13) = zcf_nStr("qmmi#66or7zvt6|u|zmkvmxk}")
    AboutMsg(14) = zcf_nStr("sook!4|~txror~h5lh4wnuzhtw~")

    
    '—— Создатель ——
    AboutMsg(16) = zcf_nStr("(’%Ессчех%фчечам)%чапфчлз%н%клфчлз%’(")
    AboutMsg(17) = zcf_nStr("{pcpg;papg{tyU`~g;{pa")
    
    '—— Идея режима "Лирика" и не только ——
    AboutMsg(19) = zcf_nStr("*ђ'Пгвш'чвбплз'%Мпчпнз%'п'кв'хймынй'ђ*")
    AboutMsg(20) = zcf_nStr("fprt|yzJq`zfUxt|y;g`")

    'Случайная фраза
    AboutMsg(23) = """" & RndPhrase & """"

    
    'Вывод текста
    MainOut Replace$(Join(AboutMsg, vbNewLine), "'", """", 1&, -1&, vbBinaryCompare), vbCenter
End Sub

Private Sub ShowMenu()
    Dim TLNG&
    FileCount = 0&
    ReDim FileList(1& To 1&)
    Call GetFiles(App.Path, False)
    
    mnuFileEn(1).Checked = InnerFile = 1
    mnuFileEn(2).Checked = InnerFile = 2
    
    For TLNG = 1& To mnuFile.UBound
        If Not TLNG > FileCount Then
            mnuFile(TLNG).Visible = True
            mnuFile(TLNG).Enabled = True
            mnuFile(TLNG).Caption = FileList(TLNG)
            mnuFile(TLNG).Checked = (InnerFile = 0) And (StrComp(sConfig.sSourceName, FileList(TLNG), vbTextCompare) = 0&)
        Else
            mnuFile(TLNG).Checked = False
            If Not TLNG = mnuFile.LBound Then
                mnuFile(TLNG).Visible = False
            Else
                mnuFile(TLNG).Enabled = False
            End If
            mnuFile(TLNG).Caption = vbNullString
        End If
    Next TLNG

    mnuOpen.Caption = Mid$(mnuOpen.Caption, 1&, InStr(1&, mnuOpen.Caption, " ", vbBinaryCompare)) & """" & sConfig.sSourceName & """"
    
    mnuOpen.Enabled = InnerFile = 0
    mnuClean.Enabled = InnerFile = 0
    mnuReload.Enabled = InnerFile = 0
    
        Call PopupMenu(mnuMain)
End Sub

Private Sub SHIT_Autogen(ByRef Index As Integer, ByRef Button As Integer)
    If Button = vbRightButton Then
        If Index = AutoGen Then 'stop
            cmdS(Index).FontBold = False
            cmdS(Index).FontItalic = False
            AutoGen = 0
        ElseIf Not AutoGen = 0 Then 'switch
            cmdS(AutoGen).FontBold = False
            cmdS(AutoGen).FontItalic = False
            AutoGen = Index
            cmdS(AutoGen).FontBold = True
            cmdS(AutoGen).FontItalic = True
        Else 'start
            AutoGen = Index
            cmdS(AutoGen).FontBold = True
            cmdS(AutoGen).FontItalic = True
        End If
        
    ElseIf Button = vbLeftButton Then
        If Not AutoGen = 0 Then
            'switch autogen
                cmdS(AutoGen).FontBold = False
                cmdS(AutoGen).FontItalic = False
                AutoGen = Index
                cmdS(AutoGen).FontBold = True
                cmdS(AutoGen).FontItalic = True
        Else
            Call SHIT(Index) 'single shit
        End If
    End If
End Sub

Private Sub MainOut(ByRef outStr As String, ByRef TextAlign As AlignmentConstants)
    txtMain.Alignment = TextAlign
    txtMain.Text = outStr
End Sub
