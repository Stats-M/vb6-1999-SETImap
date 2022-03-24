VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Редактирование существующей записи"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Комментарии к текущему блоку"
      Height          =   4425
      Index           =   4
      Left            =   -20000
      TabIndex        =   156
      Top             =   735
      Width           =   9990
      Begin VB.TextBox Text16 
         Height          =   330
         Left            =   4410
         TabIndex        =   165
         Top             =   3885
         Width           =   5265
      End
      Begin VB.TextBox Text15 
         Height          =   330
         Left            =   4410
         TabIndex        =   163
         Top             =   3255
         Width           =   5265
      End
      Begin VB.TextBox Text14 
         Height          =   330
         Left            =   4410
         TabIndex        =   162
         Top             =   2625
         Width           =   5265
      End
      Begin VB.TextBox Text13 
         Height          =   330
         Left            =   4410
         TabIndex        =   158
         Top             =   735
         Width           =   5265
      End
      Begin VB.Label Label57 
         Caption         =   "Резервное поле 3"
         Height          =   225
         Left            =   840
         TabIndex        =   164
         Top             =   3938
         Width           =   1485
      End
      Begin VB.Label Label56 
         Caption         =   "Резервное поле 2"
         Height          =   225
         Left            =   840
         TabIndex        =   161
         Top             =   3308
         Width           =   1485
      End
      Begin VB.Label Label55 
         Caption         =   "Резервное поле 1"
         Height          =   225
         Left            =   840
         TabIndex        =   160
         Top             =   2678
         Width           =   1485
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         Caption         =   $"frmEdit.frx":030A
         Height          =   435
         Left            =   420
         TabIndex        =   159
         Top             =   1785
         Width           =   8730
      End
      Begin VB.Label Label53 
         Caption         =   "Ваш комментарий (будет показан на карте)"
         Height          =   330
         Left            =   210
         TabIndex        =   157
         Top             =   735
         Width           =   3900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дополнительные параметры"
      Height          =   4425
      Index           =   3
      Left            =   -20000
      TabIndex        =   15
      Top             =   735
      Width           =   9885
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   9
         Left            =   3465
         TabIndex        =   144
         Text            =   "1.30"
         Top             =   3960
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   8
         Left            =   3465
         TabIndex        =   143
         Text            =   "1048576"
         Top             =   3584
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   7
         Left            =   3465
         TabIndex        =   142
         Text            =   "8"
         Top             =   3211
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   6
         Left            =   3465
         TabIndex        =   141
         Text            =   "2048"
         Top             =   2838
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   5
         Left            =   3465
         TabIndex        =   140
         Text            =   "0x0008"
         Top             =   2465
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   4
         Left            =   3465
         TabIndex        =   139
         Text            =   "0"
         Top             =   2092
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   3
         Left            =   3465
         TabIndex        =   138
         Text            =   "encoded"
         Top             =   1719
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   2
         Left            =   3465
         TabIndex        =   137
         Text            =   "256"
         Top             =   1365
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   1
         Left            =   3465
         TabIndex        =   136
         Text            =   "seti"
         Top             =   973
         Width           =   2640
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Index           =   0
         Left            =   3465
         TabIndex        =   135
         Text            =   "work unit"
         Top             =   600
         Width           =   2640
      End
      Begin VB.Label Label52 
         Caption         =   "1.30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   155
         Top             =   3990
         Width           =   2220
      End
      Begin VB.Label Label51 
         Caption         =   "1048576"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   154
         Top             =   3614
         Width           =   2220
      End
      Begin VB.Label Label50 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   153
         Top             =   3241
         Width           =   2220
      End
      Begin VB.Label Label49 
         Caption         =   "2048"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   152
         Top             =   2868
         Width           =   2220
      End
      Begin VB.Label Label48 
         Caption         =   "0x0008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   151
         Top             =   2495
         Width           =   2220
      End
      Begin VB.Label Label47 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   150
         Top             =   2122
         Width           =   2220
      End
      Begin VB.Label Label46 
         Caption         =   "encoded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   149
         Top             =   1749
         Width           =   2220
      End
      Begin VB.Label Label45 
         Caption         =   "256"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   148
         Top             =   1376
         Width           =   2220
      End
      Begin VB.Label Label44 
         Caption         =   "seti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   147
         Top             =   1003
         Width           =   2220
      End
      Begin VB.Label Label43 
         Caption         =   "work unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7140
         TabIndex        =   146
         Top             =   630
         Width           =   2220
      End
      Begin VB.Label Label42 
         Caption         =   "Образец"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7875
         TabIndex        =   145
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "Версия пленки (tape version)"
         Height          =   225
         Index           =   9
         Left            =   210
         TabIndex        =   134
         Top             =   3990
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Частота дискретизации (nsamples)"
         Height          =   225
         Index           =   8
         Left            =   210
         TabIndex        =   133
         Top             =   3614
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "ifft len"
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   132
         Top             =   3241
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "fft len"
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   131
         Top             =   2868
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Версия сплиттера (splitter version)"
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   130
         Top             =   2495
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Класс данных (data class)"
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   129
         Top             =   2122
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Тип данных (data type)"
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   128
         Top             =   1749
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Версия (version)"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   127
         Top             =   1376
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Задача (task)"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   126
         Top             =   1003
         Width           =   2850
      End
      Begin VB.Label Label41 
         Caption         =   "Тип (type)"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   125
         Top             =   630
         Width           =   2850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Пространственно - временные координаты"
      Height          =   4530
      Index           =   2
      Left            =   -20000
      TabIndex        =   14
      Tag             =   "2"
      Top             =   735
      Width           =   9885
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   20
         Left            =   6615
         TabIndex        =   166
         Top             =   3675
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   1
         Left            =   1575
         TabIndex        =   61
         Top             =   600
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   21
         Left            =   6615
         TabIndex        =   60
         Top             =   3990
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   19
         Left            =   6615
         TabIndex        =   59
         Top             =   3120
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   18
         Left            =   6615
         TabIndex        =   58
         Top             =   2805
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   17
         Left            =   6615
         TabIndex        =   53
         Top             =   2490
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   16
         Left            =   6615
         TabIndex        =   52
         Top             =   2175
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   15
         Left            =   6615
         TabIndex        =   51
         Top             =   1860
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   14
         Left            =   6615
         TabIndex        =   50
         Top             =   1545
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   13
         Left            =   6615
         TabIndex        =   49
         Top             =   1230
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   12
         Left            =   6615
         TabIndex        =   48
         Top             =   915
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   11
         Left            =   6615
         TabIndex        =   47
         Top             =   600
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   10
         Left            =   6615
         TabIndex        =   46
         Top             =   285
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   9
         Left            =   1575
         TabIndex        =   38
         Top             =   3120
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   8
         Left            =   1575
         TabIndex        =   37
         Top             =   2805
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   7
         Left            =   1575
         TabIndex        =   36
         Top             =   2490
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   32
         Top             =   2175
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   5
         Left            =   1575
         TabIndex        =   31
         Top             =   1860
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   4
         Left            =   1575
         TabIndex        =   30
         Top             =   1545
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   29
         Top             =   1230
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   2
         Left            =   1575
         TabIndex        =   28
         Top             =   915
         Width           =   3060
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   1575
         TabIndex        =   27
         Top             =   285
         Width           =   3060
      End
      Begin VB.Label Label58 
         Caption         =   "Точка 21"
         Height          =   225
         Left            =   5145
         TabIndex        =   167
         Top             =   4020
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Точка 20"
         Height          =   225
         Left            =   5145
         TabIndex        =   57
         Top             =   3705
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Точка 19"
         Height          =   225
         Left            =   5145
         TabIndex        =   56
         Top             =   3150
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Точка 18"
         Height          =   225
         Left            =   5145
         TabIndex        =   55
         Top             =   2835
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Точка 17"
         Height          =   225
         Left            =   5145
         TabIndex        =   54
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "Точка 16"
         Height          =   225
         Left            =   5145
         TabIndex        =   45
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Точка 15"
         Height          =   225
         Left            =   5145
         TabIndex        =   44
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Точка 14"
         Height          =   225
         Left            =   5145
         TabIndex        =   43
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Точка 13"
         Height          =   225
         Left            =   5145
         TabIndex        =   42
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Точка 12"
         Height          =   225
         Left            =   5145
         TabIndex        =   41
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Точка 11"
         Height          =   225
         Left            =   5145
         TabIndex        =   40
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Точка 10"
         Height          =   225
         Left            =   5145
         TabIndex        =   39
         Top             =   315
         Width           =   855
      End
      Begin VB.Line Line6 
         X1              =   1575
         X2              =   1575
         Y1              =   3675
         Y2              =   4305
      End
      Begin VB.Line Line5 
         X1              =   2835
         X2              =   2835
         Y1              =   3675
         Y2              =   4095
      End
      Begin VB.Line Line4 
         X1              =   3885
         X2              =   1575
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Line Line3 
         X1              =   3885
         X2              =   2835
         Y1              =   4095
         Y2              =   4095
      End
      Begin VB.Line Line2 
         X1              =   3885
         X2              =   3360
         Y1              =   3885
         Y2              =   3885
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   3360
         Y1              =   3675
         Y2              =   3885
      End
      Begin VB.Label Label19 
         Caption         =   "Local DEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3885
         TabIndex        =   35
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label Label18 
         Caption         =   "Local RA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3885
         TabIndex        =   34
         Top             =   3990
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Время"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3885
         TabIndex        =   33
         Top             =   4200
         Width           =   645
      End
      Begin VB.Label Label16 
         Caption         =   "Образец :  2451423.37662  15.814  11.39"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   3465
         Width           =   3480
      End
      Begin VB.Label Label15 
         Caption         =   "Точка 9"
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   3150
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Точка 8"
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Top             =   2835
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Точка 7"
         Height          =   225
         Left            =   210
         TabIndex        =   23
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Точка 6"
         Height          =   225
         Left            =   210
         TabIndex        =   22
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Точка 5"
         Height          =   225
         Left            =   210
         TabIndex        =   21
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Точка 4"
         Height          =   225
         Left            =   210
         TabIndex        =   20
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Точка 3"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Точка 2"
         Height          =   225
         Left            =   210
         TabIndex        =   18
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Точка 1"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Точка 0"
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Физические параметры блока"
      Height          =   4425
      Index           =   1
      Left            =   -20000
      TabIndex        =   13
      Top             =   735
      Width           =   9885
      Begin VB.OptionButton Option2 
         Caption         =   "Десятичные единицы"
         Height          =   225
         Index           =   1
         Left            =   1680
         TabIndex        =   112
         Top             =   315
         Width           =   2010
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Градусы"
         Height          =   225
         Index           =   0
         Left            =   525
         TabIndex        =   111
         Top             =   315
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.TextBox Text11 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   110
         Top             =   2175
         Width           =   435
      End
      Begin VB.TextBox Text10 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4935
         MaxLength       =   2
         TabIndex        =   109
         Top             =   2175
         Width           =   435
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Index           =   3
         Left            =   3990
         MaxLength       =   3
         TabIndex        =   108
         Top             =   2175
         Width           =   435
      End
      Begin VB.TextBox Text11 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   107
         Top             =   1685
         Width           =   435
      End
      Begin VB.TextBox Text10 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4935
         MaxLength       =   2
         TabIndex        =   106
         Top             =   1685
         Width           =   435
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3990
         MaxLength       =   2
         TabIndex        =   105
         Top             =   1685
         Width           =   435
      End
      Begin VB.TextBox Text11 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   104
         Top             =   1195
         Width           =   435
      End
      Begin VB.TextBox Text10 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4935
         MaxLength       =   2
         TabIndex        =   103
         Top             =   1195
         Width           =   435
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "+dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3990
         MaxLength       =   3
         TabIndex        =   102
         Top             =   1195
         Width           =   435
      End
      Begin VB.TextBox Text11 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   101
         Top             =   705
         Width           =   435
      End
      Begin VB.TextBox Text10 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4935
         MaxLength       =   2
         TabIndex        =   100
         Top             =   705
         Width           =   435
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3990
         MaxLength       =   2
         TabIndex        =   99
         Top             =   705
         Width           =   435
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   7
         Left            =   3990
         TabIndex        =   89
         Top             =   3990
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   6
         Left            =   3990
         TabIndex        =   88
         Top             =   3560
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   5
         Left            =   3990
         TabIndex        =   87
         Top             =   3130
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   4
         Left            =   3990
         TabIndex        =   86
         Top             =   2700
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   3
         Left            =   3990
         TabIndex        =   85
         Top             =   2175
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   3990
         TabIndex        =   84
         Top             =   1685
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   3990
         TabIndex        =   83
         Top             =   1195
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   3990
         TabIndex        =   82
         Top             =   705
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label Label40 
         Caption         =   "сек"
         Height          =   285
         Index           =   11
         Left            =   6405
         TabIndex        =   124
         Top             =   2175
         Width           =   330
      End
      Begin VB.Label Label40 
         Caption         =   "мин"
         Height          =   285
         Index           =   10
         Left            =   5460
         TabIndex        =   123
         Top             =   2175
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "град"
         Height          =   285
         Index           =   9
         Left            =   4515
         TabIndex        =   122
         Top             =   2175
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "сек"
         Height          =   285
         Index           =   8
         Left            =   6405
         TabIndex        =   121
         Top             =   1685
         Width           =   330
      End
      Begin VB.Label Label40 
         Caption         =   "мин"
         Height          =   285
         Index           =   7
         Left            =   5460
         TabIndex        =   120
         Top             =   1685
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "час"
         Height          =   285
         Index           =   6
         Left            =   4515
         TabIndex        =   119
         Top             =   1685
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "сек"
         Height          =   285
         Index           =   5
         Left            =   6405
         TabIndex        =   118
         Top             =   1195
         Width           =   330
      End
      Begin VB.Label Label40 
         Caption         =   "мин"
         Height          =   285
         Index           =   4
         Left            =   5460
         TabIndex        =   117
         Top             =   1195
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "град"
         Height          =   285
         Index           =   3
         Left            =   4515
         TabIndex        =   116
         Top             =   1195
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "сек"
         Height          =   285
         Index           =   2
         Left            =   6405
         TabIndex        =   115
         Top             =   705
         Width           =   330
      End
      Begin VB.Label Label40 
         Caption         =   "мин"
         Height          =   285
         Index           =   1
         Left            =   5460
         TabIndex        =   114
         Top             =   705
         Width           =   435
      End
      Begin VB.Label Label40 
         Caption         =   "час"
         Height          =   285
         Index           =   0
         Left            =   4515
         TabIndex        =   113
         Top             =   705
         Width           =   435
      End
      Begin VB.Label Label39 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   7140
         TabIndex        =   98
         Top             =   4020
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "9765.62"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   7140
         TabIndex        =   97
         Top             =   3590
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "1420307004.84"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   7140
         TabIndex        =   96
         Top             =   3160
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "0.567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   7140
         TabIndex        =   95
         Top             =   2730
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "+28 deg 34 min 57 sec DEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   7140
         TabIndex        =   94
         Top             =   2205
         Width           =   2535
      End
      Begin VB.Label Label39 
         Caption         =   "23 hr 03 min 53 dec RA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   7140
         TabIndex        =   93
         Top             =   1715
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "+28 deg 34 min 57 sec DEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   7140
         TabIndex        =   92
         Top             =   1230
         Width           =   2640
      End
      Begin VB.Label Label39 
         Caption         =   "23 hr 03 min 53 dec RA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   7140
         TabIndex        =   91
         Top             =   735
         Width           =   2325
      End
      Begin VB.Label Label39 
         Caption         =   "Образец"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   7770
         TabIndex        =   90
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Subband number"
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   81
         Top             =   4020
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Рабочий диапазон (subband sample rate), Гц"
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   80
         Top             =   3590
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Центр частотного диапазона (subband center), Гц"
         Height          =   435
         Index           =   5
         Left            =   210
         TabIndex        =   79
         Top             =   3055
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Угол обзора (angle range)"
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   78
         Top             =   2730
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Конечное склонение (EndDEC)"
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   77
         Top             =   2205
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Конечный угол (EndRA)"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   76
         Top             =   1715
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Начальное склонение (StartDEC) - показывается клиентом SETI@home"
         Height          =   435
         Index           =   1
         Left            =   210
         TabIndex        =   75
         Top             =   1120
         Width           =   3375
      End
      Begin VB.Label Label39 
         Caption         =   "Стартовый угол (StartRA) - показывается клиентом SETI@home"
         Height          =   435
         Index           =   0
         Left            =   210
         TabIndex        =   74
         Top             =   630
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Общие сведения об обрабатываемом блоке"
      Height          =   4530
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   630
      Width           =   9990
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3885
         TabIndex        =   71
         Top             =   3328
         Width           =   2640
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3885
         TabIndex        =   67
         Top             =   2595
         Width           =   2640
      End
      Begin VB.OptionButton Option1 
         Height          =   225
         Index           =   1
         Left            =   3360
         TabIndex        =   66
         Top             =   2625
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         Height          =   225
         Index           =   0
         Left            =   3360
         TabIndex        =   65
         Top             =   2205
         Width           =   225
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3885
         TabIndex        =   12
         Text            =   "ao1420"
         Top             =   4013
         Width           =   2640
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3885
         TabIndex        =   10
         Top             =   2175
         Width           =   2640
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3885
         TabIndex        =   8
         Top             =   1469
         Width           =   2640
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3885
         TabIndex        =   6
         Top             =   787
         Width           =   2640
      End
      Begin VB.Label Label38 
         Caption         =   "ao1420"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   73
         Top             =   4043
         Width           =   2745
      End
      Begin VB.Label Label37 
         Caption         =   "1420302732.38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   72
         Top             =   3358
         Width           =   2745
      End
      Begin VB.Label Label36 
         Caption         =   "Частота (subband base), Гц"
         Height          =   225
         Left            =   315
         TabIndex        =   70
         Top             =   3358
         Width           =   2535
      End
      Begin VB.Label Label35 
         Caption         =   "Wed Sep 01 21:02:20 1999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   69
         Top             =   2625
         Width           =   2745
      End
      Begin VB.Label Label34 
         Caption         =   "2451423.37662"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   68
         Top             =   2205
         Width           =   2850
      End
      Begin VB.Label Label33 
         Caption         =   "01se99aa.12684.28177.692336.31"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   64
         Top             =   1522
         Width           =   2955
      End
      Begin VB.Label Label32 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   63
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label31 
         Caption         =   "Образец"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8085
         TabIndex        =   62
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Телескоп (aо1420 для Аресибо)"
         Height          =   225
         Left            =   315
         TabIndex        =   11
         Top             =   4043
         Width           =   2640
      End
      Begin VB.Label Label4 
         Caption         =   "Время записи информации"
         Height          =   225
         Left            =   315
         TabIndex        =   9
         Top             =   2205
         Width           =   2850
      End
      Begin VB.Label Label3 
         Caption         =   "Имя блока"
         Height          =   225
         Left            =   315
         TabIndex        =   7
         Top             =   1522
         Width           =   2850
      End
      Begin VB.Label Label2 
         Caption         =   "Идентификационный номер записи"
         Height          =   225
         Left            =   315
         TabIndex        =   5
         Top             =   840
         Width           =   3060
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   8715
      TabIndex        =   1
      Top             =   5460
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   7140
      TabIndex        =   0
      Top             =   5460
      Width           =   1380
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5160
      Left            =   105
      TabIndex        =   3
      Top             =   150
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9102
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Общие сведения"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Физические параметры"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Координаты"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Дополнительные сведения"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Комментарии"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line10 
      X1              =   105
      X2              =   6825
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   6825
      X2              =   6825
      Y1              =   5460
      Y2              =   5985
   End
   Begin VB.Line Line8 
      X1              =   105
      X2              =   105
      Y1              =   5460
      Y2              =   5985
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   105
      X2              =   6825
      Y1              =   5985
      Y2              =   5985
   End
   Begin VB.Label Label1 
      Height          =   540
      Left            =   105
      TabIndex        =   2
      Top             =   5460
      Width           =   6735
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
    'TO DO validate and save changes
    'Unload Me
'Exit Sub    'ЗАГЛУШКА!!!
    If WU.CheckUnit(0, Text1.text) Then
        If NewMode Then
            Result = MsgBox("Запись с этим номером уже существует! Хотите ли Вы" + vbCrLf + "перезаписать новую информацию повех старой?", vbYesNo + vbExclamation, "Запись уже существует")
            If Result = vbYes Then
                EditMode = True
                NewMode = False
            Else
                frmEdit.Text1.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    Else
        If EditMode Then
            Result = MsgBox("Записи с этим номером нет в журнале! Хотите ли Вы" + vbCrLf + "записать новую информацию вместо редактирования старой?", vbYesNo + vbExclamation, "Новая запись")
            If Result = vbYes Then
                EditMode = False
                NewMode = True
            Else
                frmEdit.Text1.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    End If
    If WU.CheckUnit(1, Text2.text) Then
        If NewMode Then
            Result = MsgBox("Запись с таким именем уже существует! Хотите ли Вы" + vbCrLf + "перезаписать новую информацию повех старой?", vbYesNo + vbExclamation, "Запись уже существует")
            If Result = vbYes Then
                EditMode = True
                NewMode = False
            Else
                frmEdit.Text2.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    Else
        If EditMode Then
            Result = MsgBox("Записи с таким именем нет в журнале! Хотите ли Вы" + vbCrLf + "записать новую информацию вместо редактирования старой?", vbYesNo + vbExclamation, "Новая запись")
            If Result = vbYes Then
                EditMode = False
                NewMode = True
            Else
                frmEdit.Text2.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    End If
    WU.NumID = Text1.text
    WU.UnitName = Text2.text
    If Option2(0).Value Then
        'Отдельные поля -> нужно преобразовывать
        WU.StartRA = EncodeCoor(Text9(0).text, Text10(0).text, Text11(0).text)
        WU.StartDEC = EncodeCoor(Text9(1).text, Text10(1).text, Text11(1).text)
        WU.EndRA = EncodeCoor(Text9(2).text, Text10(2).text, Text11(2).text)
        WU.EndDEC = EncodeCoor(Text9(3).text, Text10(3).text, Text11(3).text)
    Else
        'Координаты уже представлены в нужном виде
        WU.StartRA = Text8(0).text
        WU.StartDEC = Text8(1).text
        WU.EndRA = Text8(2).text
        WU.EndDEC = Text8(3).text
    End If
    If Option1(1).Value Then
        If Not Text3.text = "" Then
            WU.TimeOfRec = Text3.text
            If Right(WU.TimeOfRec, 1) = " " Then
                WU.TimeOfRec = Left(WU.TimeOfRec, Len(WU.TimeOfRec) - 1)
            End If
        Else
            WU.TimeOfRec = "0.0"
        End If
        WU.TimeOfRec = WU.TimeOfRec & " (" & Text6.text & ")"
    End If
    With WU
        .Type_of_unit = Text12(0).text
        .Task = Text12(1).text
        .Version = Text12(2).text
        .Data_type = Text12(3).text
        .Data_class = Text12(4).text
        .Splitter_version = Text12(5).text
        .AngleRange = Text8(4).text
        .SubbandCenter = Text8(5).text
        .SubbandBase = Text7.text
        .SubbandRate = Text8(6).text
        .fft_len = Text12(6).text
        .ifft_len = Text12(7).text
        .SubbandNum = Text8(7).text
        .Receiver = Text4.text
        .Nsamples = Text12(8).text
        .TapeVer = Text12(9).text
        .Comments = Text13.text
        .Reserve1 = Text14.text
        .Reserve2 = Text15.text
        .Reserve3 = Text16.text
    End With
    If Not (WU.NumPositions = "") Then
        'Этот параметр неизменяемый. Если все-таки это поле пусто, то присвоить ему
        'значение по-умолчанию
        WU.NumPositions = "22"
    End If
    For i = 0 To 21
        If Not (WU.SetCoordX(Text5(i).text, i)) Then
            Debug.Print "ERROR has occured while trying to write to Coords(i)"
        End If
    Next i
    'Saving results
    If EditMode Then
    'TO DO  сделать refresh для DataGrid1, а, может даже Rebind
        If Not (WU.WriteHistory(WU.EncodeWU, 0)) Then
            Result = MsgBox("Ошибка записи в журнал!", vbOKOnly, "Ошибка")
        End If
        'Нужно восстановить испорченные номер и имя блока
        WU.NumID = Text1.text
        WU.UnitName = Text2.text
        WU.AddRecord 2, CLng(Val(WU.NumID))
    'Это не совсем то, что имелось в виду чуть выше (не Rebind)
    ''frmHistory.DataGrid1.ReBind
    ''frmHistory.DataGrid1.Refresh
    Else
    'Новая запись
        i = WU.GetLastNum
        'Нужно восстановить испорченные номер и имя блока
        WU.NumID = Text1.text
        WU.UnitName = Text2.text
        If CDbl(Val(WU.NumID)) > i Then
            'В конец файла
            If Not (WU.WriteHistory(WU.EncodeWU, 1)) Then
                Result = MsgBox("Ошибка записи в журнал!", vbOKOnly, "Ошибка")
                'WriteHistory сам знает когда изменять значения RegRecords и LastRecordNum
                ''LastRecordNum = LastRecordNum + 1
            End If
        Else
            'Вставка в середину журнала
            If Not (WU.WriteHistory(WU.EncodeWU, 0)) Then
                Result = MsgBox("Ошибка записи в журнал!", vbOKOnly, "Ошибка")
            End If
        End If
        WU.AddRecord 1, CLng(Val(Text1.text))
        'WriteHistory сам знает когда изменять значения RegRecords и LastRecordNum
        ''RegRecords = RegRecords + 1
        frmHistory.DataGrid1.Refresh
    End If
    frmHistory.Command3.Enabled = False
    Unload Me
End Sub

Private Sub Command2_Click()
    frmHistory.Command3.Enabled = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    i = TabStrip1.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = TabStrip1.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
        Else
            'increment the tab
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.Tabs.Count)
        Else
            'increment the tab
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(i - 1)
        End If
    End If
End Sub

Private Sub Form_Load()
Dim fColor As Long, bColor As Long
Dim i As Long
'Если EditMode=True, то все известные поля уже заполнены
'Если же нет, то известны только ID и Name
    fColor = Label1.ForeColor
    bColor = Label1.BackColor
    Label1.ForeColor = vbYellow
    Label1.BackColor = vbBlack
    Label1.Caption = "ВНИМАНИЕ! Программа НЕ ОСУЩЕСТВЛЯЕТ контроль за соответствием вводимой информации приведенным образцам! Последствия ошибки могут быть необратимы!"
    ''Label1.ForeColor = fColor
    ''Label1.BackColor = bColor
    
    Text1.text = WU.NumID
    Text2.text = WU.UnitName
    'ВСЕ! На этом кончается ТОЧНО известная информация!
    'Есть ли что-то еще - не известно! По мере сил можно будет довесить проверки...
    'TO DO
    
If EditMode Then
    If Not (WU.TimeOfRec = "") Then
        Text6.text = ExtractTime(WU.TimeOfRec)
        Text3.text = ExtractDigTime(WU.TimeOfRec)
    End If
    If Not (WU.SubbandBase = "") Then
        Text7.text = WU.SubbandBase
    End If
    ''If Not (WU.StartRA = "") Then
        ''Text9(0).text =
    ''End If
    Text9(0).text = GetHourStr(CDbl(Val(WU.StartRA)))
    Text10(0).text = GetMinStr(CDbl(Val(WU.StartRA)))
    Text11(0).text = GetSecStr(CDbl(Val(WU.StartRA)))
    Text9(1).text = GetHourStr(CDbl(Val(WU.StartDEC)))
    Text10(1).text = GetMinStr(CDbl(Val(WU.StartDEC)))
    Text11(1).text = GetSecStr(CDbl(Val(WU.StartDEC)))
    ''См. Option2_click
    ''Text8(0).text = WU.StartRA
    ''Text8(1).text = WU.StartDEC
    If Not (WU.EndRA = "") Then
        Text9(2).text = GetHourStr(CDbl(Val(WU.EndRA)))
        Text10(2).text = GetMinStr(CDbl(Val(WU.EndRA)))
        Text11(2).text = GetSecStr(CDbl(Val(WU.EndRA)))
        Text8(2).text = WU.EndRA
    End If
    If Not (WU.EndDEC = "") Then
        Text9(3).text = GetHourStr(CDbl(Val(WU.EndDEC)))
        Text10(3).text = GetMinStr(CDbl(Val(WU.EndDEC)))
        Text11(3).text = GetSecStr(CDbl(Val(WU.EndDEC)))
        Text8(3).text = WU.EndDEC
    End If
    If Not (WU.AngleRange = "") Then
        Text8(4).text = WU.AngleRange
    End If
    If Not (WU.SubbandCenter = "") Then
        Text8(5).text = WU.SubbandCenter
    End If
    If Not (WU.SubbandRate = "") Then
        Text8(6).text = WU.SubbandRate
    End If
    If Not (WU.SubbandNum = "") Then
        Text8(7).text = WU.SubbandNum
    End If
'Закладка 4
    If Not (WU.Type_of_unit = "") Then
        Text12(0).text = WU.Type_of_unit
    End If
    If Not (WU.Task = "") Then
        Text12(1).text = WU.Task
    End If
    If Not (WU.Version = "") Then
        Text12(2).text = WU.Version
    End If
    If Not (WU.Data_type = "") Then
        Text12(3).text = WU.Data_type
    End If
    If Not (WU.Data_class = "") Then
        Text12(4).text = WU.Data_class
    End If
    If Not (WU.Splitter_version = "") Then
        Text12(5).text = WU.Splitter_version
    End If
    If Not (WU.fft_len = "") Then
        Text12(6).text = WU.fft_len
    End If
    If Not (WU.ifft_len = "") Then
        Text12(7).text = WU.ifft_len
    End If
    If Not (WU.Nsamples = "") Then
        Text12(8).text = WU.Nsamples
    End If
    If Not (WU.TapeVer = "") Then
        Text12(9).text = WU.TapeVer
    End If
    For i = 0 To 21
        If Not (WU.GetCoordX(i) = "") Then
            Text5(i).text = WU.GetCoordX(i)
        End If
    Next i
    If Not (WU.Comments = "") Then
        Text13.text = WU.Comments
    End If
    If Not (WU.Reserve1 = "") Then
        Text14.text = WU.Reserve1
    End If
    If Not (WU.Reserve2 = "") Then
        Text15.text = WU.Reserve2
    End If
    If Not (WU.Reserve3 = "") Then
        Text16.text = WU.Reserve3
    End If
    If Not (WU.Receiver = "") Then
        Text4.text = WU.Receiver
    End If
    Text14.Enabled = False
    Text15.Enabled = False
    Text16.Enabled = False
End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Text3.Enabled = True
            Text6.Enabled = False
        Case 1:
            Text3.Enabled = False
            Text6.Enabled = True
    End Select
End Sub

Private Sub Option2_Click(Index As Integer)
Dim i As Long
    Select Case Index
        Case 0: 'Отдельные поля для градусов, минут, секунд
            For i = 0 To 3
                Text8(i).Visible = False
                Text9(i).Visible = True
                Text10(i).Visible = True
                Text11(i).Visible = True
                Text9(i).text = GetHourStr(CDbl(Val(Text8(i).text)))
                Text10(i).text = GetMinStr(CDbl(Val(Text8(i).text)))
                Text11(i).text = GetSecStr(CDbl(Val(Text8(i).text)))
            Next i
            For i = 0 To 11
                Label40(i).Visible = True
            Next i
            Label39(9) = "23 hr 03 min 53 dec RA"
            Label39(10) = "+28 deg 34 min 57 sec DEC"
            Label39(11) = "23 hr 03 min 53 dec RA"
            Label39(12) = "+28 deg 34 min 57 sec DEC"
        Case 1:
            For i = 0 To 3
                Text9(i).Visible = False
                Text10(i).Visible = False
                Text11(i).Visible = False
                Text8(i).Visible = True
                Text8(i).text = EncodeCoor(Text9(i).text, Text10(i).text, Text11(i).text)
            Next i
            For i = 0 To 11
                Label40(i).Visible = False
            Next i
            Label39(9) = "15.814"
            Label39(10) = "11.39"
            Label39(11) = "15.852"
            Label39(12) = "11.52"
            
    End Select
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To TabStrip1.Tabs.Count - 1
        If i = TabStrip1.SelectedItem.Index - 1 Then
            Frame1(i).Left = 210
            Frame1(i).Enabled = True
        Else
            Frame1(i).Left = -20000
            Frame1(i).Enabled = False
        End If
    Next
End Sub
