VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Графический анализ результатов"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Ручное указание диапазона"
      Height          =   330
      Left            =   8715
      TabIndex        =   13
      Top             =   2415
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   8715
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1e-4
      Max             =   100
      Scrolling       =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Обновить индекс"
      Height          =   435
      Left            =   8925
      TabIndex        =   10
      Top             =   105
      Width           =   1275
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   225
      Left            =   8610
      TabIndex        =   8
      Top             =   1680
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   5
      SelStart        =   5
      Value           =   5
      TextPosition    =   1
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5145
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   165
      Width           =   3270
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Начать!"
      Height          =   345
      Left            =   8820
      TabIndex        =   3
      Top             =   735
      Width           =   1467
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Default         =   -1  'True
      Height          =   345
      Left            =   8820
      TabIndex        =   4
      Top             =   5565
      Width           =   1467
   End
   Begin VB.PictureBox Picture1 
      Height          =   5370
      Left            =   105
      ScaleHeight     =   5310
      ScaleWidth      =   8355
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   630
      Width           =   8415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   165
      Width           =   1905
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ПРАВАЯ кнопка - конец диапазона"
      Height          =   435
      Left            =   8715
      TabIndex        =   15
      Top             =   3465
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "ЛЕВАЯ кнопка - начало диапазона"
      Height          =   435
      Left            =   8715
      TabIndex        =   14
      Top             =   2940
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Построение..."
      Height          =   225
      Left            =   8715
      TabIndex        =   12
      Top             =   4725
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   8715
      TabIndex        =   9
      Top             =   1995
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   8715
      TabIndex        =   7
      Top             =   1365
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Тип графика"
      Height          =   225
      Left            =   4095
      TabIndex        =   6
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Категория данных"
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1590
   End
End
Attribute VB_Name = "frmGport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Mode As Long     'Какой режим выбран (тип графика)
Public BandWith As Long    'Ширина полосы частот

Private Sub Check2_Click()
    'Если отметка поставлена, то погасить все, а подсказки высветить.
    'Если отметка погашена, то считается, что работа с этим графиком закончена.
    Call DisableControls
    Combo2.Enabled = False
    If Check2.Value = vbChecked Then
        Combo1.Enabled = False
        Label6.Visible = True
        Label7.Visible = True
    Else
        Combo1.Enabled = True
        Label6.Visible = False
        Label7.Visible = False
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Check2.Value = vbChecked Then
        Select Case Button
            Case 0: 'Левая кнопка
                Result = MsgBox("Левая кнопка! " & "X=" & Str(X) & " Y=" & Str(Y), vbOKOnly, "Ого!")
            Case 1: 'Правая кнопка
                Result = MsgBox("Правая кнопка! " & "X=" & Str(X) & " Y=" & Str(Y), vbOKOnly, "Ого!")
        End Select
    End If
End Sub

Private Sub Slider1_Change()
    Select Case Mode
        Case 101:   'Блоки -> частоты
            Select Case Slider1.Value
                Case 1:
                    BandWith = 25
                    Label4.Caption = "25 КГц"
                Case 2:
                    BandWith = 50
                    Label4.Caption = "50 КГц"
                Case 3:
                    BandWith = 100
                    Label4.Caption = "100 КГц"
                Case 4:
                    BandWith = 125
                    Label4.Caption = "125 КГц"
                Case 5:
                    BandWith = 250
                    Label4.Caption = "250 КГц"
                Case 6:
                    BandWith = 500
                    Label4.Caption = "500 КГц"
                Case 7:
                    BandWith = 1250
                    Label4.Caption = "1250 КГц"
            End Select
        Case 102:   'Блоки -> RA
            Select Case Slider1.Value
                Case 1:
                    BandWith = 1
                    Label4.Caption = "1 час"
                Case 2:
                    BandWith = 2
                    Label4.Caption = "2 часа"
                Case 3:
                    BandWith = 3
                    Label4.Caption = "3 часа"
                Case 4:
                    BandWith = 4
                    Label4.Caption = "4 часа"
                Case 5:
                    BandWith = 6
                    Label4.Caption = "6 часов"
            End Select
        Case 103:   'Блоки -> DEC
            Select Case Slider1.Value
                Case 1:
                    BandWith = 2
                    Label4.Caption = "2 градуса"
                Case 2:
                    BandWith = 5
                    Label4.Caption = "5 градусов"
                Case 3:
                    BandWith = 10
                    Label4.Caption = "10 градусов"
                Case 4:
                    BandWith = 25
                    Label4.Caption = "25 градусов"
            End Select
        Case 201:   'Пики -> Мощность
            Select Case Slider1.Value
                Case 1:
                    BandWith = 6
                    Label4.Caption = "6"
                Case 2:
                    BandWith = 10
                    Label4.Caption = "10"
                Case 3:
                    BandWith = 25
                    Label4.Caption = "25"
                Case 4:
                    BandWith = 50
                    Label4.Caption = "50"
                Case 5:
                    BandWith = 100
                    Label4.Caption = "100"
            End Select
    End Select
End Sub

'*******************************************
'*       Выбор категории графиков          *
'*******************************************
Private Sub Combo1_Click()
    'Разрешить выбор подпунктов из второго списка
    Combo2.Enabled = True
    'Очистить список перед его заполнением
    Combo2.Clear
    'Запретить построения графиков до выбора подпункта из второго списка
    Command2.Enabled = False
    'TO DO в зависимости от категории наполнить Combo2 пунктами...
    If Combo1.text = "Рабочие блоки" Then
        Combo2.AddItem "Распределение по частотам", 0
        Combo2.AddItem "Угловые координаты", 1
        Combo2.AddItem "Склонение", 2
    ElseIf Combo1.text = "Пики" Then
        Combo2.AddItem "Распределение по мощности", 0
        Combo2.AddItem "Мощность - сдвиг допплера", 1
        Combo2.AddItem "Мощность - несущая частота", 2
    End If
End Sub

'*******************************************
'*        Определение типа графика         *
'*******************************************
Private Sub Combo2_Click()
    If Combo1.text = "Рабочие блоки" Then
        If Combo2.text = "Распределение по частотам" Then
            Mode = 101
        ElseIf Combo2.text = "Угловые координаты" Then
            Mode = 102
        ElseIf Combo2.text = "Склонение" Then
            Mode = 103
        End If
    ElseIf Combo1.text = "Пики" Then
        If Combo2.text = "Распределение по мощности" Then
            Mode = 201
        ElseIf Combo2.text = "Мощность - сдвиг допплера" Then
            Mode = 202
        ElseIf Combo2.text = "Мощность - несущая частота" Then
            Mode = 203
        End If
    End If
    Call ChangeControls     'Подготовить настройки для данного типа графика
    Command2.Enabled = True 'Разрешить построение графика
End Sub

Private Sub ChangeControls()
    Call DisableControls
    Select Case Mode
        Case 101:
            Label3.Caption = "Ширина столбцов"
            Label3.Visible = True
            Slider1.min = 1
            Slider1.Max = 7
            Slider1.SmallChange = 1
            Slider1.LargeChange = 1
            Slider1.Value = 6
            Slider1.Visible = True
            Label4.Caption = "500 кГц"
            Label4.Visible = True
            BandWith = 500      'На случай, если установка по-умолчанию не изменится
                                '(чтобы не было деления на ноль)
        Case 102:
            Label3.Caption = "Ширина столбцов"
            Label3.Visible = True
            Slider1.min = 1
            Slider1.Max = 5
            Slider1.SmallChange = 1
            Slider1.LargeChange = 1
            Slider1.Value = 5
            Slider1.Visible = True
            Label4.Caption = "6 часов"
            Label4.Visible = True
            BandWith = 6        'На случай, если установка по-умолчанию не изменится
                                '(чтобы не было деления на ноль)
        Case 103:
            Label3.Caption = "Ширина столбцов"
            Label3.Visible = True
            Slider1.min = 1
            Slider1.Max = 4
            Slider1.SmallChange = 1
            Slider1.LargeChange = 1
            Slider1.Value = 4
            Slider1.Visible = True
            Label4.Caption = "25 градусов"
            Label4.Visible = True
            BandWith = 25       'На случай, если установка по-умолчанию не изменится
                                '(чтобы не было деления на ноль)
        Case 201:
            Label3.Caption = "Ширина столбцов"
            Label3.Visible = True
            Slider1.min = 1
            Slider1.Max = 5
            Slider1.SmallChange = 1
            Slider1.LargeChange = 1
            Slider1.Value = 4
            Slider1.Visible = True
            Label4.Caption = "50"
            Label4.Visible = True
            BandWith = 50       'На случай, если установка по-умолчанию не изменится
                                '(чтобы не было деления на ноль)
    End Select
End Sub

'Дезактивировать все объекты и подготовиться к новым командам
Private Sub DisableControls()
    Label3.Visible = False
    Label4.Visible = False
    Slider1.Visible = False
    Label7.Visible = False
    Label6.Visible = False
    Check2.Visible = False
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Check1.Value = vbChecked Then
        'Построить индекс блоков заново
        WU.WriteIndex (1)
        Check1.Value = vbUnchecked
    End If
    'TO DO Определить текущий режим
    Call DrawGraph
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Caption = "Графический анализ"
    Picture1.BackColor = vbBlack
    Combo1.AddItem "Рабочие блоки", 0
    Combo1.AddItem "Пики", 1
    Combo1.AddItem "Гауссианы", 2
    Combo1.AddItem "Импульсы", 3
    Combo1.AddItem "Триплеты", 4
    Combo1.AddItem "Другие", 5
    Combo2.Enabled = False
    'Пока не выбран тип графика - строить нельзя
    Command2.Enabled = False
End Sub

Private Sub DrawGraph()
Dim i As Long, j As Long
Dim Hits As Long
Dim TMPvalue As Long
Dim BandNums As Long    'Число столбцов
Dim OnePercent As Long  'Один процент (для индикатора прогресса)

    AutoRedraw = -1   ' Turn on AutoRedraw.
    Select Case Mode
        Case 101:   '->  Блоки->частоты
            Picture1.Scale (0, 110)-(140, 0)    ' Set custom coordinate system.
                                                'По 20 единиц про запас справа и слева
            Picture1.ForeColor = vbWhite
            Picture1.Cls
            For i = 100 To 10 Step -10
                Picture1.Line (0, i)-(2, i)     ' Draw scale marks every 10 units.
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.Print i  ' Print scale mark value on left.
                Picture1.Line (Picture1.ScaleWidth - 2, i)-(Picture1.ScaleWidth, i)
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.CurrentX = Picture1.ScaleWidth - 9
                Picture1.Print i  ' Print scale mark value on right.
            Next i
            Picture1.Line (20, 110)-(20, 10), RGB(255, 255, 255)
            Picture1.Line (70, 110)-(70, 10), RGB(255, 255, 255)
            Picture1.Line (120, 110)-(120, 10), RGB(255, 255, 255)
            Picture1.CurrentX = 5
            Picture1.CurrentY = 110
            Picture1.Print "1418,75 MHz"  ' Print scale mark value on left.
            Picture1.CurrentX = 120
            Picture1.CurrentY = 110
            Picture1.Print "1421,25 MHz"  ' Print scale mark value on left.
            
            
            BandNums = 2500 \ BandWith
            For j = 0 To BandNums - 1       'Столько проходов, сколько столбцов
                Hits = 0 'Обнулить счетчик совпадений
                For i = 1 To RegRecords     'Проверить все записи
                    WU.ReadIndex (i)
                    TMPvalue = CLng(Trunc(CDbl(TopW.freq)))
                    If TMPvalue >= (1418750000 + BandWith * j * 1000) Then
                        If TMPvalue < (1418750000 + BandWith * (j + 1) * 1000) Then
                            Hits = Hits + 1
                        End If
                    End If
                Next i
                If j / 2 = j \ 2 Then
                    Picture1.Line (20 + j * (100 \ BandNums), 0)-(20 + (j + 1) * (100 \ BandNums), Hits), RGB(0, 0, 255), BF ' blue bar.
                Else
                    Picture1.Line (20 + j * (100 \ BandNums), 0)-(20 + (j + 1) * (100 \ BandNums), Hits), RGB(255, 0, 0), BF ' red bar.
                End If
            Next j
        Case 102:   '->  Блоки->RA
            Picture1.Scale (0, 110)-(160, 0)        ' Set custom coordinate system.
                                                    'По 20 единиц про запас справа и слева
            Picture1.ForeColor = vbWhite
            Picture1.Cls
            For i = 100 To 10 Step -10
                Picture1.Line (0, i)-(2, i)     ' Draw scale marks every 10 units.
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.Print i  ' Print scale mark value on left.
                Picture1.Line (Picture1.ScaleWidth - 2, i)-(Picture1.ScaleWidth, i)
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.CurrentX = Picture1.ScaleWidth - 9
                Picture1.Print i  ' Print scale mark value on right.
            Next i
            Picture1.Line (20, 110)-(20, 10), RGB(255, 255, 255)
            Picture1.Line (80, 110)-(80, 10), RGB(255, 255, 255)
            Picture1.Line (140, 110)-(140, 10), RGB(255, 255, 255)
            Picture1.CurrentX = 12
            Picture1.CurrentY = 110
            Picture1.Print "0 RA"  ' Print scale mark value on left.
            Picture1.CurrentX = 152
            Picture1.CurrentY = 110
            Picture1.Print "24 RA"  ' Print scale mark value on left.
                        
            BandNums = 24 \ BandWith
            For j = 0 To BandNums - 1       'Столько проходов, сколько столбцов
                Hits = 0 'Обнулить счетчик совпадений
                For i = 1 To RegRecords     'Проверить все записи
                    WU.ReadIndex (i)
                    TMPvalue = CLng(Trunc(CDbl(TopW.StartRA)))
                    If TMPvalue >= (BandWith * j) Then
                        If TMPvalue < (BandWith * (j + 1)) Then
                            Hits = Hits + 1
                        End If
                    End If
                Next i
                If j / 2 = j \ 2 Then
                    Picture1.Line (20 + j * (120 \ BandNums), 0)-(20 + (j + 1) * (120 \ BandNums), Hits), RGB(0, 0, 255), BF ' blue bar.
                Else
                    Picture1.Line (20 + j * (120 \ BandNums), 0)-(20 + (j + 1) * (120 \ BandNums), Hits), RGB(255, 0, 0), BF ' red bar.
                End If
            Next j
            
        Case 103:   '->  Блоки->склонение
            Picture1.Scale (0, 110)-(140, 0)    ' Set custom coordinate system.
                                                'По 20 единиц про запас справа и слева
            Picture1.ForeColor = vbWhite
            Picture1.Cls
            For i = 100 To 10 Step -10
                Picture1.Line (0, i)-(2, i)     ' Draw scale marks every 10 units.
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.Print i  ' Print scale mark value on left.
                Picture1.Line (Picture1.ScaleWidth - 2, i)-(Picture1.ScaleWidth, i)
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.CurrentX = Picture1.ScaleWidth - 9
                Picture1.Print i  ' Print scale mark value on right.
            Next i
            Picture1.Line (20, 110)-(20, 10), RGB(255, 255, 255)
            Picture1.Line (70, 110)-(70, 10), RGB(255, 255, 255)
            Picture1.Line (120, 110)-(120, 10), RGB(255, 255, 255)
            Picture1.CurrentX = 5
            Picture1.CurrentY = 110
            Picture1.Print "-5 DEC"  ' Print scale mark value on left.
            Picture1.CurrentX = 120
            Picture1.CurrentY = 110
            Picture1.Print "+45 DEC"  ' Print scale mark value on left.
            
            BandNums = 50 \ BandWith
            For j = 0 To BandNums - 1       'Столько проходов, сколько столбцов
                Hits = 0 'Обнулить счетчик совпадений
                For i = 1 To RegRecords     'Проверить все записи
                    WU.ReadIndex (i)
                    TMPvalue = CLng(Trunc(CDbl(TopW.StartDEC)))
                    If TMPvalue >= (BandWith * j) Then
                        If TMPvalue < (BandWith * (j + 1)) Then
                            Hits = Hits + 1
                        End If
                    End If
                Next i
                If j / 2 = j \ 2 Then
                    Picture1.Line (20 + j * (100 \ BandNums), 0)-(20 + (j + 1) * (100 \ BandNums), Hits), RGB(0, 0, 255), BF ' blue bar.
                Else
                    Picture1.Line (20 + j * (100 \ BandNums), 0)-(20 + (j + 1) * (100 \ BandNums), Hits), RGB(255, 0, 0), BF ' red bar.
                End If
            Next j
        Case 201:   'Пики -> мощность
            Picture1.Scale (0, 110)-(140, 0)    ' Set custom coordinate system.
                                                'По 20 единиц про запас справа и слева
            Picture1.ForeColor = vbWhite
            Picture1.Cls
            For i = 100 To 10 Step -10
                Picture1.Line (0, i)-(2, i)     ' Draw scale marks every 10 units.
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.Print i  ' Print scale mark value on left.
                Picture1.Line (Picture1.ScaleWidth - 2, i)-(Picture1.ScaleWidth, i)
                Picture1.CurrentY = Picture1.CurrentY + 1.5   ' Move cursor position.
                Picture1.CurrentX = Picture1.ScaleWidth - 9
                Picture1.Print i  ' Print scale mark value on right.
            Next i
            Picture1.Line (20, 110)-(20, 10), RGB(255, 255, 255)
            Picture1.Line (70, 110)-(70, 10), RGB(255, 255, 255)
            Picture1.Line (120, 110)-(120, 10), RGB(255, 255, 255)
            Picture1.CurrentX = 1
            Picture1.CurrentY = 110
            Picture1.Print "Мощность 0"  ' Print scale mark value on left.
            Picture1.CurrentX = 121
            Picture1.CurrentY = 110
            Picture1.Print "Мощность 600"  ' Print scale mark value on left.
            
            Label5.Visible = True
            Label5.Refresh
            ProgressBar1.Visible = True
            
            BandNums = 600 \ BandWith
            For j = 0 To BandNums - 1       'Столько проходов, сколько столбцов
                Hits = 0 'Обнулить счетчик совпадений
                For i = 1 To RegRecords     'Проверить все записи
                    If State.ReadIndex(0, i) Then
                    'Чтение произведено успешно
                        TMPvalue = CLng(Trunc(CDbl(TopS.power)))
                        If TMPvalue >= (BandWith * j) Then
                            If TMPvalue < (BandWith * (j + 1)) Then
                                Hits = Hits + 1
                            End If
                        End If
                    End If
                Next i
                If j / 2 = j \ 2 Then
                    Picture1.Line (20 + CLng(Trunc(CDbl(j * (100 / BandNums)))), 0)-(20 + CLng(Trunc(CDbl((j + 1) * (100 / BandNums)))), Hits), RGB(0, 0, 255), BF ' blue bar.
                Else
                    Picture1.Line (20 + CLng(Trunc(CDbl(j * (100 / BandNums)))), 0)-(20 + CLng(Trunc(CDbl((j + 1) * (100 / BandNums)))), Hits), RGB(255, 0, 0), BF ' red bar.
                End If
                ProgressBar1.Value = Trunc(CDbl((j + 1) / BandNums * 99)) + 1
                ProgressBar1.Refresh
            Next j
            
            ProgressBar1.Visible = False
            Label5.Visible = False
    End Select
End Sub
