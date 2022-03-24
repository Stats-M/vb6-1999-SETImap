VERSION 5.00
Begin VB.Form frmMport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Теоретический анализ"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3 
      Caption         =   "Справка"
      Height          =   345
      HelpContextID   =   15001
      Left            =   5460
      TabIndex        =   3
      Top             =   2520
      Width           =   1467
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Закрыть"
      Height          =   345
      HelpContextID   =   15002
      Left            =   5460
      TabIndex        =   2
      Top             =   3045
      Width           =   1467
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      Height          =   3270
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   105
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Начать !"
      Default         =   -1  'True
      Height          =   345
      HelpContextID   =   15000
      Left            =   5460
      TabIndex        =   0
      Top             =   210
      Width           =   1467
   End
End
Attribute VB_Name = "frmMport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim STRout As String       'Сообщения программы (впоследствии перенести в ресурс)

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Text1.text = ""
    Call PosInIndex
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Caption = "Аналитическая информация"
End Sub

Private Sub PosInIndex()
Dim Place As Long   'Временная переменная для хранения места результата в файле индекса
Dim IndexLen As Long    'Количество записей в индексе
Dim i As Long
Dim Summa As Single     'Сумма ненулевых гауссиан
Dim NumValid As Long    'Количество ненулевых гауссиан
    IndexLen = State.GetLastRecNum(4)  'Работа с индексом пиков
    Place = -1
    For i = 1 To IndexLen
        bResult = State.ReadIndex(0, i)
        'TO DO
        'Сделать здесь проверку Win/Lin...
        If TopS.ID = WinID Then
            Place = i       'Запомнить место
            i = IndexLen    'Завершить цикл
        End If
    Next i
    
    If Place = -1 Then      'Запись не найдена!
        STRout = "Данные о лучших пиковых сигналах (рабочий блок " & Str(WinID) & ") не обнаружены. "
        STRout = STRout & "Возможно, Вы еще не запускали обработчик этой информации. "
        STRout = STRout & "Выберите пункт меню Статистика/Лучшие пики или Статистика/Лучшие гауссианы. "
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    Else
        STRout = "Пиковый сигнал данного рабочего блока находится на " & Place & " месте среди" & Str(i - 1) & " обработанных за все время. " & vbCrLf & vbCrLf
        If TopS.power > 400 Then
            STRout = STRout & "Чрезвычайно высокая мощность сигнала. Вероятно, это тестовый сигнал, подмешиваемый в некоторые рабочие блоки для контроля корректности работы клиентов. "
        ElseIf TopS.power > 250 Then
            STRout = STRout & "Очень высокая мощность сигнала, скорее всего, во время записи этого блока антенна была нацелена точно на искусственный спутник Земли. "
        ElseIf TopS.power > 200 Then
            STRout = STRout & "Высокая мощность сигнала, возможно, это результат излучения какого-либо орбитального аппарата. "
        ElseIf TopS.power > 180 Then
            STRout = STRout & "Средняя мощность сигнала, возможно, это ослабленный сигнал земного происхождения. "
        Else
            STRout = STRout & "Мощность сигнала невысока, возможно, это сигнал, пришедший из космоса. "
        End If
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    End If
    
    IndexLen = State.GetLastRecNum(3)  'Работа с индексом гауссиан
    Place = -1
    For i = 1 To IndexLen
        bResult = State.ReadIndex(1, i)
        'TO DO
        'Сделать здесь проверку Win/Lin...
        If TopG.ID = WinID Then
            Place = i       'Запомнить место
            i = IndexLen    'Завершить цикл
        End If
    Next i
    
    If Place = -1 Then      'Запись не найдена!
        STRout = "Данные о лучших гауссианах (рабочий блок " & Str(WinID) & ") не обнаружены. "
        STRout = STRout & "Возможно, Вы еще не запускали обработчик этой информации. "
        STRout = STRout & "Выберите пункт меню Статистика/Лучшие пики или Статистика/Лучшие гауссианы. "
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    Else
        STRout = "Лучшая гауссиана данного рабочего блока находится на " & Place & " месте среди" & Str(i - 1) & " обработанных за все время. " & vbCrLf & vbCrLf
        'Анализ интегрального показателя
        If TopG.average > 3 Then
            STRout = STRout & "Сигнал имеет форму, предельно близкую к идеальной! "
            STRout = STRout & "Этот блок попадет в число лучших на домашней странице проекта. "
        ElseIf TopG.average > 0.7 Then
            STRout = STRout & "Сигнал имеет форму, сильно приближающуюся к идеальной."
        ElseIf TopG.average > 0.3 Then
            STRout = STRout & "Сигнал имеет форму, достаточно близкую к идеальной."
        ElseIf TopG.average > 0.2 Then
            STRout = STRout & "Сигнал имеет форму, немного похожую на идеальную."
        ElseIf TopG.average > 0.18 Then
            STRout = STRout & "Сигнал имеет форму, плохо соответствующую идеальной."
        ElseIf TopG.average = 0 Then
            STRout = STRout & "Сильная зашумленность сигнала. Поиск гауссиан не производится."
        Else
            STRout = STRout & "Сигнал имеет форму, даже приблизительно не соответствующую идеальной."
        End If
        STRout = STRout & " "
        Text1.text = Text1.text & STRout
        'Анализ мощности сигнала
        If TopG.power > 2 Then
            STRout = "Очень высокая мощность сигнала в верхней точке гауссианы свидетельствует о том, "
            STRout = STRout & "что его источник находится очень близко от Земли и "
            STRout = STRout & "сигнал не успевает ослабнуть до момента фиксации его радиотелескопом в Аресибо."
        ElseIf TopG.power > 1.4 Then
            STRout = "Мощность сигнала в верхней точке гауссианы высокая, "
            STRout = STRout & "источник сигнала либо находится на орбите Земли (искусственное происхождение, наиболее вероятная версия), "
            STRout = STRout & "либо (маловероятно) излучается одной из мощных радиогалактик или пульсаров."
        ElseIf TopG.power > 0.75 Then
            STRout = "Мощность сигнала в верхней точке гауссианы средняя, "
            STRout = STRout & "с равной долей вероятности можно предположить как ослабленное излучение земного происхождения, "
            STRout = STRout & "так и дошедшее до Земли излучение одной из мощных радиогалактик или пульсаров."
        ElseIf TopG.power = 0 Then
            STRout = "Мощность гауссианы определить невозможно."
        Else
            STRout = "Низкая мощность в верхней точке гауссианы, "
            STRout = STRout & "вероятнее всего, это сигнал, пришедший из дальнего космоса (радиогалактики, пульсары), либо "
            STRout = STRout & "либо сильно ослабленное излучение земного происхождения."
        End If
        STRout = STRout & " "
        Text1.text = Text1.text & STRout
        'Статистический анализ
        Summa = 0
        NumValid = 0
        For i = 1 To IndexLen                   'Отсечка нулевых значений
            bResult = State.ReadIndex(1, i)
            'TO DO
            'Сделать здесь проверку Win/Lin...
            If Not (TopG.average = 0) Then
                Summa = Summa + TopG.average    'Суммируем все ненулевые величины
                NumValid = NumValid + 1
            End If
        Next i
        bResult = State.ReadIndex(1, Place)
        If Not (NumValid = 0) Then
            Summa = Summa / NumValid
        End If
        If Not (TopG.average = 0) Then
            If 1.5 * Summa < TopG.average Then
                STRout = "Интегральный показатель текущего блока значительно выше среднестатистического."
            ElseIf Summa < TopG.average Then
                STRout = "Интегральный показатель текущего блока выше среднестатистического."
            ElseIf Summa > TopG.average Then
                If Summa > 1.5 * TopG.average Then
                    STRout = "Интегральный показатель текущего блока значительно ниже среднестатистического."
                Else    'Summa > TopG.average
                    STRout = "Интегральный показатель текущего блока ниже среднестатистического."
                End If
            End If
        End If
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    End If
End Sub
