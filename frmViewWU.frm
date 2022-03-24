VERSION 5.00
Begin VB.Form frmViewWU 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметры текущего блока"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14850
   HelpContextID   =   10000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Калибровка..."
      Height          =   345
      Left            =   9450
      TabIndex        =   56
      Top             =   7035
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1030
      Left            =   8010
      ScaleHeight     =   975
      ScaleWidth      =   6090
      TabIndex        =   54
      Top             =   3045
      Width           =   6150
   End
   Begin VB.Frame Frame5 
      Caption         =   "Триплеты"
      Height          =   2535
      Left            =   7560
      TabIndex        =   37
      Top             =   4200
      Width           =   7050
      Begin VB.Line Line28 
         X1              =   4725
         X2              =   4620
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line27 
         X1              =   4620
         X2              =   6195
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line26 
         X1              =   4515
         X2              =   4725
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line25 
         X1              =   4620
         X2              =   4620
         Y1              =   945
         Y2              =   1575
      End
      Begin VB.Line Line24 
         X1              =   4515
         X2              =   4515
         Y1              =   1680
         Y2              =   315
      End
      Begin VB.Line Line23 
         X1              =   6195
         X2              =   6195
         Y1              =   1680
         Y2              =   1575
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   7
         Left            =   5460
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   6
         Left            =   3885
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   7
         Left            =   4725
         Top             =   840
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   6
         Left            =   4725
         Top             =   210
         Width           =   2430
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   53
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label42 
         Caption         =   "Смещение (Гц/сек)"
         Height          =   225
         Left            =   105
         TabIndex        =   52
         Top             =   1260
         Width           =   2535
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   49
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   48
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   47
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label36 
         Caption         =   "Праметр формы сигнала (score)"
         Height          =   225
         Left            =   105
         TabIndex        =   46
         Top             =   945
         Width           =   2535
      End
      Begin VB.Label Label35 
         Caption         =   "Период (period), сек"
         Height          =   225
         Left            =   105
         TabIndex        =   45
         Top             =   630
         Width           =   2535
      End
      Begin VB.Label Label34 
         Caption         =   "Лучший триплет (triplet power)"
         Height          =   225
         Left            =   105
         TabIndex        =   44
         Top             =   315
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Импульсы"
      Height          =   2535
      Left            =   7560
      TabIndex        =   36
      Top             =   105
      Width           =   7050
      Begin VB.Line Line22 
         X1              =   6090
         X2              =   4620
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line21 
         X1              =   4620
         X2              =   4620
         Y1              =   945
         Y2              =   1575
      End
      Begin VB.Line Line20 
         X1              =   4515
         X2              =   4515
         Y1              =   1680
         Y2              =   315
      End
      Begin VB.Line Line19 
         X1              =   6090
         X2              =   6090
         Y1              =   1680
         Y2              =   1575
      End
      Begin VB.Line Line18 
         X1              =   4725
         X2              =   4515
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line17 
         X1              =   4725
         X2              =   4620
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   5
         Left            =   5460
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   4
         Left            =   3885
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   5
         Left            =   4725
         Top             =   840
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   4
         Left            =   4725
         Top             =   210
         Width           =   2430
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   51
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label40 
         Caption         =   "Смещение (Гц/сек)"
         Height          =   225
         Left            =   105
         TabIndex        =   50
         Top             =   1260
         Width           =   2745
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   43
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   42
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   41
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label30 
         Caption         =   "Праметр формы сигнала (score)"
         Height          =   225
         Left            =   105
         TabIndex        =   40
         Top             =   945
         Width           =   2745
      End
      Begin VB.Label Label29 
         Caption         =   "Период (period), сек"
         Height          =   225
         Left            =   105
         TabIndex        =   39
         Top             =   630
         Width           =   2745
      End
      Begin VB.Label Label28 
         Caption         =   "Лучший импульс (pulse power)"
         Height          =   225
         Left            =   105
         TabIndex        =   38
         Top             =   315
         Width           =   2745
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7455
      Top             =   2940
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Краткий отчет"
      Height          =   345
      Left            =   11235
      TabIndex        =   35
      Top             =   7035
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Обновить"
      Height          =   345
      Left            =   7770
      TabIndex        =   27
      Top             =   7035
      Width           =   1467
   End
   Begin VB.Frame Frame3 
      Caption         =   "Пики и гауссианы"
      Height          =   3375
      Left            =   210
      TabIndex        =   6
      Top             =   3990
      Width           =   7050
      Begin VB.Line Line16 
         X1              =   840
         X2              =   840
         Y1              =   1680
         Y2              =   1890
      End
      Begin VB.Line Line15 
         X1              =   4305
         X2              =   840
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line14 
         X1              =   4305
         X2              =   4305
         Y1              =   420
         Y2              =   1680
      End
      Begin VB.Line Line13 
         X1              =   4725
         X2              =   4305
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line12 
         X1              =   3465
         X2              =   3465
         Y1              =   1785
         Y2              =   1890
      End
      Begin VB.Line Line11 
         X1              =   4410
         X2              =   3465
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line10 
         X1              =   4410
         X2              =   4410
         Y1              =   735
         Y2              =   1785
      End
      Begin VB.Line Line9 
         X1              =   4725
         X2              =   4410
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line8 
         X1              =   4830
         X2              =   4830
         Y1              =   1785
         Y2              =   1890
      End
      Begin VB.Line Line7 
         X1              =   4515
         X2              =   4830
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line6 
         X1              =   4515
         X2              =   4515
         Y1              =   1050
         Y2              =   1785
      End
      Begin VB.Line Line5 
         X1              =   4725
         X2              =   4515
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line4 
         X1              =   4620
         X2              =   4620
         Y1              =   1365
         Y2              =   1680
      End
      Begin VB.Line Line3 
         X1              =   4725
         X2              =   4620
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line2 
         X1              =   6405
         X2              =   4620
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         X1              =   6405
         X2              =   6405
         Y1              =   1890
         Y2              =   1680
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Label26"
         Height          =   225
         Left            =   4935
         TabIndex        =   33
         Top             =   3045
         Width           =   1590
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Label25"
         Height          =   225
         Left            =   5040
         TabIndex        =   32
         Top             =   2730
         Width           =   1485
      End
      Begin VB.Label Label24 
         Caption         =   "Смещение для лучшей гауссианы (Гц/сек)"
         Height          =   225
         Left            =   420
         TabIndex        =   31
         Top             =   3045
         Width           =   3375
      End
      Begin VB.Label Label23 
         Caption         =   "Смещение для лучшего пикового сигнала (Гц/сек)"
         Height          =   225
         Left            =   420
         TabIndex        =   30
         Top             =   2730
         Width           =   3900
      End
      Begin VB.Label Label4 
         Caption         =   "Интегральный показатель"
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   1365
         Width           =   2850
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   13
         Top             =   1365
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   3
         Left            =   4725
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   3
         Left            =   5670
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   2
         Left            =   4200
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   1
         Left            =   2730
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   645
         Index           =   0
         Left            =   210
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   2
         Left            =   4725
         Top             =   945
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   1
         Left            =   4725
         Top             =   630
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   225
         Index           =   0
         Left            =   4725
         Top             =   315
         Width           =   2430
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   12
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   11
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3255
         TabIndex        =   10
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Параметр формы сигнала (gaussian fit)"
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   1050
         Width           =   2955
      End
      Begin VB.Label Label2 
         Caption         =   "Лучшая гауссиана (gaussian power)"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   735
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Лучший пиковый сигнал (spike power)"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   2955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Общая информация о блоке"
      Height          =   2745
      Left            =   210
      TabIndex        =   5
      Top             =   1050
      Width           =   7050
      Begin VB.Label Label27 
         Height          =   225
         Left            =   3360
         TabIndex        =   34
         Top             =   945
         Width           =   3060
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PROCESSING"
         Height          =   435
         Left            =   3360
         TabIndex        =   29
         Top             =   2310
         Width           =   3480
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000FF00&
         Height          =   330
         Left            =   3360
         TabIndex        =   28
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Arecibo Radio Observatory"
         Height          =   225
         Left            =   3360
         TabIndex        =   26
         Top             =   1890
         Width           =   3480
      End
      Begin VB.Label Label21 
         Caption         =   "Источник"
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   1890
         Width           =   2220
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3360
         TabIndex        =   24
         Top             =   2205
         Width           =   3480
      End
      Begin VB.Label Label17 
         Caption         =   "Прогресс обработки"
         Height          =   225
         Left            =   210
         TabIndex        =   23
         Top             =   2265
         Width           =   2220
      End
      Begin VB.Label Label16 
         Caption         =   "1.234567890"
         Height          =   225
         Left            =   3360
         TabIndex        =   22
         Top             =   1575
         Width           =   3060
      End
      Begin VB.Label Label15 
         Caption         =   "Частота, Гц"
         Height          =   225
         Left            =   210
         TabIndex        =   21
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label Label14 
         Caption         =   "23 may 1999"
         Height          =   225
         Left            =   3360
         TabIndex        =   20
         Top             =   1260
         Width           =   3165
      End
      Begin VB.Label Label13 
         Caption         =   "Дата записи"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Label12 
         Caption         =   "XXX RA XXX DEC"
         Height          =   225
         Left            =   3360
         TabIndex        =   18
         Top             =   630
         Width           =   2850
      End
      Begin VB.Label Label11 
         Caption         =   "Координаты"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label10 
         Caption         =   "52"
         Height          =   225
         Left            =   3360
         TabIndex        =   16
         Top             =   315
         Width           =   435
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Индекс SETImap"
         Height          =   225
         Left            =   1680
         TabIndex        =   15
         Top             =   315
         Width           =   1590
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Default         =   -1  'True
      Height          =   345
      Left            =   13020
      TabIndex        =   4
      Top             =   7035
      Width           =   1467
   End
   Begin VB.Frame Frame1 
      Caption         =   "Выбор блока"
      Height          =   750
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   7050
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3570
         TabIndex        =   3
         Text            =   "Linux в C:\setilin"
         Top             =   315
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Другой"
         Height          =   330
         Index           =   1
         Left            =   2625
         TabIndex        =   2
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Клиент для Windows"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Label Label44 
      Caption         =   "На графике импульсы показаны желтым цветом, триплеты - красным"
      Height          =   225
      Left            =   8295
      TabIndex        =   55
      Top             =   2835
      Width           =   5580
   End
End
Attribute VB_Name = "frmViewWU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit        ' 0   1   2   3   4  стерерть этот комментарий потом
Dim Indy(7, 4) As Long 'Min Max End Cur Direction position for indicators

Private Sub Command1_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Command2_Click()
    Call DisplayInfo
End Sub

Private Sub Command3_Click()
Dim hFile As Long

    hFile = FreeFile
    If UseDefaultRF = 1 Then
        'Используем файл по-умолчанию
        Open (App.path & BackSlash & ReportFile) For Output As hFile
        Result = MsgBox("Краткий отчет записан в файл по-умолчанию " + ReportFile + vbCrLf + "расположенный в папке программы SETImap." + vbCrLf + "Изменить имя файла можно в главном меню" + vbCrLf + "(пункт Вид / Настройки...)", vbInformation + vbOKOnly, "Краткий отчет")
    Else
        'Используем настройки
        Open ReportFileReg For Output As hFile
        Result = MsgBox("Краткий отчет записан в файл" + vbCrLf + ReportFileReg + vbCrLf + "Изменить имя файла можно в главном меню" + vbCrLf + "(пункт Вид / Настройки...)", vbInformation + vbOKOnly, "Краткий отчет")
    End If
    If Option1(0).Value Then
        Write #hFile, "WINDOWS CLIENT ID", WinID
    Else
        Write #hFile, "LINUX CLIENT ID", LinID
    End If
    Write #hFile, "From"
    Print #hFile, Label12.Caption
    Print #hFile, Label27.Caption
    Write #hFile, "Date"
    Print #hFile, Label14.Caption
    Write #hFile, "Frequency"
    Print #hFile, Label16.Caption
    Write #hFile, "--------------"
    Write #hFile, "Spike power"
    Print #hFile, Label5.Caption
    Write #hFile, "at chirp rate"
    Print #hFile, Label25.Caption
    Write #hFile, "Gaussian power"
    Print #hFile, Label6.Caption
    Write #hFile, "Gaussian fit"
    Print #hFile, Label7.Caption
    Write #hFile, "at chirp rate"
    Print #hFile, Label26.Caption
    Write #hFile, "--------------"
    Write #hFile, "Pulse power"
    Print #hFile, Label31.Caption
    Write #hFile, "at chirp rate"
    Print #hFile, Label41.Caption
    Write #hFile, "Pulse score"
    Print #hFile, Label33.Caption
    Write #hFile, "Pulse period"
    Print #hFile, Label32.Caption
    Write #hFile, "--------------"
    Write #hFile, "Triple power"
    Print #hFile, Label37.Caption
    Write #hFile, "at chirp rate"
    Print #hFile, Label43.Caption
    Write #hFile, "Triple score"
    Print #hFile, Label39.Caption
    Write #hFile, "Triple period"
    Print #hFile, Label38.Caption
    Close (hFile)
End Sub

Private Sub Autorange()
Dim TMPvalue As Long
Dim Changes As Boolean  'Если были изменения, то, если разрешено в настройках,
                        'сохранить изменения в реестре.
Dim strErrText As String
    On Error GoTo AutorangeErr
    
    Changes = False
    'Spike -> power
strErrText = "Spike->power: autorange"
    TMPvalue = Trunc(CDbl(State.bs_power * 10000))
    If TMPvalue > MaxSpower Then
        Result = MsgBox("Параметр Spike->power превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 10000) + vbCrLf + "Текущее граничное значение: " & Str(MaxSpower / 10000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxSpower = TMPvalue
            Changes = True
        End If
    End If
    'Gaussian -> power
strErrText = "Gaussian->power: autorange"
    TMPvalue = Trunc(CDbl(State.bg_power * 100000))
    If TMPvalue > MaxGpower Then
        Result = MsgBox("Параметр Gaussian->power превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxGpower / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxGpower = TMPvalue
            Changes = True
        End If
    End If
    'Gaussian -> fit
strErrText = "Gaussian->fit: autorange"
    TMPvalue = Trunc(CDbl(State.bg_chisq * 100000))
    If Not (TMPvalue = 0) Then
        'Если =0, то просто значение отсутствует! Не сравнивать дальше!
        If TMPvalue < MaxGfit Then
            Result = MsgBox("Параметр Gaussian->fit лучше предельного предполагаемого значения." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее  граничное значение: " & Str(MaxGfit / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
            If Result = vbYes Then
                MaxGfit = TMPvalue
                Changes = True
            End If
        End If
    End If
    'Gaussian -> integral
strErrText = "Gaussian->integral: autorange"
    If Not (State.bg_chisq = 0) Then
        'Исключить деление на ноль!!!
        TMPvalue = Trunc(CDbl((State.bg_power * 100000) / State.bg_chisq))
        If Not (TMPvalue = 0) Then
        'Если =0, то просто значение отсутствует! Не сравнивать дальше!
            If TMPvalue > MaxGintegr Then
                Result = MsgBox("Параметр Gaussian->intergal parameter превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxGintegr / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
                If Result = vbYes Then
                    MaxGintegr = TMPvalue
                    Changes = True
                End If
            End If
        End If
    End If
    'Pulse -> power
strErrText = "Pulse->power: autorange"
    TMPvalue = Trunc(CDbl(State.bp_power * 100000))
    If TMPvalue > MaxPpower Then
        Result = MsgBox("Параметр Pulse->power превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxPpower / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxPpower = TMPvalue
            Changes = True
        End If
    End If
    'Pulse -> score
strErrText = "Pulse->Score: autorange"
    TMPvalue = Trunc(CDbl(State.bp_score * 100000))
    If TMPvalue > MaxPscore Then
        Result = MsgBox("Параметр Pulse->score превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxPscore / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxPscore = TMPvalue
            Changes = True
        End If
    End If
    'Triplet -> power
strErrText = "Triplet->power: autorange"
    TMPvalue = Trunc(CDbl(State.bt_power * 100000))
    If TMPvalue > MaxTpower Then
        Result = MsgBox("Параметр Triplet->power превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxTpower / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxTpower = TMPvalue
            Changes = True
        End If
    End If
    'Triplet -> score
strErrText = "Triplet->score: autorange"
    TMPvalue = Trunc(CDbl(State.bt_score * 100000))
    If TMPvalue > MaxTscore Then
        Result = MsgBox("Параметр Triplet->score превысил предполагаемое максимальное значение." + vbCrLf + "Новое расчетное значение: " & Str(TMPvalue / 100000) + vbCrLf + "Текущее граничное значение: " & Str(MaxTscore / 100000) + vbCrLf + "Провести корректировку?", vbExclamation + vbYesNo, "Автокалибровка")
        If Result = vbYes Then
            MaxTscore = TMPvalue
            Changes = True
        End If
    End If
    If Changes Then
    'Изменения произошли!
strErrText = "Was changes: autorange"
        If EnableRegSave Then
        'Настройки разрешают запись в реестр
strErrText = "Saving autorange parameters: autorange"
            SaveSetting App.Title, "AutoRange", "MaxPscore", MaxPscore
            SaveSetting App.Title, "AutoRange", "MaxPpower", MaxPpower
            SaveSetting App.Title, "AutoRange", "MaxTscore", MaxTscore
            SaveSetting App.Title, "AutoRange", "MaxTpower", MaxTpower
            SaveSetting App.Title, "AutoRange", "MaxGpower", MaxGpower
            SaveSetting App.Title, "AutoRange", "MaxGfit", MaxGfit
            SaveSetting App.Title, "AutoRange", "MaxGintegr", MaxGintegr
            SaveSetting App.Title, "AutoRange", "MaxSpower", MaxSpower
        End If
    End If
    Exit Sub
AutorangeErr:
    Err.Raise vbObjectError, "Autorange", strErrText
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    If DoLinux = 0 Then
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Combo1.Enabled = False
        Frame1.Enabled = False
    End If
    ''If DoLinux = 0 Then
    ''    Option1(1).Left = -20000
    ''    Combo1.Left = -20000
    ''Else
    ''    Option1(1).Left = 2625
    ''    Combo1.Left = 3570
    ''End If
    ClearGraph
    DisplayInfo
End Sub

'**********************************
'* Загружает рисунки по-умолчанию *
'**********************************
Private Sub ClearGraph()
Dim i As Long
    For i = 0 To 7
        Image1(i).Picture = LoadResPicture(110, vbResBitmap)
    Next i
    For i = 0 To 7
        Image2(i).Picture = LoadResPicture(130, vbResBitmap)
    Next i
    Picture1.Cls
End Sub

'***********************************************
'*  Подготавливает переменные для определения  *
'*  длины анимационной последовательности      *
'***********************************************
Private Sub InitAnimation()
Dim i As Long
    For i = 0 To 7
        Indy(i, 0) = Indy(i, 2) - 1
        If Indy(i, 0) < 0 Then
            Indy(i, 0) = 0
        End If
        Indy(i, 1) = Indy(i, 2) + 2
        If Indy(i, 1) > 16 Then
            Indy(i, 1) = 16
        End If
        Indy(i, 3) = 0  'исходное значение
        Indy(i, 4) = 1  'возрастает
    Next i
End Sub

Private Sub Animate()
    Timer1.Interval = AnimTick
    Timer1.Enabled = True
End Sub

Private Sub DisplayInfo()
Dim MathResult As Long
Dim i As Long
Dim Success As Boolean 'Признак успешного чтения данных
Dim FileMode As Long    'Необходимо для правильной работы cWU, хранит значение режима
                        '(Win/Lin) и передает его в cWU.GetFilePath...
Dim strErrText As String
    On Error GoTo DisplayInfoErr

    Success = False
    MathResult = 0
    If Option1(0).Value Then
    'Выбран клиент под Windows
        FileMode = 1
        If State.CheckFile(0) Then
        'Файл State.sah существует
            If State.DecodeState(State.ReadFile(0)) Then
            'Успешное декодирование - производим вывод на экран
                Success = True
                If Not (UpdateOnStartup = 1) Then
        'Сохранение текущих параметров. Можно включить, т.к. обычно
        'сохранение идет только при вызове этого окна. Если оно показывается
        'достаточно долгое время, то возможна потеря информации.
        'Недостаток: частая перезапись информации.
                    State.Interchange 0
                    If State.EncodeHistory Then
                        bResult = State.WriteHistory(1, WinID)
                    End If
                End If
            End If
        End If
        Label10.Caption = WinID
    Else
    'Выбран клиент под Linux или что похлеще...
    'TO DO
    'Здесь нужно будет сделать проверку: какой именно пункт из Combo1 выбран...
        FileMode = 2
        If State.CheckFile(1) Then
        'Файл State.sah существует
            If State.DecodeState(State.ReadFile(1)) Then
            'Успешное декодирование - производим вывод на экран
                Success = True
                If Not (UpdateOnStartup = 1) Then
        'Сохранение текущих параметров. Можно включить, т.к. обычно
        'сохранение идет только при вызове этого окна. Если оно показывается
        'достаточно долгое время, то возможна потеря информации.
        'Недостаток: частая перезапись информации.
                    State.Interchange 1
                    If State.EncodeHistory Then
                        bResult = State.WriteHistory(1, LinID)
                    End If
                End If
            End If
        End If
        Label10.Caption = LinID
    End If
'ПАРАМЕТРЫ ПРОЧИТАНЫ ИЗ ФАЙЛОВ - ПРОИЗВОДИМ ВЫВОД НА ЭКРАН
    If Success Then
        'Провести автокалибровку
        Call Autorange
        
        'Изменение индикатора прогресса: красный неподвижен, а зеленый увеличивается
strErrText = "Progress bar evaluating"
        Label18.Width = Trunc(Label20.Width * (CRtoPercent(State.cr) / 100))
        Label19.Width = Label20.Width - Label18.Width
        Label19.Left = Label18.Left + Label18.Width
        If CRtoPercent(State.cr) > 90 Then
            Label20.Caption = "APPROX 1 HOUR TO GO"
        ElseIf CRtoPercent(State.cr) < 10 Then
            Label20.Caption = "STARTING WITH WORKUNIT"
        Else
            Label20.Caption = "PROCESSING"
        End If
        
        'Аналоговые индикаторы и цифровые значения соответствующих параметров
        'SPIKE power
'''        MathResult = CDbl(State.bs_power) \ 20
'''        If MathResult > 15 Then
'''            Image2(0).Picture = LoadResPicture(146, vbResBitmap)
'''            MathResult = 15
'''        Else
'''            Image2(0).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        End If
'''        Image1(0).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
strErrText = "State->power bar evaluating"
        If State.bs_power < 250 Then
        'Экпотенциальная зависимость
            MathResult = Trunc(CDbl((State.bs_power * 12 * 10000) / 2300000))
        Else
            MathResult = Trunc(CDbl(12 + ((State.bs_power - 230) * 5 * 10000) / (MaxSpower - 230)))
        End If
        Label5.Caption = Str(State.bs_power)
'''        Image1(0).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
'''        Image2(0).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
        Indy(0, 2) = MathResult
        
        
        'GAUSSIAN power
'''        MathResult = Trunc(CDbl(State.bg_power / 0.2))
'''        If MathResult > 15 Then
'''            Image2(1).Picture = LoadResPicture(146, vbResBitmap)
'''            MathResult = 15
'''        Else
'''            Image2(1).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        End If
strErrText = "Gaussian->power bar evaluating"
        MathResult = Trunc(CDbl((State.bg_power * 16 * 100000) / MaxGpower))
        Label6.Caption = Str(State.bg_power)
'''        Image1(1).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
'''        Image2(1).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
        Indy(1, 2) = MathResult
        
        'GAUSSIAN fit
        'Выбираем картинку ВРУЧНУЮ!!!
        'Формула не очень-то получилась. Трудности связаны с неизвестным максимумом,
        'нелинейность параметра в (0;1) и (5;inf)
strErrText = "Gaussian->fit bar evaluating"
        MathResult = Trunc(CDbl(State.bg_chisq)) * 10
        If MathResult > 250 Then
            MathResult = 16
        ElseIf MathResult > 200 Then
            MathResult = 15
        ElseIf MathResult > 160 Then
            MathResult = 14
        ElseIf MathResult > 130 Then
            MathResult = 13
        ElseIf MathResult > 105 Then
            MathResult = 12
        ElseIf MathResult > 80 Then
            MathResult = 11
        ElseIf MathResult > 60 Then
            MathResult = 10
        ElseIf MathResult > 40 Then
            MathResult = 9
        ElseIf MathResult > 25 Then
            MathResult = 8
        ElseIf MathResult > 15 Then
            MathResult = 7
        ElseIf MathResult > 10 Then
            MathResult = 6
        ElseIf MathResult > 8 Then
            MathResult = 5
        ElseIf MathResult > 6 Then
            MathResult = 4
        ElseIf MathResult > 4 Then
            MathResult = 3
        ElseIf MathResult > 2 Then
            MathResult = 2
        ElseIf MathResult > 1 Then
            MathResult = 1
        ElseIf MathResult = 0 Then  'Это спецслучай, когда гауссианы отсутствуют.
            MathResult = 16
        Else
            MathResult = 0
        End If
        Label7.Caption = Str(State.bg_chisq)
'''        Image1(2).Picture = LoadResPicture(126 - MathResult, vbResBitmap)
'''        Image2(2).Picture = LoadResPicture(146 - MathResult, vbResBitmap)
        Indy(2, 2) = 16 - MathResult
        
        'GAUSSIAN integral
'''        If Not (State.bg_chisq = 0) Then
'''            MathResult = Trunc((State.bg_power / State.bg_chisq) * 100)
'''            Label8.Caption = Str(State.bg_power / State.bg_chisq)
'''        Else
'''        'Исключить деление на ноль!!!
'''            MathResult = 0
'''            Label8.Caption = 0
'''        End If
'''        If MathResult > 50 Then
'''            MathResult = 11 + Trunc(MathResult / 200)
'''        Else
'''            MathResult = Trunc(MathResult / 5)
'''        End If
'''        If MathResult > 15 Then
'''            Image2(3).Picture = LoadResPicture(146, vbResBitmap)
'''            MathResult = 15
'''        Else
'''            Image2(3).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        End If
'''        Image1(3).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
strErrText = "Gaussian->integral bar evaluating"
        If Not (State.bg_chisq = 0) Then
            If Trunc(CDbl(State.bg_power / State.bg_chisq)) < 1 Then
            'Экпотенциальная зависимость
                MathResult = Trunc(CDbl(((State.bg_power / State.bg_chisq) * 12 * 100000) / 100000))
            Else
                MathResult = Trunc(CDbl(12 + (((State.bg_power / State.bg_chisq) - 1) * 4 * 100000) / MaxGintegr - 100000))
            End If
            Label8.Caption = Str(State.bg_power / State.bg_chisq)
        Else
        'Исключить деление на ноль!!!
            MathResult = 0
            Label8.Caption = 0
        End If
'''        Image1(3).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
'''        Image2(3).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
        Indy(3, 2) = MathResult
        
        If Not (State.bg_chisq = 0) Then
            Label8.Caption = Str(State.bg_power / State.bg_chisq)
        Else
        'Исключить деление на ноль!!!
            Label8.Caption = 0
        End If
        
        'Импульс - мощность
'''        MathResult = Trunc(CDbl(State.bp_power / 0.2))
'''        If MathResult > 16 Then
'''            MathResult = 16
'''        End If
strErrText = "Pulse->power bar evaluating"
        MathResult = Trunc(CDbl((State.bp_power * 16 * 100000) / MaxPpower))
        Label31.Caption = Str(State.bp_power)
'''        Image2(4).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        Image1(4).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
        Indy(4, 2) = MathResult
        
        'Импульс - показатель
'''        MathResult = Trunc((CDbl(State.bp_score) * 16) / 1.2)
'''        If MathResult > 16 Then
'''            MathResult = 16
'''        End If
strErrText = "Pulse->score bar evaluating"
        MathResult = Trunc(CDbl((State.bp_score * 16 * 100000) / MaxPscore))
        Label33.Caption = Str(State.bp_score)
'''        Image2(5).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        Image1(5).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
        Indy(5, 2) = MathResult
        
        'Триплет - мощность
'''        MathResult = Trunc(CDbl(State.bt_power / 0.2))
'''        If MathResult > 16 Then
'''            MathResult = 16
'''        End If
strErrText = "Triplet->power bar evaluating"
        MathResult = Trunc(CDbl((State.bt_power * 16 * 100000) / MaxTpower))
        Label37.Caption = Str(State.bt_power)
'''        Image2(6).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        Image1(6).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
        Indy(6, 2) = MathResult
        
        'Триплет - показатель
'''        MathResult = Trunc((CDbl(State.bt_score) * 16) / 1.2)
'''        If MathResult > 16 Then
'''            MathResult = 16
'''        End If
strErrText = "Triplet->score bar evaluating"
        MathResult = Trunc(CDbl((State.bt_score * 16 * 100000) / MaxTscore))
        Label39.Caption = Str(State.bt_score)
'''        Image2(7).Picture = LoadResPicture(130 + MathResult, vbResBitmap)
'''        Image1(7).Picture = LoadResPicture(110 + MathResult, vbResBitmap)
        Indy(7, 2) = MathResult
        
strErrText = "Doppler shift bars evaluating"
        'Сдвиги Допплера для всех главных параметров
        Label25.Caption = Str(State.bs_rate)
        Label26.Caption = Str(State.bg_rate)
        Label41.Caption = Str(State.bp_chirp_rate)
        Label43.Caption = Str(State.bt_chirp_rate)
        
strErrText = "Others bar evaluating"
        'Другая информация о главных параметрах
        Label32.Caption = Str(State.bp_period)
        Label38.Caption = Str(State.bt_period)
        
        'Заполнение информационных полей - может занять некоторое время
        'Проверка блока
        If (Dir(WU.GetFilePath(FileMode), vbNormal) <> "") Then
            WU.DecodeWU (WU.ReadFile(FileMode))
            Label12.Caption = " " & DecodeRA(CDbl(Val(WU.StartRA))) & " RA"
            Label27.Caption = "+" & DecodeDEC(CDbl(Val(WU.StartDEC))) & " DEC"
            Label14.Caption = ExtractTime(WU.TimeOfRec)
            Label16.Caption = WU.SubbandBase
            If WU.Receiver = "ao1420" Then
                Label22.Caption = "Arecibo Radio Observatory"
            Else
                Label22.Caption = "Unknown source - " & WU.Receiver
            End If
        End If
        
strErrText = "Calling DisplayGraph"
        'Рисование импульсов / триплетов
        DisplayGraph
        
        If AllowAnim Then
strErrText = "Start animation sequence"
            'Вычисление фаз анимаций на основе полученных результатов
            Call InitAnimation
            'Пуск таймера
            Call Animate
        Else
strErrText = "No animation"
            'Произвести обычный вывод на экран
            For i = 0 To 7
                Image2(i).Picture = LoadResPicture(130 + Indy(i, 2), vbResBitmap)
                Image1(i).Picture = LoadResPicture(110 + Indy(i, 2), vbResBitmap)
            Next i
        End If
        
        StatusStr.Caption = "Просмотр сведений об обнаруженных в текущем блоке сигналах."
    Else
strErrText = "No STATE.SAH file founded"
        'Нет файла State.sah! Попробуем найти Result.sah
        ClearGraph  'Очистим графику - это ЧУЖИЕ значения
        If Option1(0).Value Then
            If OutResult.CheckFile(0, 1) Then
            'Выбран клиент под Windows
                Label18.Width = Label20.Width
                Label19.Width = 0
                Label19.Left = Label18.Left + Label18.Width
                Label20.Caption = "PROCESSING COMPLETE"
            Else
                Label18.Width = 0
                Label19.Width = Label20.Width
                Label19.Left = Label18.Left
                Label20.Caption = "NO DATA TO DISPLAY"
            End If
        Else
            'Выбран клиент под Linux или что похлеще...
            'TO DO
            'Здесь нужно будет сделать проверку: какой именно пункт из Combo1 выбран...
            'См. начало этой функции
            If OutResult.CheckFile(1, 1) Then
            'Выбран клиент под WIndows
                Label18.Width = Label20.Width
                Label19.Width = 0
                Label19.Left = Label18.Left + Label18.Width
                Label20.Caption = "PROCESSING COMPLETE"
            Else
                Label18.Width = 0
                Label19.Width = Label20.Width
                Label19.Left = Label18.Left
                Label20.Caption = "NO DATA TO DISPLAY"
            End If
        End If
    End If  'Вывод завершен
    Exit Sub
DisplayInfoErr:
    Err.Raise vbObjectError, "DisplayInfo", strErrText
End Sub

Private Sub DisplayGraph()
Dim DataStringP As String, DataStringT As String    'Сигналы (pulse/triplet)
Dim sValueP As String, sValueT As String    'Величина сигнала (pulse/triplet) (HEX)
Dim lValueP As Long, lValueT As Long        'Величина сигнала (pulse/triplet) (DEC)
Dim i As Long
Dim PulseY As Long, TripletY As Long
Dim X As Long
            
    DataStringP = GetTokenEx("bp_pot=", State.ReadFile(0), "bt_score=", False)
    DataStringT = GetTokenEx("bt_pot=", State.ReadFile(0), "*", False)
    Picture1.BackColor = vbBlack
    ''Picture1.ForeColor = vbYellow
    Picture1.CurrentX = 0
    Picture1.CurrentY = 512
    For i = 1 To 1023 Step 2
        sValueP = Mid(DataStringP, i, 2)
        sValueT = Mid(DataStringT, i, 2)
        lValueP = CLng(Val("&H" & sValueP))
        lValueT = CLng(Val("&H" & sValueT))
        X = (i - 1) * 6
        PulseY = 1024 - lValueP * 4
        TripletY = 1024 - lValueT * 4
        If PulseY < TripletY Then   'PulseY на графике ВЫШЕ...
            Picture1.ForeColor = vbYellow
            Picture1.Line (X, PulseY)-(X, 1024)
            Picture1.Line (X, PulseY)-(X + 6, PulseY)
            Picture1.Line (X + 6, PulseY)-(X + 6, 1024)
            Picture1.ForeColor = vbRed
            Picture1.Line (X, TripletY)-(X, 1024)
            Picture1.Line (X, TripletY)-(X + 6, TripletY)
            Picture1.Line (X + 6, TripletY)-(X + 6, 1024)
        Else                        'TripleY на графике ВЫШЕ...
            Picture1.ForeColor = vbRed
            Picture1.Line (X, TripletY)-(X, 1024)
            Picture1.Line (X, TripletY)-(X + 6, TripletY)
            Picture1.Line (X + 6, TripletY)-(X + 6, 1024)
            Picture1.ForeColor = vbYellow
            Picture1.Line (X, PulseY)-(X, 1024)
            Picture1.Line (X, PulseY)-(X + 6, PulseY)
            Picture1.Line (X + 6, PulseY)-(X + 6, 1024)
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Command2.Enabled = True
    frmMain.mnuWUInfo.Enabled = True
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Combo1.Enabled = False
        Case 1:
            Combo1.Enabled = True
    End Select
    DisplayInfo
End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim NoJob As Boolean
    NoJob = True
    For i = 0 To 7
'''        If Indy(i, 4) <> 3 Then
'''            NoJob = False
'''            If Indy(i, 3) < Indy(i, 2) Then
'''                Indy(i, 3) = Indy(i, 3) + 1
'''            Else
'''                Indy(i, 4) = 3
'''            End If
'''            Image1(i).Picture = LoadResPicture(110 + Indy(i, 3), vbResBitmap)
'''            Image2(i).Picture = LoadResPicture(130 + Indy(i, 3), vbResBitmap)
'''        End If
        
        Select Case Indy(i, 4)
            Case 0: 'Увеличение на единицу (финальное)
                NoJob = False
                If Indy(i, 3) < Indy(i, 2) Then
                    Indy(i, 3) = Indy(i, 3) + 1
                End If
                Indy(i, 4) = 3
            Case 1: 'Увеличение
                NoJob = False
                If Indy(i, 3) < Indy(i, 1) Then
                    Indy(i, 3) = Indy(i, 3) + 1
                Else
                    Indy(i, 4) = 2
                    'Меняем направление, этот кадр пропускаем (замираем на месте)
                End If
            Case 2: 'Уменьшение
                NoJob = False
                If Indy(i, 3) > Indy(i, 0) Then
                    Indy(i, 3) = Indy(i, 3) - 1
                Else
                    Indy(i, 4) = 0
                    'Меняем направление, этот кадр пропускаем (замираем на месте)
                End If
            Case 3: 'Стоп!
        End Select
        Image1(i).Picture = LoadResPicture(110 + Indy(i, 3), vbResBitmap)
        Image2(i).Picture = LoadResPicture(130 + Indy(i, 3), vbResBitmap)
    Next i
    If NoJob Then
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
End Sub
