VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1575
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Справка"
      Height          =   375
      Left            =   105
      TabIndex        =   18
      Top             =   4515
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4515
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4515
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "&Apply"
      Top             =   4515
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame2 
         Caption         =   "Файлы отчетов"
         Height          =   1380
         Left            =   210
         TabIndex        =   19
         Top             =   2205
         Width           =   5265
         Begin VB.CommandButton Command2 
            Height          =   350
            Left            =   4620
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   630
            Width           =   400
         End
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   210
            TabIndex        =   23
            Top             =   630
            Width           =   4320
         End
         Begin VB.Label Label5 
            Caption         =   "По-умолчанию используется файл sreport.txt в папке SETImap"
            Height          =   225
            Left            =   210
            TabIndex        =   25
            Top             =   1050
            Width           =   4845
         End
         Begin VB.Label Label4 
            Caption         =   "Файл краткого отчета о текущих результатах:"
            Height          =   225
            Left            =   210
            TabIndex        =   20
            Top             =   315
            Width           =   4215
         End
      End
      Begin VB.Frame fraSample4 
         Caption         =   "Файлы данных"
         Height          =   1815
         Left            =   210
         TabIndex        =   10
         Tag             =   "Sample 4"
         Top             =   180
         Width           =   5265
         Begin VB.CheckBox Check7 
            Caption         =   "Производить импорт журналов при отсутствии текущего рабочего блока"
            Height          =   435
            Left            =   210
            TabIndex        =   36
            Top             =   315
            Width           =   4845
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Установки обработки рабочих блоков"
         Height          =   3390
         Left            =   210
         TabIndex        =   29
         Tag             =   "Sample 3"
         Top             =   210
         Width           =   5190
         Begin VB.OptionButton Option1 
            Caption         =   "Только по команде пользователя"
            Height          =   330
            Index           =   0
            Left            =   525
            TabIndex        =   34
            Top             =   630
            Value           =   -1  'True
            Width           =   3060
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Автоматически при старте"
            Height          =   330
            Index           =   1
            Left            =   525
            TabIndex        =   33
            Top             =   945
            Width           =   3060
         End
         Begin VB.Frame Frame1 
            Caption         =   "Повторные результаты"
            Height          =   1065
            Left            =   315
            TabIndex        =   30
            Top             =   1470
            Width           =   4530
            Begin VB.OptionButton Option2 
               Caption         =   "Пропускать"
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   32
               Top             =   315
               Value           =   -1  'True
               Width           =   3480
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Перезаписывать поверх старых"
               Height          =   225
               Index           =   1
               Left            =   210
               TabIndex        =   31
               Top             =   630
               Width           =   3585
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Обработка результатов"
            Height          =   1065
            Left            =   315
            TabIndex        =   35
            Top             =   315
            Width           =   4530
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Настройки карты звездного неба"
         Height          =   3390
         Left            =   210
         TabIndex        =   8
         Tag             =   "Sample 2"
         Top             =   210
         Width           =   5295
         Begin MSComctlLib.Slider Slider2 
            Height          =   225
            Left            =   210
            TabIndex        =   16
            Top             =   2310
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   397
            _Version        =   393216
            LargeChange     =   1
            Max             =   2
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   225
            Left            =   210
            TabIndex        =   14
            Top             =   1680
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   397
            _Version        =   393216
            LargeChange     =   1
            Max             =   2
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Выделять последний блок цветом"
            Height          =   330
            Left            =   210
            TabIndex        =   13
            Top             =   840
            Width           =   4845
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Автоматически перерисовывать (падение скорости)"
            Height          =   330
            Left            =   210
            TabIndex        =   12
            Top             =   420
            Width           =   4845
         End
         Begin VB.Label Label2 
            Caption         =   "Размер маркера"
            Height          =   225
            Left            =   315
            TabIndex        =   17
            Top             =   2100
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Вид маркера"
            Height          =   225
            Left            =   315
            TabIndex        =   15
            Top             =   1470
            Width           =   3060
         End
         Begin VB.Image Image1 
            Height          =   750
            Left            =   3885
            Top             =   1680
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Настройки SETImap"
         Height          =   3390
         Left            =   208
         TabIndex        =   4
         Tag             =   "Sample 1"
         Top             =   207
         Width           =   5295
         Begin VB.CheckBox Check8 
            Caption         =   "Отключить проверку Linux-клиента при старте SETImap"
            Height          =   330
            Left            =   315
            TabIndex        =   37
            Top             =   2100
            Width           =   4740
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Анимация"
            Height          =   330
            Left            =   315
            TabIndex        =   27
            Top             =   1365
            Width           =   1065
         End
         Begin MSComctlLib.Slider Slider3 
            Height          =   225
            Left            =   1785
            TabIndex        =   26
            Top             =   1680
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   397
            _Version        =   393216
            LargeChange     =   1
            Max             =   4
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Обновлять текущие результаты только при старте"
            Height          =   330
            Left            =   315
            TabIndex        =   22
            Top             =   1050
            Width           =   4635
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Сохранять результаты автокалибровки (рекомендуется)"
            Height          =   330
            Left            =   315
            TabIndex        =   21
            Top             =   735
            Width           =   4635
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Показывать информационные сообщения"
            Enabled         =   0   'False
            Height          =   330
            Left            =   315
            TabIndex        =   11
            Top             =   420
            Width           =   4635
         End
         Begin VB.Label Label6 
            Enabled         =   0   'False
            Height          =   225
            Left            =   1785
            TabIndex        =   28
            Top             =   1470
            Width           =   2850
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Общие"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Карта"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Результаты"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Файлы"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedStrip As Long

Private Sub Check2_Click()
    RedrawOnStartup = Check2.Value
End Sub

Private Sub Check3_Click()
    LastInColor = Check3.Value
End Sub

Private Sub Check4_Click()
    EnableRegSave = Check4.Value
End Sub

Private Sub Check5_Click()
    UpdateOnStartup = Check5.Value
End Sub

Private Sub Check6_Click()
    AllowAnim = Check6.Value
    If AllowAnim = 0 Then
        Slider3.Enabled = False
        Label6.Enabled = False
    Else
        Slider3.Enabled = True
        Label6.Enabled = True
    End If
End Sub

Private Sub Check7_Click()
    DoImport = Check7.Value
End Sub

Private Sub Check8_Click()
    DoLinux = Check8.Value
End Sub

Private Sub cmdApply_Click()
    SaveSetting App.Title, "Starmap", "MarkerType", MarkerType
    SaveSetting App.Title, "Starmap", "MarkerSize", MarkerSize
    SaveSetting App.Title, "Starmap", "RedrawOnStartup", RedrawOnStartup
    SaveSetting App.Title, "Starmap", "LastInColor", LastInColor
    SaveSetting App.Title, "Settings", "AutoShowWU", AutoShowWU
    SaveSetting App.Title, "Settings", "EnableRegSave", EnableRegSave
    SaveSetting App.Title, "Settings", "UpdateOnStartup", UpdateOnStartup
    SaveSetting App.Title, "Settings", "SplitterOverwrite", SplitterOverwr
    SaveSetting App.Title, "Settings", "AllowAnim", AllowAnim
    SaveSetting App.Title, "Settings", "DoLinux", DoLinux
    SaveSetting App.Title, "Settings", "DoImport", DoImport
    If AllowAnim = 1 Then
        'Время показа кадра сохраняем только если анимация разрешена
        SaveSetting App.Title, "Settings", "AnimationTick", AnimTick
    End If
    If Text1.text = "" Then
        'Файл по-умолчанию
        UseDefaultRF = 1    'Да, использовать
        SaveSetting App.Title, "Settings", "UseDefaultReportFile", UseDefaultRF
        DeleteSetting App.Title, "Settings", "ReportFile"
    Else
        'Файл указан
        'Заодно заполним переменную на тот случай, если пользователь сразу
        'захочет создать отчет без предварительного рестарта программы
        ReportFileReg = Text1.text
        UseDefaultRF = 0
        SaveSetting App.Title, "Settings", "UseDefaultReportFile", UseDefaultRF
        SaveSetting App.Title, "Settings", "ReportFile", ReportFileReg
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call cmdApply_Click
    Unload Me
End Sub

Private Sub Command1_Click()
    'ToDo: Add 'Command1_Click' code.
    MsgBox "Help Code goes here!"
End Sub

'************************************************
'*  Показать окно Open для выбора файла отчета  *
'************************************************
Private Sub Command2_Click()
    On Error GoTo ErrHandler
    'CancelError is True. Нажатие Cancel в диалоговом окне вызовет ошибку,
    'что позволит отследить это действие пользователя (отмену открытия файла).
    With CommonDialog1
        'Остлеживание нажатия Cancel
        .CancelError = True
        'Заголовок окна
        .DialogTitle = "Файл краткого отчета..."
        'Установка фильтров
        .Filter = "Все файлы (*.*)|*.*|Текстовые файлы (*.txt)|*.txt|Файлы документов (*.doc)|*.doc"
        'Выбрать по-умолчанию тип файлов TXT
        .FilterIndex = 1
        'Показать окно
        .ShowOpen
    End With
''''    CommonDialog1.ShowOpen
''''    'Заполнить поле именем файла
    Text1.text = CommonDialog1.FileName
    'TO DO Запомнить это имя!
    Exit Sub

ErrHandler:
    'Была нажата кнопка Cancel.
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    Slider1.Value = MarkerType
    Slider2.Value = MarkerSize
    Image1.Picture = LoadResPicture(150 + Slider1.Value * 10 + Slider2.Value, vbResBitmap)
    Check2.Value = RedrawOnStartup
    Check3.Value = LastInColor
    Check4.Value = EnableRegSave
    Check5.Value = UpdateOnStartup
    Check6.Value = AllowAnim
    Check7.Value = DoImport
    Check8.Value = DoLinux
    Select Case AnimTick
        Case 10:
            Slider3.Value = 0
            'Slider по-умолчанию тут стоит, поэтому при загрузке и таких настройках
            'текст не выводился. Теперь все должно быть ОК.
            Label6.Caption = "Скорость анимации: очень быстро"
        Case 25:    Slider3.Value = 1
        Case 50:    Slider3.Value = 2
        Case 100:   Slider3.Value = 3
        Case 250:   Slider3.Value = 4
    End Select
    If AllowAnim = 1 Then
        Slider3.Enabled = True
        Label6.Enabled = True
    Else
        Slider3.Enabled = False
        Label6.Enabled = False
    End If
    Option1(AutoShowWU).Value = True
    Option2(SplitterOverwr).Value = True 'При SplittterOverwr=0 (нет) Opt2(0)=True ("Пропускать")
    Command2.Picture = LoadResPicture(180, vbResBitmap)
    If Not (UseDefaultRF) Then
        Text1.text = ReportFileReg
    End If
    
End Sub

Private Sub Option1_Click(Index As Integer)
    AutoShowWU = CLng(Index)
End Sub

Private Sub Option2_Click(Index As Integer)
    SplitterOverwr = CLng(Index)
End Sub

Private Sub Slider1_Change()
    Image1.Picture = LoadResPicture(150 + Slider1.Value * 10 + Slider2.Value, vbResBitmap)
    MarkerType = Slider1.Value
End Sub

Private Sub Slider2_Change()
    Image1.Picture = LoadResPicture(150 + Slider1.Value * 10 + Slider2.Value, vbResBitmap)
    MarkerSize = Slider2.Value
End Sub

Private Sub Slider3_Change()
    Label6.Caption = "Скорость анимации: "
    Select Case Slider3.Value
        Case 0:
            AnimTick = 10
            Label6.Caption = Label6.Caption & "очень быстро"
        Case 1:
            AnimTick = 25
            Label6.Caption = Label6.Caption & "быстро"
        Case 2:
            AnimTick = 50
            Label6.Caption = Label6.Caption & "средне"
        Case 3:
            AnimTick = 100
            Label6.Caption = Label6.Caption & "медленно"
        Case 4:
            AnimTick = 250
            Label6.Caption = Label6.Caption & "очень медленно"
    End Select
End Sub

Private Sub tbsOptions_Click()
Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
End Sub
