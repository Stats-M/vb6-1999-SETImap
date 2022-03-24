VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETImap"
   ClientHeight    =   8820
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   14685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   14685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "INFO"
      Height          =   345
      Left            =   13020
      TabIndex        =   5
      Top             =   7980
      Width           =   1467
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   7230
      Left            =   105
      ScaleHeight     =   7170
      ScaleWidth      =   14415
      TabIndex        =   4
      Top             =   525
      Width           =   14475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Показать!"
      Height          =   345
      Left            =   210
      TabIndex        =   2
      Top             =   7980
      Width           =   1467
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Macro"
            Object.ToolTipText     =   "Macro"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help What's This"
            Object.ToolTipText     =   "Help What's This"
            ImageKey        =   "Help What's This"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3570
      Top             =   7875
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4935
      Top             =   7875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0864
            Key             =   "Help What's This"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   9300
      X2              =   9300
      Y1              =   8460
      Y2              =   8735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000010&
      X1              =   45
      X2              =   45
      Y1              =   8460
      Y2              =   8735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   45
      X2              =   9295
      Y1              =   8460
      Y2              =   8460
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   45
      X2              =   9295
      Y1              =   8730
      Y2              =   8730
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   8505
      Width           =   8940
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9975
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9975
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   435
      Left            =   1890
      TabIndex        =   1
      Top             =   7980
      Width           =   10935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuSave 
         Caption         =   "Сохранить картинку..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "Вид"
      Begin VB.Menu mnuCurrentWU 
         Caption         =   "Показывать текущий блок"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPrevWU 
         Caption         =   "Показывать предыдущие блоки"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuWUNumber 
         Caption         =   "Показывать номер блока"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuWUComment 
         Caption         =   "Показывать комментарии"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuHyp11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAllWU 
         Caption         =   "Показать координаты блоков"
      End
      Begin VB.Menu mnuShowBorders 
         Caption         =   "Показать границы"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuHyp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Увеличить"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Очистить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuHyp10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWhereGaussians 
         Caption         =   "Источники лучших гауссиан"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuWhereSpikes 
         Caption         =   "Источники лучших пиков"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuHyp9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRange 
         Caption         =   "Калибровка..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVisuals 
         Caption         =   "Настройки..."
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Информация"
      Begin VB.Menu mnuHistory 
         Caption         =   "Журнал блоков"
      End
      Begin VB.Menu mnuHistoryState 
         Caption         =   "Журнал результатов"
      End
      Begin VB.Menu mnuHyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserInfo 
         Caption         =   "О пользователе"
      End
      Begin VB.Menu mnuWUInfo 
         Caption         =   "О текущем блоке"
      End
   End
   Begin VB.Menu mnuStats 
      Caption         =   "Статистика"
      Begin VB.Menu mnuTopSpikes 
         Caption         =   "Лучшие пики..."
      End
      Begin VB.Menu mnuTopGauss 
         Caption         =   "Лучшие гауссианы..."
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraphs 
         Caption         =   "Графики..."
      End
      Begin VB.Menu mnuTheory 
         Caption         =   "Теория..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Справка"
      Begin VB.Menu mnuContents 
         Caption         =   "Вызов справки"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpOnHelp 
         Caption         =   "Как пользоваться справкой"
      End
      Begin VB.Menu mnuHyp8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
   End
   Begin VB.Menu mnuHideOnMap 
      Caption         =   "НаКарте"
      Visible         =   0   'False
      Begin VB.Menu mnuHTopSpikes 
         Caption         =   "Лучшие пики..."
      End
      Begin VB.Menu mnuHTopGauss 
         Caption         =   "Лучшие гауссианы..."
      End
      Begin VB.Menu mnuHyp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHGraphs 
         Caption         =   "Графики..."
      End
      Begin VB.Menu mnuHTheory 
         Caption         =   "Теория..."
      End
      Begin VB.Menu mnuHyp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHHistory 
         Caption         =   "Журнал блоков"
      End
      Begin VB.Menu mnuHHistoryState 
         Caption         =   "Журнал результатов"
      End
      Begin VB.Menu mnuHyp6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHUserInfo 
         Caption         =   "О пользователе"
      End
      Begin VB.Menu mnuHWUInfo 
         Caption         =   "О текущем блоке"
      End
      Begin VB.Menu mnuHyp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHZoom 
         Caption         =   "Увеличить"
      End
      Begin VB.Menu mnuHClear 
         Caption         =   "Очистить"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mWU As cWU
Attribute mWU.VB_VarHelpID = -1

Private Sub Command1_Click()
''''''''''''''Dim NextLine As String
''''''''''''''Dim sfile As String
Dim i As Long
''Dim A As Long

    For i = 1 To RegRecords - 1
        WU.ClearAll (0)
        WU.DecodeHistory (WU.ReadHistory(i, 0))
        ShowPosition WU.StartRA, WU.StartDEC, "", "", False, 1
        StatusStr.Caption = "Выполнено (%) " + Str((i * 100) \ RegRecords) + " (" + Str(i) + " блоков)"
        StatusStr.Refresh
    Next i
'Последний блок - отдельно, чтобы не проводить постоянно проверки в цикле
    WU.ClearAll (0)
    WU.DecodeHistory (WU.ReadHistory(RegRecords, 0))
    If LastInColor = 1 Then
        ShowPosition WU.StartRA, WU.StartDEC, "", "", True, 1
    Else
        ShowPosition WU.StartRA, WU.StartDEC, "", "", False, 1
    End If
    StatusStr.Caption = "Выполнено 100%"
    StatusStr.Refresh
    
    frmMain.Refresh
    
    'Open SETIpath & "\" & FileUser For Input As #1
''''''''''''''    Open App.path & "\OldBlocks\Result25.txt" For Input As #1
    'Open App.Path & "\" & FileState For Input As #1
    'Open App.Path & "\" & FileUser For Input As #1
''    A = 0
''''''''''''''    NextLine = ""
''''''''''''''    sfile = ""
    ''For i = 1 To 350
        ''Line Input #1, NextLine
        ''NextLine = NextLine & Input(1, #1)
        ''If NextLine = Chr(10) Or NextLine = Chr(13) Then
            ''NextLine = "*"
            ''a = a + 1
            ''If a = 25 Then
                ''i = 255
            ''End If
        ''End If
'        If Not (i < 800) Then
            ''Label1.Caption = Label1.Caption + NextLine
'        End If
    ''Next i
    ''Label1.Caption = Label1.Caption + NextLine
''''''''''''''    Do Until EOF(1)
''''''''''''''        NextLine = Input(1, #1)
''''''''''''''        If NextLine = Chr(10) Or NextLine = Chr(13) Then
''''''''''''''            NextLine = "*"
''''''''''''''        End If
''''''''''''''        Label1.Caption = Label1.Caption + NextLine
''''''''''''''        sfile = sfile + NextLine
''''''''''''''    Loop
''''''''''''''    Label1.Caption = Right(sfile, 400)
'    Close #1
'
'    Open App.Path & "\" & FileUser For Input As #1
'    Do Until EOF(1)
'        NextLine = NextLine + Input(1, #1)
'    Loop
''    For i = 1 To 1500
''        NextLine = NextLine + Input(1, #1)
''    Next i
''''''''''''''    Close #1
''    Label1.Caption = "URAA " + GetToken("coord14=", NextLine)
''    Label2.Caption = "Button has been pressed!"
End Sub

Private Sub Command2_Click()
    Call mnuWUInfo_Click
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then        ' Check if right mouse button was clicked.
      PopupMenu mnuStats     ' Display the Stats menu as a pop-up menu.
   End If
End Sub

Private Sub Form_Load()
    'Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    'Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    'Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    'Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    Command1.ToolTipText = "Показать на карте координаты обработанных блоков"
    Command2.ToolTipText = "Показать информацию о текущем рабочем блоке"
    Label1.Caption = "Используйте ПРАВУЮ кнопку мыши на карте для вызова меню, ЛЕВУЮ - для получения краткой информации о звезде." + vbCrLf
    Label1.Caption = Label1.Caption & "Используйте Shift+ЛЕВУЮ кнопку мыши на карте для подсветки звезд, для которых можно вывести краткую справку."
    
    Form_Resize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    StatusStr.Caption = "Ready..."  'Thats damned status bar control doesn't work!
                                    'I'd forced to use text labels instead...
                                    'What is goin' on here? Who can tell me?
End Sub

Private Sub Form_Resize()
Dim NewFrmWidth As Long     'Variable defined in reason of very slow controls
                            'properties access methods. We'll replace them with
                            'simple math operations and keep our Pentium-II
                            'happy with his favourite Long 32-bit num's  :)
Dim NewFrmHeight As Long    'Same thing!

    NewFrmWidth = frmMain.ScaleWidth
    NewFrmHeight = frmMain.ScaleHeight
    Line1.X2 = NewFrmWidth
    Line2.X2 = NewFrmWidth
    Line3.X2 = NewFrmWidth - 100
    Line4.X2 = NewFrmWidth - 100
    Line6.X1 = NewFrmWidth - 100
    Line6.X2 = NewFrmWidth - 100
    Label2.Width = (NewFrmWidth - 100) - Line5.X1 - 100
    
    'Загрузка карты неба
    If (Dir(App.path & "\skymap.bmp") <> "") Then
        Set Picture1.Picture = LoadPicture(App.path & "\skymap.bmp")
    Else
        Result = MsgBox("Файл звездного неба (skymap.bmp) не обнаружен." + vbCrLf + "Возможно, он находится в другой папке и/или диске" + vbCrLf + "Хотите чтобы SETImap поискала этот файл на Вашем компьютере?", vbYesNo + vbExclamation, "Файл не найден")
        'TO DO Здесь провести поиск и если не найдено ничего, установить Result=vbNo
        If Result = vbNo Then
            Result = MsgBox("SETImap может открыть страницу на сервере SETI@home, " + vbCrLf + "на которой находится необходимый файл." + vbCrLf + "Хотите чтобы SETImap подключилась к Интернету?", vbYesNo + vbQuestion, "Файл не найден")
                If Result = vbYes Then
                    Result = MsgBox("Пожалуйста, сохраните загруженную карту как файл с именем" + vbCrLf + """skymap.bmp"" (Bitmap filetype) в папке SETImap", vbOKOnly + vbInformation, "Файл не найден")
                    Dim RetVal As Double
                    RetVal = Shell("Start http://www.setiathome.com", vbMaximizedFocus)
                Else
                    Result = MsgBox("Старт SETImap без карты звездного неба.", vbOKOnly + vbInformation, "Файл не найден")
                End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = Forms.Count To 1 Step -1
        Unload Forms(i - 1)
    Next
    
    'close all sub forms
    'For i = Forms.Count - 1 To 1 Step -1
        'Unload Forms(i)
    'Next
    
        'If Me.WindowState <> vbMinimized Then
        'SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        'SaveSetting App.Title, "Settings", "MainTop", Me.Top
        'SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        'SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    'End If
    ''SaveSetting App.Title, "Settings", "NumOfHistoryRec", RegRecords
    ''SaveSetting App.Title, "Settings", "LastRecordNum", LastRecordNum
    SaveRegSettings
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuClear_Click()
    Picture1.Cls
End Sub

Private Sub mnuContents_Click()
'Open Help and display the Contents screen
''    With dlgCommonDialog
''        .HelpCommand = cdlHelpContents
''        .HelpFile = App.HelpFile
''        .ShowHelp
''    End With
    ShowCHMHelp
End Sub

Private Sub mnuExit_Click()
    Set WU = Nothing
    Set State = Nothing
    Set UserInfo = Nothing
    Set OutResult = Nothing
    Set StatusStr = Nothing
    ''Set Spike = Nothing
    ''Set Gauss = Nothing
    Unload Me
    End
End Sub

Private Sub mnuGraphs_Click()
    Load frmGport
    frmGport.Show
End Sub

Private Sub mnuHelpOnHelp_Click()
    With dlgCommonDialog
        .HelpCommand = cdlHelpHelpOnHelp
        .HelpFile = App.HelpFile
        .ShowHelp
    End With
End Sub

Private Sub mnuHGraphs_Click()
    mnuGraphs_Click
End Sub

Private Sub mnuHHistory_Click()
    mnuHistory_Click
End Sub

Private Sub mnuHHistoryState_Click()
    mnuHistoryState_Click
End Sub

Private Sub mnuHistory_Click()
    StatusStr.Caption = "Пожалуйста, подождите. Загрузка журнала может занять несколько секунд..."
    StatusStr.Refresh
    showWU = True
    Load frmHistory
    frmHistory.Show
End Sub

Private Sub mnuHistoryState_Click()
    StatusStr.Caption = "Пожалуйста, подождите. Загрузка журнала может занять несколько секунд..."
    StatusStr.Refresh
    showWU = False
    Load frmHistory
    frmHistory.Show
End Sub

Private Sub mnuHTheory_Click()
    mnuTheory_Click
End Sub

Private Sub mnuHTopGauss_Click()
    mnuTopGauss_Click
End Sub

Private Sub mnuHTopSpikes_Click()
    mnuTopSpikes_Click
End Sub

Private Sub mnuHUserInfo_Click()
    mnuUserInfo_Click
End Sub

Private Sub mnuHWUInfo_Click()
    mnuWUInfo_Click
End Sub

Private Sub mnuHZoom_Click()
    mnuZoom_Click
End Sub

Private Sub mnuHClear_Click()
    mnuClear_Click
End Sub

Private Sub mnuRange_Click()
    'TO DO Menu...
    'MsgBox "Открыть окно калибровки приборов"
End Sub

Private Sub mnuShowAllWU_Click()
    Call Command1_Click
End Sub

Private Sub mnuShowBorders_Click()
    Picture1.ForeColor = vbWhite
    Picture1.Line (Picture1.ScaleLeft, 2100)-(Picture1.ScaleLeft + Picture1.ScaleWidth, 2100)
    Picture1.Line (Picture1.ScaleLeft, 3690)-(Picture1.ScaleLeft + Picture1.ScaleWidth, 3690)
End Sub

Private Sub mnuTheory_Click()
    Load frmMport
    frmMport.Show
End Sub

Private Sub mnuTopGauss_Click()
    Load frmStats
    frmStats.Option1(1).Value = True
    frmStats.Show
End Sub

Private Sub mnuTopSpikes_Click()
    Load frmStats
    frmStats.Option1(0).Value = True
    frmStats.Show
End Sub

Private Sub mnuUserInfo_Click()
    Load frmUser
    frmUser.Show vbModal, Me
End Sub

Private Sub mnuVisuals_Click()
    Load frmOptions
    frmOptions.Show vbModal, Me
End Sub

'**********************************************************
'*    Показывает на карте места, сигналы из которых       *
'*           несли в себе лучшие гауссианы                *
'**********************************************************
Private Sub mnuWhereGaussians_Click()
Dim i As Long
Dim strTmp As String
Dim Total As Long           'Всего сигналов
Dim Matched As Long         'Из них обработано лучших

    Total = State.GetLastRecNum(3)
    Matched = 0
    'For i = 1 To Total
    For i = Total To 1 Step -1
        If State.ReadIndex(1, i) Then   'Target=1 - гауссианы
            strTmp = Format((TopG.average * 10), "0.0000000")
            If Val(Left(strTmp, 1)) > 2 Then
                WU.ClearAll (0)
                WU.DecodeHistory (WU.ReadHistory(TopG.ID, 0))
                'ShowPosition WU.StartRA, WU.StartDEC, "", "", True, 2
                ShowPosition WU.StartRA, WU.StartDEC, Format((TopG.average * 100), "0.00000"), "", True, 2
                Matched = Matched + 1
            End If
        End If
    StatusStr.Caption = "Обработано " + Str(Total - i) + " результатов (" + Str(((Total - i) * 100) \ Total) + " %), из них лучших " + Str(Matched) + " (" + Str((Matched * 100) \ Total) + " %)"
    StatusStr.Refresh
    Next i
    StatusStr.Caption = "Показать источники лучших гауссиан: выполнено 100%. Показано всего источников: " + Str(Matched) + " (" + Str((Matched * 100) \ Total) + " % от общего числа результатов)."
    StatusStr.Refresh
End Sub

'**********************************************************
'*    Показывает на карте места, сигналы из которых       *
'*         несли в себе лучшие пиковые сигналы            *
'**********************************************************
Private Sub mnuWhereSpikes_Click()
Dim i As Long
Dim Total As Long           'Всего сигналов
Dim Matched As Long         'Из них обработано лучших
    
    Total = State.GetLastRecNum(4)
    Matched = 0
    ''For i = 1 To Total
    For i = Total To 1 Step -1
        If State.ReadIndex(0, i) Then   'Target=0 - пики
            If Int(TopS.power) > 200 Then
                WU.ClearAll (0)
                WU.DecodeHistory (WU.ReadHistory(TopS.ID, 0))
                'ShowPosition WU.StartRA, WU.StartDEC, "", "", True, 3
                ShowPosition WU.StartRA, WU.StartDEC, Format(TopS.power, "0.00000"), "", True, 3
                Matched = Matched + 1
            End If
        End If
    StatusStr.Caption = "Обработано " + Str(Total - i) + " результатов (" + Str(((Total - i) * 100) \ Total) + " %), из них лучших " + Str(Matched) + " (" + Str((Matched * 100) \ Total) + " %)"
    StatusStr.Refresh
    Next i
    StatusStr.Caption = "Показать источники лучших пиков: выполнено 100%. Показано всего источников: " + Str(Matched) + " (" + Str((Matched * 100) \ Total) + " % от общего числа результатов)."
    StatusStr.Refresh
End Sub

Private Sub mnuWUInfo_Click()
    Load frmViewWU
    frmViewWU.Show
End Sub

Private Sub mnuZoom_Click()
    'TO DO Menu...
    'MsgBox "Сделать zoom"
End Sub

Private Sub mWU_WriteComplete()
    Result = MsgBox("Операция записи журнала успешно завершена!", vbOKOnly, "Запись журнала")
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim TwipsPerGrad As Double      'Число твипсов в одном градусе
Dim TwipsPerDEC As Double       'Число твипсов в одном градусе склонения
Dim dec As Double, ra As Double
    TwipsPerGrad = Picture1.ScaleWidth / 24
    TwipsPerDEC = Picture1.ScaleHeight / 180
    dec = 90 - Y / TwipsPerDEC
    If x < Picture1.ScaleWidth / 2 Then
        ra = 12 - x / TwipsPerGrad
    Else
        ra = 36 - x / TwipsPerGrad
    End If
    StatusStr.Caption = "Угол " & DecodeRA(ra) & "| Склонение " & DecodeDEC(dec)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim message As String
    If Button = 2 Then       ' Check if right mouse button was clicked.
        PopupMenu mnuHideOnMap     ' Display the OnMap menu as a pop-up menu.
    End If
    If Button = 1 Then
        If Shift = 0 Then
            'Result = MsgBox("Левая кнопка! " & "X=" & Str(X) & " Y=" & Str(Y), vbOKOnly, "Ого! Shift=0")
        Else
            'Подсветка звезд
            'Result = MsgBox("Левая кнопка! " & "X=" & Str(X) & " Y=" & Str(Y), vbOKOnly, "Ого! Shift= " & Str(Shift))
            Picture1.ForeColor = vbRed
            Picture1.Circle (3690, 3315), 100
            Picture1.Circle (3690, 3315), 75
            Picture1.Circle (3690, 3315), 50
            Picture1.Circle (7125, 3015), 100
            Picture1.Circle (7125, 3015), 75
            Picture1.Circle (7125, 3015), 50
        End If
        If x > 3640 And x < 3740 Then
            If Y > 3265 And Y < 3365 Then
            message = "Правое плечо Ориона образует звезда Бетельгейзе (от арабского ""Beit Algueze"")," + vbCrLf
            message = message & "что означает ""armpit of the giant"". Звезда находится на расстоянии" + vbCrLf
            message = message & "520 световых лет от Земли, а ее свет начал свой путь незадолго до" + vbCrLf
            message = message & "путешествия Колумба. Диаметр звезды по оценкам составляет 480...800 миль," + vbCrLf
            message = message & "делая ее одной из самых больших звезд, видимых невооруженным глазом." + vbCrLf
            message = message & "Если поместить Бетельгейзе на место нашего Солнца, то внешние слои звезды" + vbCrLf
            message = message & "находились бы за орбитой Марса."
            Result = MsgBox(message, vbOKOnly, "Звезда - информация")
            End If
        End If
        If x > 7075 And x < 7175 Then
            If Y > 2965 And Y < 3065 Then
            message = "Algenib созвездия Пегаса" + vbCrLf
            message = message & "Расстояние - 479 световых лет."
            Result = MsgBox(message, vbOKOnly, "Звезда - информация")
            End If
        End If
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Save"
            'ToDo: Add 'Save' button code.
            MsgBox "Add 'Save' button code."
        Case "Print"
            'ToDo: Add 'Print' button code.
            MsgBox "Add 'Print' button code."
        Case "Macro"
            'ToDo: Add 'Macro' button code.
            MsgBox "Add 'Macro' button code."
        Case "Properties"
            'MsgBox "Add 'Properties' button code."
            Call mnuVisuals_Click
        Case "Help"
            'ToDo: Add 'Help' button code.
            MsgBox "Add 'Help' button code."
        Case "Help What's This"
            'ToDo: Add 'Help What's This' button code.
            MsgBox "Add 'Help What's This' button code."
    End Select
End Sub

Function ShowPosition(pstart_ra As String, pstart_dec As String, p_cmt As String, comment As String, pcurr As Boolean, Mode As Long) As Boolean
Dim ix As Integer, iy As Integer
Dim ra As Double, dec As Double
Dim TwipsPerGrad As Double      'Число твипсов в одном градусе
Dim TwipsPerDEC As Double      'Число твипсов в одном градусе склонения
Dim text As String
Dim i As Long

Dim iii As Long

    i = 0
    TwipsPerGrad = Picture1.ScaleWidth / 24
    TwipsPerDEC = Picture1.ScaleHeight / 180
    ra = MyStrToFloat(pstart_ra)
    dec = MyStrToFloat(pstart_dec)
    Select Case Mode
        Case 1:
            If ra < 12 Then
                'Точка на левой половине карты
                ix = (Picture1.ScaleWidth) / 2 - Trunc(ra * TwipsPerGrad)
            Else
                'Точка на правой половине карты
                ix = (Picture1.ScaleWidth) / 2 + Trunc((24 - ra) * TwipsPerGrad)
            End If
            iy = (Picture1.ScaleHeight) / 2 - Trunc(dec * TwipsPerDEC)
            If pcurr Then
                Picture1.ForeColor = vbRed
            Else
                Picture1.ForeColor = vbYellow
            End If
            Picture1.DrawWidth = 1
            ix = ix
            iy = iy
            
            Select Case MarkerType
                Case 0: 'Крест
                    Picture1.Line (ix - 100 - 50 * MarkerSize, iy)-(ix + 100 + 50 * MarkerSize, iy)
                    Picture1.Line (ix, iy - 100 - 50 * MarkerSize)-(ix, iy + 100 + 50 * MarkerSize)
                Case 1: 'Окружность
                    Picture1.Circle (CInt(ix), CInt(iy)), 50 + 50 * MarkerSize
                Case 2: 'Треугольник
                    Picture1.Line (ix - 50 - 50 * MarkerSize, iy + 75 + 50 * MarkerSize)-(ix, iy - 60 - 50 * MarkerSize)
                    Picture1.Line (ix, iy - 60 - 50 * MarkerSize)-(ix + 50 + 50 * MarkerSize, iy + 75 + 50 * MarkerSize)
                    Picture1.Line (ix + 50 + 50 * MarkerSize, iy + 75 + 50 * MarkerSize)-(ix - 50 - 50 * MarkerSize, iy + 75 + 50 * MarkerSize)
            End Select
        Case 2:
            If ra < 12 Then
                'Точка на левой половине карты
                ix = (Picture1.ScaleWidth) / 2 - Trunc(ra * TwipsPerGrad)
            Else
                'Точка на правой половине карты
                ix = (Picture1.ScaleWidth) / 2 + Trunc((24 - ra) * TwipsPerGrad)
            End If
            iy = (Picture1.ScaleHeight) / 2 - Trunc(dec * TwipsPerDEC)
iii = Val(Left(p_cmt, 2))
If iii > 40 Then
    iii = 40
End If
iii = iii - 17
            For i = 1 To 200
                'Picture1.ForeColor = RGB(100 - Int(i / 2), 100 - Int(i / 2), 255 - i)
Picture1.ForeColor = RGB(82 + iii * 5 - Int(i / 2), 82 + iii * 5 - Int(i / 2), 255 - i)
                Picture1.Circle (CInt(ix), CInt(iy)), i
            Next i
        Case 3:
            If ra < 12 Then
                'Точка на левой половине карты
                ix = (Picture1.ScaleWidth) / 2 - Trunc(ra * TwipsPerGrad)
            Else
                'Точка на правой половине карты
                ix = (Picture1.ScaleWidth) / 2 + Trunc((24 - ra) * TwipsPerGrad)
            End If
            iy = (Picture1.ScaleHeight) / 2 - Trunc(dec * TwipsPerDEC)


iii = Val(Left(p_cmt, InStr(1, p_cmt, ",", vbTextCompare))) / 10
If iii > 40 Then
    iii = 40
End If
iii = iii - 19
            
            For i = 1 To 200
                'Picture1.ForeColor = RGB(100 - Int(i / 2), 255 - i, 100 - Int(i / 2))
Picture1.ForeColor = RGB(95 + iii * 5 - Int(i / 2), 255 - i, 95 + iii * 5 - Int(i / 2))
                Picture1.Circle (CInt(ix), CInt(iy)), i
            Next i
    End Select
    ''If ShowWUNumber1.Checked Then begin
      ''text:=p_cmt;
    ''end;
    ''If ShowComment.Checked Then begin
      ''If ShowWUNumber1.Checked Then begin
        ''text:=text+' ';
      ''end;
      ''text:=text+comment;
    ''end;
    ''StarMap.Canvas.Font.Color:=clYellow;
    ''StarMap.Canvas.textout(ix+6+s, iy+6+s, text);
  ''except
    ''MessageDlg('An error occured while trying to show the following coordinates on the map:'+chr(13)+
               '''RA = '+pstart_ra+chr(13)+
               '''DEC = '+pstart_dec+chr(13),
               ''mtError,[mbOK], 0);
  ''end;
''//  StatusBar1.Panels[0].Text:=IntToStr(TRUNC(ra))+' hr '+IntToStr(TRUNC((RA-TRUNC(RA))*60))+' min RA  |  '
''//                    +IntToStr(TRUNC(DEC))+' deg '+IntToStr(TRUNC((DEC-TRUNC(DEC))*60))+' min DEC';;
''end;
    'Picture1.DrawWidth = 1
End Function

Private Sub tbToolBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then       ' Check if right mouse button was clicked.
      PopupMenu mnuStats     ' Display the Stats menu as a pop-up menu.
   End If
End Sub

Public Sub RunServices()
    If RedrawOnStartup Then
        frmMain.Picture1.AutoRedraw = True
        frmMain.AutoRedraw = True
    Else
        frmMain.Picture1.AutoRedraw = False
        frmMain.AutoRedraw = False
    End If
    If AutoShowWU = 1 Then
        Command2.Value = True   'Fire the Click event
        ''Call mnuWUInfo_Click
    End If
End Sub
