VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistory 
   Caption         =   "Журнал"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5370
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   9472
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8715
      TabIndex        =   3
      Top             =   5670
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Редактировать"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5670
      TabIndex        =   2
      Top             =   5670
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Добавить пропущенную запись"
      Height          =   435
      Left            =   2310
      TabIndex        =   1
      Top             =   5670
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Показ"
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   5670
      Width           =   1695
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Привязывает к DataGrid другому источнику (более полному) и наоборот
    showWU = Not (showWU)
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    BindDataSource
    Command1.Enabled = True
    Command4.Enabled = True
    If showWU Then
        Me.Caption = "Журнал учета блоков"
        Command1.Caption = "Показ результатов"
        Command2.Enabled = True
    Else
        Me.Caption = "Журнал результатов обработки блоков"
        Command1.Caption = "Показ блоков"
    End If
End Sub

Private Sub Command2_Click()
    Load frmInput
    frmInput.Show vbModal, Me
    If EditMode Or NewMode Then
        Load frmEdit
        frmEdit.Show vbModal, Me
        frmHistory.DataGrid1.ReBind
        frmHistory.DataGrid1.Refresh
        'TO DO Зачем это писать дважды?
        ''frmHistory.DataGrid1.ReBind
        ''frmHistory.DataGrid1.Refresh
        InitForm (0)    'Сейчас режим показа рабочих блоков
    End If
End Sub

'Нажата кнопка "РЕДАКТИРОВАТЬ"
Private Sub Command3_Click()
    If showWU Then
        'Загрузка редактора рабочих блоков
        WU.ClearAll (0)     'Стереть все, чтобы нормально инициировать окно редактора
        If Not (WU.DecodeHistory(WU.ReadHistory(EditID, 1))) Then
            Result = MsgBox("Ошибка чтения журнала. Операция отменена.", vbOKOnly, "Ошибка")
            Debug.Print "EditID" & EditID
        Else
            EditMode = True
            Load frmEdit
            frmEdit.Show vbModal, Me
            'Обновить привязку. Надо ли это делать?
            frmHistory.DataGrid1.ReBind
            frmHistory.DataGrid1.Refresh
            InitForm (0)   'Сейчас режим показа рабочих блоков
        End If
    Else
        'загрузка редактора результатов
        If State.ReadHistory(EditID) = 0 Then
            Load frmREditor
            frmREditor.Show vbModal, Me
            'Обновить привязку. Надо ли это делать?
            frmHistory.DataGrid1.ReBind
            frmHistory.DataGrid1.Refresh
            InitForm (1)   'Сейчас режим показа результатов
        End If
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Print the Text, row, and column of the cell the user clicked.
    Debug.Print DataGrid1.text; DataGrid1.Row; DataGrid1.Col; LastRow; LastCol
    'получить номер выделенной стороки и ID
    EditRowNum = DataGrid1.Row
    DataGrid1.Col = 0
    EditID = CLng(Val(DataGrid1.text))
    Debug.Print "Getting WU ID: " & EditID
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

'**********************************************************
'*                 Инициализация колонок                  *
'*  Mode=0: режим показа журнала рабочих блоков           *
'*  Mode=1: режим показа журнала результатов (State)      *
'**********************************************************
Private Sub InitForm(ByVal Mode As Long)
' TO DO сделать чтение ширины из реестра и обязательно присваивать эти значения переменным
    Select Case Mode
        Case 0:    'Сейчас режим показа рабочих блоков
            DataGrid1.Columns(0).Width = 500
            DataGrid1.Columns(1).Width = 1500
            DataGrid1.Columns(2).Width = 3000
            DataGrid1.Columns(3).Width = 2100
            DataGrid1.Columns(4).Width = 1500
            DataGrid1.Columns(5).Width = 1000
            DataGrid1.Columns(6).Width = 1000
            DataGrid1.Columns(7).Width = 1000
            DataGrid1.Columns(8).Width = 1000
            DataGrid1.Columns(9).Width = 1000
            DataGrid1.Columns(10).Width = 1000
        Case 1:
            DataGrid1.Columns(0).Width = 500
            DataGrid1.Columns(1).Width = 750
            DataGrid1.Columns(2).Width = 1200
            DataGrid1.Columns(3).Width = 1100
            DataGrid1.Columns(4).Width = 1500
            DataGrid1.Columns(5).Width = 1000
            DataGrid1.Columns(6).Width = 1300
            DataGrid1.Columns(7).Width = 1200
            DataGrid1.Columns(8).Width = 1200
            DataGrid1.Columns(9).Width = 1200
            DataGrid1.Columns(10).Width = 1200
            DataGrid1.Columns(11).Width = 1200
            DataGrid1.Columns(12).Width = 1200
    End Select
End Sub

Private Sub BindDataSource()
    If showWU Then
        Me.Caption = "Журнал учета блоков"
        Command1.Caption = "Показ результатов"
        If Not (WUbind) Then
            WUbind = Not (WUbind)   'Поднять флаг запрещения повторной привязки (для исключения дублирования записей)
            WU.AddRecord 0  'Загрузить ВСЕ Сведения из журнала
        End If
        Set DataGrid1.DataSource = WU
        Call InitForm(0)    'Сейчас режим показа рабочих блоков
    Else
        Me.Caption = "Журнал результатов обработки блоков"
        Command1.Caption = "Показ блоков"
        Command2.Enabled = False
        If Not (Sbind) Then
            Sbind = Not (Sbind)     'Поднять флаг запрещения повторной привязки (для исключения дублирования записей)
            State.AddRecord 0  'Загрузить ВСЕ Сведения из журнала
        End If
        Set DataGrid1.DataSource = State
        Call InitForm(1)    'Сейчас режим показа результатов
    End If
End Sub

Private Sub Form_Load()
    '' ' Create a new NamesData Object
    ''Set datNames = New NamesData
    ''Нам это не надо, т.к. классы инициализируются в модуле при старте программы!
    
    ' Bind the DataGrid to the new DataSource datNames
    BindDataSource
    StatusStr.Caption = "Журнал загружен."
    StatusStr.Refresh
End Sub

Private Sub Form_Resize()
    If frmHistory.ScaleWidth < 10500 Then
        frmHistory.Width = 10500
    End If
    If frmHistory.ScaleHeight < 5000 Then
        frmHistory.Height = 5000
    End If
    DataGrid1.Width = frmHistory.ScaleWidth - 175
    DataGrid1.Height = frmHistory.ScaleHeight - 1000
    Command1.Top = frmHistory.ScaleHeight - 600
    Command2.Top = frmHistory.ScaleHeight - 600
    Command3.Top = frmHistory.ScaleHeight - 600
    Command4.Top = frmHistory.ScaleHeight - 600
    Command1.Left = (frmHistory.ScaleWidth - 10410) \ 2 + 210
    Command2.Left = Command1.Left + 2100
    Command3.Left = Command2.Left + 3360
    Command4.Left = Command3.Left + 3045
End Sub
