VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class is represent 1 work unit header"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

'local variable(s) to hold property value(s)
Private mvarStartRA As String 'local copy
Private mvarStartDEC As String 'local copy
Private mvarEndRA As String 'local copy
Private mvarEndDEC As String 'local copy
Private mvarUnitName As String 'local copy
Private mvarVersion As String 'local copy
Private mvarTimeOfRec As String 'local copy
Private mvarSubbandCenter As String 'local copy
Private mvarSubbandBase As String 'local copy
Private mvarSubbandRate As String 'local copy
Private mvarSubbandNum As String 'local copy
Private mvarReceiver As String 'local copy
Private mvarNsamples As String 'local copy
Private mvarTapeVer As String 'local copy
Private mvarNumPositions As String 'local copy
Private mvarAngleRange As String 'local copy
Private mvarComments As String 'local copy              'КОММЕНТАРИЙ
Private mvarNumID As String 'local copy
Private mvarType_of_unit As String 'local copy
Private mvarTask As String 'local copy
Private mvarData_type As String 'local copy
Private mvarData_class As String 'local copy
Private mvarSplitter_version As String 'local copy
Private mvarfft_len As String 'local copy
Private mvarifft_len As String 'local copy
Private mvarReserve1 As String 'local copy
Private mvarReserve2 As String 'local copy
Private mvarReserve3 As String 'local copy

Private mvarexistWU As Boolean 'local copy
Private mvarhFile As Long 'local copy

Private Coords(21) As String
'Private existWU As Boolean
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent WriteComplete[(arg1, arg2, ... , argn)]
Public Event WriteComplete()
Private WithEvents rsWUhistory As ADODB.Recordset
Attribute rsWUhistory.VB_VarHelpID = -1

Public Property Let existWU(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.existWU = 5
    mvarexistWU = vData
End Property

Public Property Get existWU() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.existWU
    existWU = mvarexistWU
End Property

Public Property Let Reserve3(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Reserve3 = 5
    mvarReserve3 = vData
End Property

Public Property Get Reserve3() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Reserve3
    Reserve3 = mvarReserve3
End Property

Public Property Let Reserve2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Reserve2 = 5
    mvarReserve2 = vData
End Property

Public Property Get Reserve2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Reserve2
    Reserve2 = mvarReserve2
End Property

Public Property Let Reserve1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Reserve1 = 5
    mvarReserve1 = vData
End Property

Public Property Get Reserve1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Reserve1
    Reserve1 = mvarReserve1
End Property

Public Property Let ifft_len(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ifft_len = 5
    mvarifft_len = vData
End Property

Public Property Get ifft_len() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ifft_len
    ifft_len = mvarifft_len
End Property

Public Property Let fft_len(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fft_len = 5
    mvarfft_len = vData
End Property

Public Property Get fft_len() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fft_len
    fft_len = mvarfft_len
End Property

Public Property Let Splitter_version(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Splitter_version = 5
    mvarSplitter_version = vData
End Property

Public Property Get Splitter_version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Splitter_version
    Splitter_version = mvarSplitter_version
End Property

Public Property Let Data_class(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Data_class = 5
    mvarData_class = vData
End Property

Public Property Get Data_class() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Data_class
    Data_class = mvarData_class
End Property

Public Property Let Data_type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Data_type = 5
    mvarData_type = vData
End Property

Public Property Get Data_type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Data_type
    Data_type = mvarData_type
End Property

Public Property Let Task(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Task = 5
    mvarTask = vData
End Property

Public Property Get Task() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Task
    Task = mvarTask
End Property

Public Property Let Type_of_unit(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type_of_unit = 5
    mvarType_of_unit = vData
End Property

Public Property Get Type_of_unit() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Type_of_unit
    Type_of_unit = mvarType_of_unit
End Property

Public Property Let NumID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumID = 5
    mvarNumID = vData
End Property

Public Property Get NumID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumID
    NumID = mvarNumID
End Property

Public Property Let NumPositions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumPositions = 5
    mvarNumPositions = vData
End Property

Public Property Get NumPositions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumPositions
    NumPositions = mvarNumPositions
End Property

Public Property Let hFile(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hFile = 5
    mvarhFile = vData
End Property

Public Property Get hFile() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hFile
    hFile = mvarhFile
End Property

Public Property Let Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comments = 5
    mvarComments = vData
End Property

Public Property Get Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    Comments = mvarComments
End Property

Public Property Let AngleRange(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AngleRange = 5
    mvarAngleRange = vData
End Property

Public Property Get AngleRange() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AngleRange
    AngleRange = mvarAngleRange
End Property

Public Property Let TapeVer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TapeVer = 5
    mvarTapeVer = vData
End Property

Public Property Get TapeVer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TapeVer
    TapeVer = mvarTapeVer
End Property

Public Property Let Nsamples(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nsamples = 5
    mvarNsamples = vData
End Property

Public Property Get Nsamples() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nsamples
    Nsamples = mvarNsamples
End Property

Public Property Let Receiver(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Receiver = 5
    mvarReceiver = vData
End Property

Public Property Get Receiver() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Receiver
    Receiver = mvarReceiver
End Property

Public Property Let SubbandNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubbandNum = 5
    mvarSubbandNum = vData
End Property

Public Property Get SubbandNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubbandNum
    SubbandNum = mvarSubbandNum
End Property

Public Property Let SubbandRate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubbandRate = 5
    mvarSubbandRate = vData
End Property

Public Property Get SubbandRate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubbandRate
    SubbandRate = mvarSubbandRate
End Property

Public Property Let SubbandBase(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubbandBase = 5
    mvarSubbandBase = vData
End Property

Public Property Get SubbandBase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubbandBase
    SubbandBase = mvarSubbandBase
End Property

Public Property Let SubbandCenter(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubbandCenter = 5
    mvarSubbandCenter = vData
End Property

Public Property Get SubbandCenter() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubbandCenter
    SubbandCenter = mvarSubbandCenter
End Property

Public Property Let TimeOfRec(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeOfRec = 5
    mvarTimeOfRec = vData
End Property

Public Property Get TimeOfRec() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeOfRec
    TimeOfRec = mvarTimeOfRec
End Property

Public Property Let Version(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Version = 5
    mvarVersion = vData
End Property

Public Property Get Version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property

Public Property Let UnitName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitName = 5
    mvarUnitName = vData
End Property

Public Property Get UnitName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitName
    UnitName = mvarUnitName
End Property

Public Property Let EndDEC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndDEC = 5
    mvarEndDEC = vData
End Property

Public Property Get EndDEC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndDEC
    EndDEC = mvarEndDEC
End Property

Public Property Let EndRA(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndRA = 5
    mvarEndRA = vData
End Property

Public Property Get EndRA() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndRA
    EndRA = mvarEndRA
End Property

Public Property Let StartDEC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartDEC = 5
    mvarStartDEC = vData
End Property

Public Property Get StartDEC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartDEC
    StartDEC = mvarStartDEC
End Property

Public Property Let StartRA(ByVal vData As String)
Attribute StartRA.VB_Description = "Starting angle"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartRA = 5
    mvarStartRA = vData
End Property

Public Property Get StartRA() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartRA
    StartRA = mvarStartRA
End Property

'**********************************************************
'*           Читает значения массива координат            *
'**********************************************************
Public Function GetCoordX(Index As Long) As String
    On Error GoTo GetCoordXErr
    If Index >= 0 Then
        If Index < 22 Then
            GetCoordX = Coords(Index)
        Else
            GetCoordX = ""
        End If
    Else
        GetCoordX = ""
    End If
Exit Function
GetCoordXErr:

    Call RaiseError(MyUnhandledError, "cWU:GetCoordX Method")
End Function

'**********************************************************
'*       Устанавливает значения массива координат         *
'**********************************************************
Public Function SetCoordX(sfile As String, Index As Long) As Boolean
    On Error GoTo SetCoordXErr
    If Index >= 0 Then
        If Index < 22 Then
            Coords(Index) = sfile
            SetCoordX = True
        Else
            SetCoordX = False
        End If
    Else
        SetCoordX = False
    End If
Exit Function
SetCoordXErr:

    Call RaiseError(MyUnhandledError, "cWU:SetCoordX Method")
End Function

'**********************************************************
'*            Очищает некоторые поля объекта              *
'*         Mode = 0 - Очистить все поля объекта           *
'*            Mode = 1 - все кроме ID и Name              *
'**********************************************************
Public Function ClearAll(ByVal Mode As Long) As String
Dim i As Long
    On Error GoTo ClearAllErr
    
    Select Case Mode
        Case 0:     'Полная очистка
            NumID = "-1"
            UnitName = ""
        Case 1:     'Оставить ID и Name
    End Select
    Type_of_unit = ""
    Task = ""
    Version = ""
    Data_type = ""
    Data_class = ""
    Splitter_version = ""
    StartRA = ""
    StartDEC = ""
    EndRA = ""
    EndDEC = ""
    AngleRange = ""
    TimeOfRec = ""
    SubbandCenter = ""
    SubbandBase = ""
    SubbandRate = ""
    fft_len = ""
    ifft_len = ""
    SubbandNum = ""
    Receiver = ""
    Nsamples = ""
    TapeVer = ""
    NumPositions = ""
    For i = 0 To 21
        Coords(i) = ""
    Next i
    Reserve1 = ""
    Reserve2 = ""
    Reserve3 = ""
            
Exit Function
ClearAllErr:

    Call RaiseError(MyUnhandledError, "cWU:ClearAll Method")
End Function

'LINUX compatible
'**********************************************************
'*           Проверить существование файла                *
'*       Mode = 0 - путь к файлу журнала                  *
'*       Mode = 1 - путь к текущему блоку                 *
'*       Mode = 2 - путь к блоку клиента под Linux        *
'*       Mode = 3 - путь к файлу-индексу блоков           *
'**********************************************************
Public Function CheckFile(ByVal Mode As Long) As Boolean
Dim CheckResult As Boolean
    On Error GoTo CheckFileErr
    
    CheckResult = False
    If (Dir(GetFilePath(Mode), vbNormal) <> "") Then
        CheckResult = True
    End If
    CheckFile = CheckResult
    Exit Function
    
CheckFileErr:
    CheckFile = False
    Call RaiseError(MyUnhandledError, "cWU:CheckFile Method")
End Function

'LINUX compatible
'**********************************************************
'*  Получить полное имя файла с данными о текущем блоке   *
'*       Mode = 0 - путь к файлу журнала                  *
'*       Mode = 1 - путь к текущему блоку                 *
'*       Mode = 2 - путь к блоку клиента под Linux        *
'*       Mode = 3 - путь к файлу-индексу блоков           *
'**********************************************************
Public Function GetFilePath(ByVal Mode As Long) As String
    On Error GoTo GetFilePathErr
    
    Select Case Mode
        Case 0:     'Чтение из журнала
            GetFilePath = App.path & "\" & Datafile
        Case 1:     'Чтение непосредственно из WU
            GetFilePath = SETIpath & "\" & FileWU
        Case 2:     'Чтение блока, обрабатываемого клиентом под Linux
            GetFilePath = LinuxPath & "\" & FileWULinux
        Case 3:     'Путь к файлу-индексу блоков
            GetFilePath = App.path & "\" & IndexFileW
    End Select
    Exit Function
    
GetFilePathErr:

    Call RaiseError(MyUnhandledError, "cWU:GetFilePath Method")
End Function

'LINUX compatible
'EFRS (extended failures reporting system)
'**********************************************************
'*  Получить полное имя файла с данными о текущем блоке   *
'*       Mode = 0 - добавить только запись с номером ID   *
'*       Mode = 1 - построить индекс заново               *
'*       ID - номер блока (для Mode = 0)                  *
'*       (нужна еще проверка перезаписи)                  *
'*Сортировка - по номеру блока (т.е. не нужно сортировать)*
'*  WriteIndex = 0 - неизвестная ошибка                   *
'*  WriteIndex = 1 - успешное завершение                  *
'*  WriteIndex = 2 - записи с данным ID в журнале нет     *
'*  WriteIndex = 3 - ID не задан (режим Mode = 0)         *
'*  WriteIndex = 4 - ошибка при вызове DecodeHostory      *
'*  WriteIndex = 5 - файл журнала не найден               *
'**********************************************************
Public Function WriteIndex(ByVal Mode As Long, Optional ByVal ID As Long = -6) As Long
Dim hFileInd As Long        'указатель на файл SETItopw.dat
Dim sfile As String         'для чтения информации из журнала
Dim LastRecNum As Long
Dim i As Long, j As Long
Dim TMPrec As tTopW         'Временная запись-пустышка
    On Error GoTo WriteIndexErr
    
    WriteIndex = 0
    sfile = ""
    TMPrec.freq = 0         'Очистка полей записи-пустышки
    TMPrec.ID = -6
    TMPrec.StartDEC = 0
    TMPrec.StartRA = 0
    TMPrec.time = ""
    'Проверим файл-индекс
    If Not (CheckFile(3)) Then
        Mode = 1    'Файла индексов еще нет - строим заново
    End If
    Select Case Mode
        Case 0: 'Добавить (если надо) в существующий файл
        ''АЛГОРИТМ
        ''Если ID задан - прочитать запись из журнала, иначе - ошибка, выход из функции
        ''Искать в индексе запись с таким же ID, если ID < LastID
        ''Если найдена, то Put на ее место новую запись, иначе добавить (с учетом номера)
            If Not (ID = -6) Then
                'Индекс задан - проверяем длину файла
                sfile = ReadHistory(ID, 1)  'Читаем запись по индексу
                If Not (sfile = "") Then 'Чтение журнала успешно
                    If DecodeHistory(sfile) Then
                        TopW.ID = NumID
                        TopW.freq = SubbandBase
                        TopW.StartDEC = StartDEC
                        TopW.StartRA = StartRA
                        TopW.time = ExtractTime(TimeOfRec)
                        'Открыть файл индекса
                        hFileInd = FreeFile
                        Open GetFilePath(3) For Random As hFileInd Len = Len(TopW)
                        LastRecNum = LOF(hFileInd) \ Len(TopW)
                        If ID > LastRecNum Then
                            'ID дальше конца файла
                            If ((ID - LastRecNum) > 1) Then
                                'Заполнить промежуточные записи пустышками (ID=-6)
                                'будем использовать TMPrec
                                For i = LastRecNum + 1 To ID - 1
                                    Seek hFileInd, i
                                    Put hFileInd, , TMPrec
                                Next i
                                'Записать информацию
                            End If
                            'Теперь ID указывает на запись, следующую за последней
                        End If
                        'Подготовительная работа проведена (если надо было)
                        'записываем информацию!
                        Put hFileInd, ID, TopW
                        Close hFileInd
                        WriteIndex = 1 'успешное завершение
                    Else
                        WriteIndex = 4 'ошибка при вызове DecodeHostory
                    End If
                Else
                    WriteIndex = 2 'записи с данным ID в журнале нет
                End If
            Else
                WriteIndex = 3 'ID не задан (режим Mode = 0)
            End If
        Case 1: 'Построить индекс заново
            'Открываем журнал - загружаем записи ВРУЧНУЮ!!!
            'Не обращаемся к ReadHistory!!!
            If CheckFile(0) Then
                'Журнал обнаружен
                'Если индекс существует - стереть его
                If CheckFile(3) Then
                    Kill GetFilePath(3)
                End If
                hFile = FreeFile
                Open GetFilePath(0) For Input As hFile
                For i = 1 To RegRecords 'перебираем записи поштучно и определяем их ID
                    Line Input #hFile, sfile
                    If DecodeHistory(sfile) Then
                        'Подготавливаем запись для индекса
                        TopW.ID = CLng(Val(NumID))
                        TopW.freq = CLng(Val(SubbandBase))
                        TopW.StartDEC = CSng(Val(StartDEC))
                        TopW.StartRA = CSng(Val(StartRA))
                        TopW.time = ExtractTime(TimeOfRec)
                        'Открыть файл индекса
                        hFileInd = FreeFile
                        Open GetFilePath(3) For Random As hFileInd Len = Len(TopW)
                        LastRecNum = LOF(hFileInd) \ Len(TopW)
                        'Записываем индекс в соответствии с NumID записи в журнале
                        If NumID > LastRecNum Then
                            'ID дальше конца файла
                            If ((NumID - LastRecNum) > 1) Then
                                'Заполнить промежуточные записи пустышками (ID=-6)
                                'будем использовать TMPrec
                                For j = LastRecNum + 1 To NumID - 1
                                    Seek hFileInd, j
                                    Put hFileInd, , TMPrec
                                Next j
                                'Записать информацию
                            End If
                            'Теперь NumID указывает на запись, следующую за последней
                        End If
                        'Подготовительная работа проведена (если надо было)
                        'записываем информацию!
                        Put hFileInd, NumID, TopW
                        Close hFileInd
                        WriteIndex = 1 'успешное завершение
                    Else
                        WriteIndex = 4 'ошибка при вызове DecodeHostory
                        i = RegRecords
                    End If
                Next i
                Close hFile
                If WriteIndex < 2 Then
                    'Успешное завершение. Надеюсь, что все ошибки были перехвачены
                    'и WriteIndex=0 осталось с начала функции
                    WriteIndex = 1
                End If
            Else
                WriteIndex = 5 'файл журнала не найден
            End If
    End Select
    Exit Function
    
WriteIndexErr:
    WriteIndex = 0
    Call RaiseError(MyUnhandledError, "cWU:WriteIndex Method")
End Function

'LINUX compatible
'**********************************************************
'*      Читает запись из файла индекса SETItopw.dat       *
'*              ID содержит индекс записи                 *
'**********************************************************
Public Function ReadIndex(ByVal ID As Long) As Boolean
Dim hFileInd As Long        'указатель на файл SETItopw.dat
Dim LastRecNum As Long
    On Error GoTo ReadIndexErr
    
    ReadIndex = False
    If CheckFile(3) Then
        hFileInd = FreeFile
        Open GetFilePath(3) For Random As hFileInd Len = Len(TopW)
        LastRecNum = LOF(hFileInd) \ Len(TopW)
        If LastRecNum < ID Then 'ID больше, чем есть записей в файле!
            ReadIndex = False
        Else    'читать запись из файла
            Get hFileInd, ID, TopW
            ReadIndex = True
        End If
        Close hFileInd
    End If
    Exit Function
    
ReadIndexErr:
    ReadIndex = False
    Call RaiseError(MyUnhandledError, "cWU:ReadIndex Method")
End Function

'**********************************************************
'*    Прочитать заданный параметр из шифрованной записи   *
'**********************************************************
Public Function GetEncodedToken(ByVal tokenname As String, ByVal sfile As String) As String
Dim i As Long, j As Long
    On Error GoTo GetEncodedTokenErr
    
    GetEncodedToken = ""
    Select Case tokenname
        Case "NumID":
            i = InStr(2, sfile, "$", vbTextCompare)
            GetEncodedToken = Mid(sfile, 2, i - 2)
        Case "Name":
            i = 0
            For j = 1 To 5
                i = InStr(i + 1, sfile, "$", vbTextCompare)
            Next j
            j = InStr(i + 1, sfile, "$", vbTextCompare)
            GetEncodedToken = Mid(sfile, i + 1, j - i - 1)
        Case "Comments":
            i = 0
            For j = 1 To 25
                i = InStr(i + 1, sfile, "$", vbTextCompare)
            Next j
            j = InStr(i + 1, sfile, "$", vbTextCompare)
            GetEncodedToken = Mid(sfile, i + 1, j - i - 1)
    End Select
    Exit Function
    
GetEncodedTokenErr:
    Call RaiseError(MyUnhandledError, "cWU:GetEncodedToken Method")
End Function

'**********************************************************
'*         Прочитать заголовок из файла-журнала           *
'*             ID - номер записи в журнале                *
'*       Mode=0: ID - порядковый номер записи в файле     *
'*       Mode=1: ID - идентификационный номер блока       *
'**********************************************************
Public Function ReadHistory(ByVal ID As Long, ByVal Mode As Long) As String
Dim sfile As String
Dim i As Long
    On Error GoTo ReadHistoryErr
    sfile = ""
    If (Dir(GetFilePath(0), vbNormal) <> "") Then
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFilePath(0) For Input As hFile
        Select Case Mode
            Case 0:
                For i = 1 To ID
                    Line Input #hFile, sfile
                Next i
                ReadHistory = sfile
            Case 1:
                Do Until EOF(hFile)
                    Line Input #hFile, sfile
                    If GetEncodedToken("NumID", sfile) = CStr(ID) Then
                        ReadHistory = sfile
                    End If
                Loop
            End Select
        Close hFile
    Else
        ReadHistory = ""
    End If
    
    Exit Function
ReadHistoryErr:
    Call RaiseError(MyUnhandledError, "cWU:ReadHistory Method")
End Function

'**********************************************************
'*          Дешифрует запись из файла-журнала             *
'*      sfile - строка, полученная от ReadHistory         *
'**********************************************************
Public Function DecodeHistory(sfile As String) As Boolean
Dim i As Long, j As Long, x As Long
    On Error GoTo DecodeHistoryErr
        i = InStr(1, sfile, "$", vbTextCompare)
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        NumID = Mid(sfile, i + 1, j - i - 1)
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Type_of_unit = Mid(sfile, i + 1, j - i - 1)
        If Type_of_unit = "#" Then
            Type_of_unit = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Task = Mid(sfile, i + 1, j - i - 1)
        If Task = "#" Then
            Task = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Version = Mid(sfile, i + 1, j - i - 1)
        If Version = "#" Then
            Version = ""
        End If
       
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        UnitName = Mid(sfile, i + 1, j - i - 1)
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Data_type = Mid(sfile, i + 1, j - i - 1)
        If Data_type = "#" Then
            Data_type = ""
        End If
       
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Data_class = Mid(sfile, i + 1, j - i - 1)
        If Data_class = "#" Then
            Data_class = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Splitter_version = Mid(sfile, i + 1, j - i - 1)
        If Splitter_version = "#" Then
            Splitter_version = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        StartRA = Mid(sfile, i + 1, j - i - 1)
        If StartRA = "#" Then
            StartRA = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        StartDEC = Mid(sfile, i + 1, j - i - 1)
        If StartDEC = "#" Then
            StartDEC = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        EndRA = Mid(sfile, i + 1, j - i - 1)
        If EndRA = "#" Then
            EndRA = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        EndDEC = Mid(sfile, i + 1, j - i - 1)
        If EndDEC = "#" Then
            EndDEC = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        AngleRange = Mid(sfile, i + 1, j - i - 1)
        If AngleRange = "#" Then
            AngleRange = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        TimeOfRec = Mid(sfile, i + 1, j - i - 1)
        If TimeOfRec = "#" Then
            TimeOfRec = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        SubbandCenter = Mid(sfile, i + 1, j - i - 1)
        If SubbandCenter = "#" Then
            SubbandCenter = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        SubbandBase = Mid(sfile, i + 1, j - i - 1)
        If SubbandBase = "#" Then
            SubbandBase = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        SubbandRate = Mid(sfile, i + 1, j - i - 1)
        If SubbandRate = "#" Then
            SubbandRate = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        fft_len = Mid(sfile, i + 1, j - i - 1)
        If fft_len = "#" Then
            fft_len = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        ifft_len = Mid(sfile, i + 1, j - i - 1)
        If ifft_len = "#" Then
            ifft_len = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        SubbandNum = Mid(sfile, i + 1, j - i - 1)
        If SubbandNum = "#" Then
            SubbandNum = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Receiver = Mid(sfile, i + 1, j - i - 1)
         If Receiver = "#" Then
            Receiver = ""
        End If
       
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Nsamples = Mid(sfile, i + 1, j - i - 1)
        If Nsamples = "#" Then
            Nsamples = ""
        End If
       
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        TapeVer = Mid(sfile, i + 1, j - i - 1)
        If TapeVer = "#" Then
            TapeVer = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        NumPositions = Mid(sfile, i + 1, j - i - 1)
        If NumPositions = "#" Then
            NumPositions = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Comments = Mid(sfile, i + 1, j - i - 1)
        If Comments = "#" Then
            Comments = ""
        End If
        
        For x = 0 To 21
            i = j
            j = InStr(i + 1, sfile, "$", vbTextCompare)
            Coords(x) = Mid(sfile, i + 1, j - i - 1)
            If Coords(x) = "#" Then
                Coords(x) = ""
            End If
        Next x
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Reserve1 = Mid(sfile, i + 1, j - i - 1)
        If Reserve1 = "#" Then
            Reserve1 = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Reserve2 = Mid(sfile, i + 1, j - i - 1)
        If Reserve2 = "#" Then
            Reserve2 = ""
        End If
        
        i = j
        j = InStr(i + 1, sfile, "$", vbTextCompare)
        Reserve3 = Mid(sfile, i + 1, j - i - 1)
        If Reserve3 = "#" Then
            Reserve3 = ""
        End If

        DecodeHistory = True
    Exit Function
DecodeHistoryErr:
    Call RaiseError(MyUnhandledError, "cWU:DecodeHistory Method")
End Function

'LINUX compatible
'**********************************************************
'*      Прочитать заголовок текущего рабочего блока       *
'*  Mode=1: файла типа Work_unit.sah (клиент для Windows) *
'*  Mode=2: файла типа work_uni.sah (клиент для Linux)    *
'**********************************************************
Public Function ReadFile(ByVal Mode As Long) As String
Dim sfile As String
Dim i As Long
    On Error GoTo ReadFileErr
    
    If (Dir(GetFilePath(Mode), vbNormal) <> "") Then
        sfile = ""
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFilePath(Mode) For Input As hFile
        ''For i = 1 To 2000
            ''sfile = sfile + Input(1, #1)
        ''Next i
        'OK, try another tactics...  :)
        sfile = Input(2000, #1)
        Close hFile
        ReadFile = sfile
    Else
        ReadFile = ""
    End If
    Exit Function
ReadFileErr:

    Call RaiseError(MyUnhandledError, "cWU:ReadFile Method")
End Function

'**********************************************************
'*          Добавляет запись в базу данных для            *
'*          последующей демонстрации в DataGrid           *
'*   Mode = 0 инициализация списка при запуске программы  *
'*          Mode = 1 добавление только 1 записи           *
'*        Mode = 2 обновление существующей записи         *
'**********************************************************
Public Sub AddRecord(ByVal Mode As Long, Optional ByVal RecNum As Long)
Dim sfile As String
Dim i As Long
'ЭКВИВАЛЕНТНЫЕ ОПЕРАТОРЫ
'rsWUhistory!ID = ...
'rsWUhistory.Fields.Item("ID") = ...
On Error GoTo AddRecordErr

Select Case Mode
Case 0:
    If (Dir(GetFilePath(0), vbNormal) <> "") Then
        sfile = ""
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFilePath(0) For Input As hFile
        For i = 1 To RegRecords
            Line Input #hFile, sfile
            DecodeHistory (sfile)
            If Not (rsWUhistory.EOF) Then
                rsWUhistory.MoveLast
            End If
            rsWUhistory.AddNew
            With rsWUhistory.Fields
                .Item("ID") = NumID
                .Item("Comments") = Comments
                .Item("Unit Name") = UnitName
                .Item("Recording time (GMT)") = ExtractTime(TimeOfRec)
                .Item("Base frequency") = SubbandBase
                .Item("Start RA") = StartRA
                .Item("Start DEC") = StartDEC
                .Item("End RA") = EndRA
                .Item("End DEC") = EndDEC
                .Item("Receiver") = Receiver
                .Item("Tape") = TapeVer
            End With
            rsWUhistory.Update
        Next i
        Close hFile
        rsWUhistory.MoveFirst ' Move to the beginning of the recordset.
''Open GetFilePath(0) For Append As hFile
''Debug.Print "FILE DEBUGGER REPORT"
''Debug.Print Seek(hFile)
''Close hFile
    End If
Case 1:
    'Если нужно добавить только 1 запись, то вызывать ReadHistory
    If DecodeHistory(ReadHistory(RecNum, 1)) Then
        rsWUhistory.MoveFirst
        Do While CLng(Val(rsWUhistory.Fields.Item("ID"))) < CLng(Val(NumID))
            rsWUhistory.MoveNext
        Loop
        rsWUhistory.AddNew
        With rsWUhistory.Fields
            .Item("ID") = NumID
            .Item("Comments") = Comments
            .Item("Unit Name") = UnitName
            .Item("Recording time (GMT)") = ExtractTime(TimeOfRec)
            .Item("Base frequency") = SubbandBase
            .Item("Start RA") = StartRA
            .Item("Start DEC") = StartDEC
            .Item("End RA") = EndRA
            .Item("End DEC") = EndDEC
            .Item("Receiver") = Receiver
            .Item("Tape") = TapeVer
        End With
        rsWUhistory.Update
        rsWUhistory.MoveFirst ' Move to the beginning of the recordset.
    End If
Case 2:
    If DecodeHistory(ReadHistory(RecNum, 1)) Then
        rsWUhistory.MoveFirst
        i = 0
        Do While CLng(Val(rsWUhistory.Fields.Item("ID"))) < CLng(Val(NumID))
            rsWUhistory.MoveNext
            i = i + 1
        Loop
        If CLng(Val(rsWUhistory.Fields.Item("ID"))) = CLng(Val(NumID)) Then
            ''rsWUhistory.Fields.Delete (i)
            ''rsWUhistory.AddNew
            With rsWUhistory.Fields
                .Item("ID") = NumID
                .Item("Comments") = Comments
                .Item("Unit Name") = UnitName
                .Item("Recording time (GMT)") = ExtractTime(TimeOfRec)
                .Item("Base frequency") = SubbandBase
                .Item("Start RA") = StartRA
                .Item("Start DEC") = StartDEC
                .Item("End RA") = EndRA
                .Item("End DEC") = EndDEC
                .Item("Receiver") = Receiver
                .Item("Tape") = TapeVer
            End With
            rsWUhistory.Update
            rsWUhistory.MoveFirst ' Move to the beginning of the recordset.
        End If
    End If
End Select
Exit Sub
AddRecordErr:

    Call RaiseError(MyUnhandledError, "cWU:AddRecord Method")
End Sub

'**********************************************************
'*         Дешифрует заголовок блока из строки            *
'*       sfile - строка, полученная от  ReadFile          *
'**********************************************************
Public Function DecodeWU(sfile As String) As Boolean
Dim i As Long
Dim stopper As String
    On Error GoTo DecodeWUErr

        stopper = Chr(10)
        NumID = "-1"
        Type_of_unit = GetToken("type=", sfile, stopper)
        Task = GetToken("task=", sfile, stopper)
        Version = GetToken("version=", sfile, stopper)
        UnitName = GetToken("name=", sfile, stopper)
        Data_type = GetToken("data_type=", sfile, stopper)
        Data_class = GetToken("data_class=", sfile, stopper)
        Splitter_version = GetToken("splitter_version=", sfile, stopper)
        StartRA = GetToken("start_ra=", sfile, stopper)
        StartDEC = GetToken("start_dec=", sfile, stopper)
        EndRA = GetToken("end_ra=", sfile, stopper)
        EndDEC = GetToken("end_dec=", sfile, stopper)
        AngleRange = GetToken("angle_range=", sfile, stopper)
        TimeOfRec = GetToken("time_recorded=", sfile, stopper)
        SubbandCenter = GetToken("subband_center=", sfile, stopper)
        SubbandBase = GetToken("subband_base=", sfile, stopper)
        SubbandRate = GetToken("subband_sample_rate=", sfile, stopper)
        fft_len = GetToken("fft_len=", sfile, stopper)
        ifft_len = GetToken("ifft_len=", sfile, stopper)
        SubbandNum = GetToken("subband_number=", sfile, stopper)
        Receiver = GetToken("receiver=", sfile, stopper)
        Nsamples = GetToken("nsamples=", sfile, stopper)
        TapeVer = GetToken("tape_version=", sfile, stopper)
        NumPositions = GetToken("num_positions=", sfile, stopper)
        For i = 0 To 21
            Coords(i) = GetToken("coord" & CStr(i) & "=", sfile, stopper)
        Next i
        DecodeWU = True
    Exit Function
DecodeWUErr:

    Call RaiseError(MyUnhandledError, "cWU:DecodeWU Method")
End Function

'**********************************************************
'*        Зашифровывает заголовок блока в строку          *
'**********************************************************
Public Function EncodeWU() As String
Dim sfile As String
Dim i As Long
    On Error GoTo EncodeWUErr

    sfile = sfile + "$"                         'Division mark
    
    If NumID = "" Then
        NumID = "#"                             'Empty string
    End If
    sfile = sfile + NumID + "$"
    
    If Type_of_unit = "" Then
        Type_of_unit = "#"                      'Empty string
    End If
    sfile = sfile + Type_of_unit + "$"
    
    If Task = "" Then
        Task = "#"                              'Empty string
    End If
    sfile = sfile + Task + "$"
        
    If Version = "" Then
        Version = "#"                           'Empty string
    End If
    sfile = sfile + Version + "$"
    
    If UnitName = "" Then
        UnitName = "#"                          'Empty string
    End If
    sfile = sfile + UnitName + "$"
    
    If Data_type = "" Then
        Data_type = "#"                         'Empty string
    End If
    sfile = sfile + Data_type + "$"
    
    If Data_class = "" Then
        Data_class = "#"                        'Empty string
    End If
    sfile = sfile + Data_class + "$"
    
    If Splitter_version = "" Then
        Splitter_version = "#"                  'Empty string
    End If
    sfile = sfile + Splitter_version + "$"
    
    If StartRA = "" Then
        StartRA = "#"                           'Empty string
    End If
    sfile = sfile + StartRA + "$"
    
    If StartDEC = "" Then
        StartDEC = "#"                          'Empty string
    End If
    sfile = sfile + StartDEC + "$"
    
    If EndRA = "" Then
        EndRA = "#"                             'Empty string
    End If
    sfile = sfile + EndRA + "$"
    
    If EndDEC = "" Then
        EndDEC = "#"                            'Empty string
    End If
    sfile = sfile + EndDEC + "$"
    
    If AngleRange = "" Then
        AngleRange = "#"                        'Empty string
    End If
    sfile = sfile + AngleRange + "$"
    
    If TimeOfRec = "" Then
        TimeOfRec = "#"                         'Empty string
    End If
    sfile = sfile + TimeOfRec + "$"
    
    If SubbandCenter = "" Then
        SubbandCenter = "#"                     'Empty string
    End If
    sfile = sfile + SubbandCenter + "$"
    
    If SubbandBase = "" Then
        SubbandBase = "#"                       'Empty string
    End If
    sfile = sfile + SubbandBase + "$"
    
    If SubbandRate = "" Then
        SubbandRate = "#"                       'Empty string
    End If
    sfile = sfile + SubbandRate + "$"
    
    If fft_len = "" Then
        fft_len = "#"                          'Empty string
    End If
    sfile = sfile + fft_len + "$"
    
    If ifft_len = "" Then
        ifft_len = "#"                          'Empty string
    End If
    sfile = sfile + ifft_len + "$"
    
    If SubbandNum = "" Then
        SubbandNum = "#"                        'Empty string
    End If
    sfile = sfile + SubbandNum + "$"
    
    If Receiver = "" Then
        Receiver = "#"                          'Empty string
    End If
    sfile = sfile + Receiver + "$"
    
    If Nsamples = "" Then
        Nsamples = "#"                          'Empty string
    End If
    sfile = sfile + Nsamples + "$"
    
    If TapeVer = "" Then
        TapeVer = "#"                          'Empty string
    End If
    sfile = sfile + TapeVer + "$"
    
    If NumPositions = "" Then
        NumPositions = "#"                      'Empty string
    End If
    sfile = sfile + NumPositions + "$"
    
    If Comments = "" Then
        Comments = "#"                          'Empty string
    End If
    sfile = sfile + Comments + "$"
    
    For i = 0 To 21
        If Coords(i) = "" Then
            Coords(i) = "#"                     'Empty string
        End If
        sfile = sfile + Coords(i) + "$"
    Next i
    
    If Reserve1 = "" Then
        Reserve1 = "#"                          'Empty string
    End If
    sfile = sfile + Reserve1 + "$"
    
    If Reserve2 = "" Then
        Reserve2 = "#"                          'Empty string
    End If
    sfile = sfile + Reserve2 + "$"

    If Reserve3 = "" Then
        Reserve3 = "#"                          'Empty string
    End If
    sfile = sfile + Reserve3 + "$"
    
    EncodeWU = sfile
    
    Exit Function
EncodeWUErr:

    Call RaiseError(MyUnhandledError, "cWU:EncodeWU Method")
End Function

'**********************************************************
'*      Записывает зашифрованный заголовок в журнал       *
'*       sfile - зашифрованная строка (от EncodeWU)       *
'*Mode=0 - изменение существующей записи или вставка новой*
'*   Mode = 1 - добавление новой записи В КОНЕЦ файла     *
'**********************************************************
Public Function WriteHistory(ByVal sfile As String, ByVal Mode As Long) As Boolean
Dim s As String, Num As String
Dim hTMPFile As Long, i As Long
Dim WriteDone As Boolean
'Проверить файл на наличие одноИМЕННОГО блока и блока с тем же ID - в ввызывающей функции!!!
    On Error GoTo WriteHistoryErr

    hFile = FreeFile
    WriteHistory = False
    
    If (Dir(GetFilePath(0), vbNormal) <> "") Then   'Если журнал существует, то
        If FileLen(GetFilePath(0)) < 50 Then        'Если он пуст (т.к. запись гораздо длиннее 50 байтов)
            Open GetFilePath(0) For Output As hFile
            Print #hFile, sfile
            Close (hFile)
            RegRecords = RegRecords + 1
            LastRecordNum = GetLastNum
            WriteHistory = True
            RaiseEvent WriteComplete        'Yield to application about completed task
        Else    'Если файл не пуст, то нужно аккуратно добавить или перезаписать
            ''ТЕСТИРОВАНИЕ НЕ ВЫЯВИЛО ОШИБОК ПРИ ОТСУТСТВИИ СЛЕДУЮЩИХ ИНСТРУКЦИЙ
            ''Open GetFilePath(0) For Input As hFile
            ''Do
            ''    Line Input #hFile, s
            ''    NumOfRecords = NumOfRecords + 1
            ''Loop While Not (EOF(hFile))
            '''Seek hFile, 1
            ''Close (hFile)
            Select Case Mode
                Case 0:     'Модификация существующей записи
                    WriteDone = False
                    Num = NumID 'Запомнить ID блока пока он не перезаписался другим значением
                    hFile = FreeFile
                    Open GetFilePath(0) For Input As hFile
                    hTMPFile = FreeFile
                    Open App.path & "\tmpdata.tmp" For Output As hTMPFile
                    For i = 1 To RegRecords
                        Line Input #hFile, s
                        If Not (DecodeHistory(s)) Then
                            'Слабое утешение: сбоя программы не избежать
                            'Произошло нарушение целостности файлового формата
                            WriteHistory = False
                            Exit For
                            Exit Function   'Это ерунда, но для успокоения совести...
                        End If
                        If WriteDone Then
                            'Новый блок уже добавлен, поэтому просто перекачиваем
                            'оставшиеся записи...
                            Print #hTMPFile, s
                        ElseIf CDbl(Val(Num)) > CDbl(Val(NumID)) Then
                            'Если ID нового блока больше, чем у только что прочитанного,
                            'то пока пропустить запись нового блока
                            Print #hTMPFile, s
                        ElseIf CDbl(Val(Num)) = CDbl(Val(NumID)) Then
                            'ID нового блока равен ID старого! Перезаписываем...
                            Print #hTMPFile, sfile
                            WriteDone = True
                            Debug.Print sfile
                        ElseIf Not (WriteDone) Then
                            'Осталась последняя возможность: ID нового блока меньше,
                            'чем у только что прочитанного, но запись еще не
                            'производилась, поэтому вставляем новый блок и
                            'записываем сразу за ним старый...
                            Print #hTMPFile, sfile
                            Print #hTMPFile, s
                            RegRecords = RegRecords + 1
                            WriteDone = True
                        End If
                    Next i
                    Close hFile
                    Close hTMPFile
                    WriteHistory = True
                    Kill GetFilePath(0)
                    Name App.path & "\tmpdata.tmp" As GetFilePath(0)
                    RaiseEvent WriteComplete        'Yield to application about completed task
                   'Kill App.Path & "\tmpdata.tmp"
                Case 1:     'Добавление новой записи
                    Open GetFilePath(0) For Append As hFile
                    Print #hFile, sfile
                    Close (hFile)
                    RegRecords = RegRecords + 1
                    LastRecordNum = GetLastNum
                    WriteHistory = True
                    RaiseEvent WriteComplete        'Yield to application about completed task
            End Select
        End If
    Else                                            'Файл не существует
        Open GetFilePath(0) For Output As hFile
        Print #hFile, sfile
        Close (hFile)
        WriteHistory = True
        RegRecords = RegRecords + 1
        LastRecordNum = GetLastNum
        RaiseEvent WriteComplete        'Yield to application about completed task
    End If
    
    Exit Function
WriteHistoryErr:

    Call RaiseError(MyUnhandledError, "cWU:WriteHistory Method")
End Function

'**********************************************************
'*       Восстанавливает журнал при нестандартном         *
'*             завершении предыдущей сессии               *
'*       LogState хранится в реестре и обновляется        *
'*       после каждой операции перезаписи журнала         *
'**********************************************************
Public Function CrushRecover(LogState As Long) As Boolean
    On Error GoTo CrushRecoverErr
'TO DO - константы должны соответствовать во всех частях проекта
    Select Case LogState
        Case 1:     'Началось обновление, но запись не успела начаться
        Case 2:     'Шел процесс создания временного файла
        Case 3:     'Временный файл был создан, но удалить старый или
                    'переименовать временный файл не удалось
    End Select
    
    Exit Function
CrushRecoverErr:

    Call RaiseError(MyUnhandledError, "cWU:CrushRecover Method")
End Function

'**********************************************************
'*    Проверяет соответствие числа зарегестрированных     *
'*    блоков в реестре действительному положению дел      *
'**********************************************************
Public Function CheckRegSettings(NumRec As Long, Repair As Boolean) As Boolean
Dim s As String
Dim NumOfRecords As Long
    On Error GoTo CheckRegSettingsErr

    hFile = FreeFile
    NumOfRecords = 0
    If (Dir(GetFilePath(0), vbNormal) <> "") Then   'Если журнал существует, то
        If FileLen(GetFilePath(0)) > 50 Then        'Если он не пуст
            Open GetFilePath(0) For Input As hFile
            Do
                Line Input #hFile, s
                NumOfRecords = NumOfRecords + 1
            Loop While Not (EOF(hFile))
            Close (hFile)
        End If
    End If
    If NumOfRecords = NumRec Then   'Число записей соответствует числу в реестре
        CheckRegSettings = True     'Исправлять не надо
    Else
        CheckRegSettings = False    'Обнаружено несоответствие
        If Repair Then              'Если нужно исправить, то...
            SaveSetting App.Title, "Settings", "NumOfHistoryRec", NumOfRecords
            RegRecords = NumOfRecords
        End If
    End If

    Exit Function
CheckRegSettingsErr:

    Call RaiseError(MyUnhandledError, "cWU:CheckRegSettings Method")
End Function

'**********************************************************
'*     Возвращает номер блока в журнале по его имени      *
'**********************************************************
Public Function GetIDbyName(name As String) As String
Dim sfile As String
Dim fResult As String
    On Error GoTo GetIDbyNameErr
    
    GetIDbyName = ""
    fResult = ""
    hFile = FreeFile
    
    If (Dir(GetFilePath(0), vbNormal) <> "") Then   'Если журнал существует, то
        If FileLen(GetFilePath(0)) > 50 Then        'Если он не пуст
            Open GetFilePath(0) For Input As hFile
                Do
                    Line Input #hFile, sfile
                    If GetEncodedToken("Name", sfile) = name Then
                        fResult = GetEncodedToken("NumID", sfile)
                    End If
                Loop While Not (EOF(hFile))
            Close (hFile)
        End If
    End If
    GetIDbyName = fResult
    
    Exit Function
GetIDbyNameErr:

    Call RaiseError(MyUnhandledError, "cWU:GetIDbyName Method")
End Function

'**********************************************************
'*Проверяет наличие блока с заданным номером и/или именем *
'*        Mode - определяет режим работы функции          *
'*               Mode=0 - проверка ID                     *
'*               Mode=1 - проверка Name                   *
'*               Mode=2 - проверка обоих параметров       *
'**********************************************************
Public Function CheckUnit(ByVal Mode As Long, ByVal IDorNAME As String, Optional ByVal NAMEonly As String) As Boolean
Dim s As String
    On Error GoTo CheckUnitErr
    
    CheckUnit = False
    hFile = FreeFile
    
    If (Dir(GetFilePath(0), vbNormal) <> "") Then   'Если журнал существует, то
        If FileLen(GetFilePath(0)) > 50 Then        'Если он не пуст
            Open GetFilePath(0) For Input As hFile
            Select Case Mode
                Case 0:
                    Do
                        Line Input #hFile, s
                        If GetEncodedToken("NumID", s) = IDorNAME Then
                            CheckUnit = True
                        End If
                    Loop While Not (EOF(hFile))
                Case 1:
                    Do
                        Line Input #hFile, s
                        If GetEncodedToken("Name", s) = IDorNAME Then
                            CheckUnit = True
                        End If
                    Loop While Not (EOF(hFile))
                Case 2:
                    Do
                        Line Input #hFile, s
                        If GetEncodedToken("NumID", s) = IDorNAME Then
                            If GetEncodedToken("Name", s) = NAMEonly Then
                                CheckUnit = True
                            End If
                        End If
                    Loop While Not (EOF(hFile))
                End Select
            Close (hFile)
        End If
    End If
    
    Exit Function
CheckUnitErr:

    Call RaiseError(MyUnhandledError, "cWU:CheckUnit Method")
End Function

'**********************************************************
'*           Определяет ID последней записи               *
'**********************************************************
Public Function GetLastNum() As Long
Dim s As String
Dim Current As Long
Dim MaxID As Long
    On Error GoTo GetLastNumErr

    hFile = FreeFile
    MaxID = 0
    If (Dir(GetFilePath(0), vbNormal) <> "") Then   'Если журнал существует, то
        If FileLen(GetFilePath(0)) > 50 Then        'Если он не пуст
            Open GetFilePath(0) For Input As hFile
            Do
                Line Input #hFile, s
                Current = GetEncodedToken("NumID", s)
                If Current > MaxID Then
                    MaxID = Current
                End If
            Loop While Not (EOF(hFile))
            Close (hFile)
        End If
    End If
    GetLastNum = MaxID
    Exit Function
GetLastNumErr:

    Call RaiseError(MyUnhandledError, "cWU:GetLastNum Method")
End Function

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    Set Data = rsWUhistory
End Sub

Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
    
    ' Add the names of the new datamember to the DataMember collection
    ' This allows other objects to see the available DataMembers
    DataMembers.Add "WUhistory"
    
    If (Dir(GetFilePath(1), vbNormal) <> "") Then
        existWU = True      'Режим чтения из WU-файла
    Else
        existWU = False
    End If

    Set rsWUhistory = New ADODB.Recordset ' Set the object variable.

    ' Create a recordset with two fields and open the recordset. The
    ' first record has an integer data type and the second is a string,
    ' with a maximum of 256 characters. The CursorType is set to
    ' OpenStatic--an updatable snapshot of a set of records. The
    ' LockType is set to LockOptimistic to allow updates to the
    ' recordset
    With rsWUhistory
        .Fields.Append "ID", adBSTR
        .Fields.Append "Comments", adBSTR, 255
        .Fields.Append "Unit Name", adBSTR, 255
        .Fields.Append "Recording time (GMT)", adBSTR, 255
        .Fields.Append "Base frequency", adBSTR, 255
        .Fields.Append "Start RA", adBSTR, 255
        .Fields.Append "Start DEC", adBSTR, 255
        .Fields.Append "End RA", adBSTR, 255
        .Fields.Append "End DEC", adBSTR, 255
        .Fields.Append "Receiver", adBSTR, 255
        .Fields.Append "Tape", adBSTR, 255
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    'LastRecordNum = GetLastNum
End Sub
'
''**********************************************************
''*         Выделяет заданный параметр из строки.          *
''* Функция работает с символом перевода строки в качестве *
''* разделителя параметров. Этот способ применяется только *
''*    при обработке файлов рабочих блоков. Для файлов     *
''*    результатов используется функция GetEncodedToken    *
''**********************************************************
'Private Function GetToken(tokenname As String, ByVal psfile As String) As String
'Dim res As String
'Dim i As Long, StartPos As Long, EndPos As Long
'    res = ""
'    i = InStr(1, psfile, tokenname, vbTextCompare)
'    If i <> 0 Then
'        StartPos = i + Len(tokenname)
'        Do While (Mid(psfile, StartPos, 1) = " ")
'            StartPos = StartPos + 1
'        Loop
'        EndPos = InStr(StartPos, psfile, Chr(10), vbTextCompare)
'        res = Mid(psfile, StartPos, EndPos - StartPos)
'    End If
'    'Trim spaces
'    Do While (Left(res, 1) Like " ")
'        res = Right(res, Len(res) - 1)
'    Loop
'    Do While (Right(res, 1) Like " ")
'        res = Left(res, Len(res) - 1)
'    Loop
'    GetToken = res
'End Function

Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


#If DebugMode Then
    Public Property Get ClassDebugID()
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
