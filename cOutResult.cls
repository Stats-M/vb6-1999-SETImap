VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class used cOutfile and cResfile to gather info"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Private Type tGaussRec
    WUID As Long    'Номер блока
    'ID As Long      'Номер записи для одного и того же блока
    Label As Long     'Идентификация источника информации
    peak As Single
    mean As Single
    ra As Single
    dec As Single
    time As Double
    freq As Double
    sigma As Single
    chisqr As Single
    fft_len As Long
    chirp_rate As Single
    maxpow As Single
    Reserve1 As Long
    Reserve2 As Long
End Type

Private Type tSpikeRec
    WUID As Long    'Номер блока
    'ID As Long      'Номер записи для одного и того же блока
    Label As Long     'Идентификация источника информации
    power As Single
    ra As Single
    dec As Single
    time As Double
    freq As Double
    fft_len As Long
    chirp_rate As Single
    Reserve1 As Long
    Reserve2 As Long
End Type

Private Type tPulseRec
    WUID As Long    'Номер блока
    'ID As Long      'Номер записи для одного и того же блока
    Label As Long     'Идентификация источника информации
    power As Single
    mean As Single
    period As Single
    ra As Single
    dec As Single
    time As Double
    freq As Double
    fft_len As Long
    chirp_rate As Single
    snr As Single
    thresh As Single
    len_prof As Long
    Reserve1 As Long
    Reserve2 As Long
End Type

Private Type tTripletRec
    WUID As Long    'Номер блока
    'ID As Long      'Номер записи для одного и того же блока
    Label As Long     'Идентификация источника информации
    power As Single
    mean As Single
    period As Single
    ra As Single
    dec As Single
    time As Double
    freq As Double
    fft_len As Long
    chirp_rate As Single
    Reserve1 As Long
    Reserve2 As Long
End Type

Private Const HistoryMode = 2   'Режим работы с журналами

Private mvartime As Double 'local copy
Private mvarfreq As Double 'local copy
Private mvarfft_len As Long 'local copy
Private mvarra As Single 'local copy
Private mvardec As Single 'local copy
Private mvarchirp_rate As Single 'local copy

Private mvarpeak As Single 'local copy  -> for Gaussian data types only
Private mvarmean As Single 'local copy  -> for Gaussian data types only
Private mvarsigma As Single 'local copy  -> for Gaussian data types only
Private mvarchisqr As Single 'local copy  -> for Gaussian data types only
Private mvarmaxpow As Single 'local copy  -> for Gaussian data types only

Private mvarpower As Single 'local copy -> for Spikes data types only

Private mvarperiod As Single 'local copy -> for Pulses and Triplets data types
Private mvarsnr As Single 'local copy
Private mvarthresh As Single 'local copy
Private mvarlen_prof As Long 'local copy

Private mvarhGFile As Long 'local copy -> Gaussians file handler
Private mvarhSFile As Long 'local copy -> Spikes file handler
Private mvarhPFile As Long 'local copy -> Pulses file handler
Private mvarhTFile As Long 'local copy -> Triplets file handler
Private mvarhFile As Long 'local copy  -> Universal file handler

Private VerifyMode As Boolean   'Флаг осуществления проверки при записи,
                                'показывает, что при записи нужно будет просмотреть содержимое файла данных
Private GaussRec As tGaussRec   'Запись для работы с файлом данных (SETIgaus.dat)
Private SpikeRec As tSpikeRec   'Запись для работы с файлом данных (SETIspik.dat)
Private PulseRec As tPulseRec   'Запись для работы с файлом данных (SETIpuls.dat)
Private TripletRec As tTripletRec 'Запись для работы с файлом данных (SETItrip.dat)

Public Property Let len_prof(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.len_prof = 5
    mvarlen_prof = vData
End Property

Public Property Get len_prof() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.len_prof
    len_prof = mvarlen_prof
End Property

Public Property Let thresh(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.thresh = 5
    mvarthresh = vData
End Property

Public Property Get thresh() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.thresh
    thresh = mvarthresh
End Property

Public Property Let snr(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.snr = 5
    mvarsnr = vData
End Property

Public Property Get snr() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.snr
    snr = mvarsnr
End Property

Public Property Let period(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.period = 5
    mvarperiod = vData
End Property

Public Property Get period() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.period
    period = mvarperiod
End Property

Public Property Let hTFile(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hTFile = 5
    mvarhTFile = vData
End Property

Public Property Get hTFile() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hTFile
    hTFile = mvarhTFile
End Property

Public Property Let hPFile(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hPFile = 5
    mvarhPFile = vData
End Property

Public Property Get hPFile() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hPFile
    hPFile = mvarhPFile
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

Public Property Let maxpow(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.maxpow = 5
    mvarmaxpow = vData
End Property

Public Property Get maxpow() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.maxpow
    maxpow = mvarmaxpow
End Property

Public Property Let hSFile(ByVal vData As Long)
    On Error GoTo hSFileLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hSFile = 5
    mvarhSFile = vData
    Exit Property
hSFileLetErr:
    Call RaiseError(MyUnhandledError, "cSpike:hSFile Property Let")
End Property

Public Property Get hSFile() As Long
    On Error GoTo hSFileGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hSFile
    hSFile = mvarhSFile
    Exit Property
hSFileGetErr:
    Call RaiseError(MyUnhandledError, "cSpike:hSFile Property Get")
End Property

Public Property Let hGFile(ByVal vData As Long)
    On Error GoTo hGFileLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hGFile = 5
    mvarhGFile = vData
    Exit Property
hGFileLetErr:
    Call RaiseError(MyUnhandledError, "cSpike:hGFile Property Let")
End Property

Public Property Get hGFile() As Long
    On Error GoTo hGFileGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hGFile
    hGFile = mvarhGFile
    Exit Property
hGFileGetErr:
    Call RaiseError(MyUnhandledError, "cSpike:hGFile Property Get")
End Property

Public Property Let power(ByVal vData As Single)
    On Error GoTo powerLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.power = 5
    mvarpower = vData
    Exit Property
powerLetErr:
    Call RaiseError(MyUnhandledError, "cSpike:power Property Let")
End Property

Public Property Get power() As Single
    On Error GoTo powerGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.power
    power = mvarpower
    Exit Property
powerGetErr:
    Call RaiseError(MyUnhandledError, "cSpike:power Property Get")
End Property

Public Property Let chisqr(ByVal vData As Single)
    On Error GoTo chisqrLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chisqr = 5
    mvarchisqr = vData
    Exit Property
chisqrLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:chisqr Property Let")
End Property

Public Property Get chisqr() As Single
    On Error GoTo chisqrGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chisqr
    chisqr = mvarchisqr
    Exit Property
chisqrGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:chisqr Property Get")
End Property

Public Property Let sigma(ByVal vData As Single)
    On Error GoTo sigmaLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sigma = 5
    mvarsigma = vData
    Exit Property
sigmaLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:sigma Property Let")
End Property

Public Property Get sigma() As Single
    On Error GoTo sigmaGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sigma
    sigma = mvarsigma
    Exit Property
sigmaGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:sigma Property Get")
End Property

Public Property Let chirp_rate(ByVal vData As Single)
    On Error GoTo chirp_rateLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chirp_rate = 5
    mvarchirp_rate = vData
    Exit Property
chirp_rateLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:chirp_rate Property Let")
End Property

Public Property Get chirp_rate() As Single
    On Error GoTo chirp_rateGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chirp_rate
    chirp_rate = mvarchirp_rate
    Exit Property
chirp_rateGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:chirp_rate Property Get")
End Property

Public Property Let mean(ByVal vData As Single)
    On Error GoTo meanLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mean = 5
    mvarmean = vData
    Exit Property
meanLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:mean Property Let")
End Property

Public Property Get mean() As Single
    On Error GoTo meanGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mean
    mean = mvarmean
    Exit Property
meanGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:mean Property Get")
End Property

Public Property Let peak(ByVal vData As Single)
    On Error GoTo peakLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.peak = 5
    mvarpeak = vData
    Exit Property
peakLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:peak Property Let")
End Property

Public Property Get peak() As Single
    On Error GoTo peakGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.peak
    peak = mvarpeak
    Exit Property
peakGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:peak Property Get")
End Property

Public Property Let dec(ByVal vData As Single)
    On Error GoTo decLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dec = 5
    mvardec = vData
    Exit Property
decLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:dec Property Let")
End Property

Public Property Get dec() As Single
    On Error GoTo decGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dec
    dec = mvardec
    Exit Property
decGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:dec Property Get")
End Property

Public Property Let ra(ByVal vData As Single)
    On Error GoTo raLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ra = 5
    mvarra = vData
    Exit Property
raLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:ra Property Let")
End Property

Public Property Get ra() As Single
    On Error GoTo raGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ra
    ra = mvarra
    Exit Property
raGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:ra Property Get")
End Property

Public Property Let fft_len(ByVal vData As Long)
    On Error GoTo fft_lenLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fft_len = 5
    mvarfft_len = vData
    Exit Property
fft_lenLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:fft_len Property Let")
End Property

Public Property Get fft_len() As Long
    On Error GoTo fft_lenGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fft_len
    fft_len = mvarfft_len
    Exit Property
fft_lenGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:fft_len Property Get")
End Property

Public Property Let freq(ByVal vData As Double)
    On Error GoTo freqLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.freq = 5
    mvarfreq = vData
    Exit Property
freqLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:freq Property Let")
End Property

Public Property Get freq() As Double
    On Error GoTo freqGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.freq
    freq = mvarfreq
    Exit Property
freqGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:freq Property Get")
End Property

Public Property Let time(ByVal vData As Double)
    On Error GoTo timeLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.time = 5
    mvartime = vData
    Exit Property
timeLetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:time Property Let")
End Property

Public Property Get time() As Double
    On Error GoTo timeGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.time
    time = mvartime
    Exit Property
timeGetErr:
    Call RaiseError(MyUnhandledError, "cOutResult:time Property Get")
End Property
'END PROPERTIES LET/GET SECTION
'******************************

'LINUX compatible
'**********************************************************
'*           Проверить существование файла                *
'*  Mode=0: клиент для Windows                            *
'*  Mode=1: клиент для Linux                              *
'*  Mode=2: работа с журналом                             *
'* Target=0: Проверить существование файла Outfile.sah    *
'* Target=1: Проверить существование файла Result.sah     *
'* Target=2: Проверить существование файла SETIgaus.dat   *
'* Target=3: Проверить существование файла SETIspik.dat   *
'* Target=4: Проверить существование файла SETIpuls.dat   *
'* Target=5: Проверить существование файла SETItrip.dat   *
'**********************************************************
Public Function CheckFile(ByVal Mode As Long, ByVal Target As Long) As Boolean
Dim CheckResult As Boolean
    On Error GoTo CheckFileErr
    
    CheckResult = False
    If (Dir(GetFileName(Mode, Target), vbNormal) <> "") Then
        CheckResult = True
    End If
    CheckFile = CheckResult
    Exit Function
    
CheckFileErr:
    CheckFile = False
    Call RaiseError(MyUnhandledError, "cOutResult:CheckFile Method")
End Function

'**********************************************************
'*     Проверяет существование записи в журнале           *
'*    Mode=0: Проверка существования конкретной записи    *
'*            блока WUID                                  *
'*    Mode=1: Проверка определенного параметра записи     *
'*            для блока WUID (будущее использование)      *
'* Target=2: Получить расположение файла типа SETIgaus.dat*
'* Target=3: Получить расположение файла типа SETIspik.dat*
'* Target=4: Получить расположение файла типа SETIpuls.dat*
'* Target=5: Получить расположение файла типа SETItrip.dat*
'**********************************************************
Public Function CheckRecord(Mode As Long, Target As Long, Optional WUID As Long = -1, Optional ID As Long = -1) As Boolean
Dim LastRecNum As Long
Dim i As Long
    On Error GoTo CheckRecordErr

    CheckRecord = False
    If CheckFile(HistoryMode, Target) Then 'Файл существует - открываем его
'ВНИМАНИЕ! Здесь могут быть проблемы с указателями (файл уже может быть открыт)!!!
'TO DO
        Select Case Target
            Case 2:
                Dim TMPgaussrec As tGaussRec
                hFile = FreeFile
                Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(TMPgaussrec)
                LastRecNum = LOF(hFile) \ Len(TMPgaussrec)
                    If Not LastRecNum = 0 Then
                        Seek hFile, 1   'переместиться в начало
                        Select Case Mode    'Выбрать режим проверки
                            Case 0:
                                If Not WUID = -1 Then   'Задан ли номер блока?
                                    For i = 1 To LastRecNum     'Проверить все записи
                                        Get hFile, , TMPgaussrec    'Читать запись
                                        If TMPgaussrec.WUID = WUID Then 'Да, есть запись для этого блока
                                            CheckRecord = True
                                            i = LastRecNum      'Остановить проверку
                                        End If
                                    Next i
                                End If
                        End Select
                    End If
            Case 3:
                Dim TMPspikerec As tSpikeRec
                hFile = FreeFile
                Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(TMPspikerec)
                LastRecNum = LOF(hFile) \ Len(TMPspikerec)
                    If Not LastRecNum = 0 Then
                        Seek hFile, 1   'переместиться в начало
                        Select Case Mode    'Выбрать режим проверки
                            Case 0:
                                If Not WUID = -1 Then   'Задан ли номер блока?
                                    For i = 1 To LastRecNum     'Проверить все записи
                                        Get hFile, , TMPspikerec    'Читать запись
                                        If TMPspikerec.WUID = WUID Then 'Да, есть запись для этого блока
                                            CheckRecord = True
                                            i = LastRecNum      'Остановить проверку
                                        End If
                                    Next i
                                End If
                        End Select
                    End If
            Case 4:
                Dim TMPpulserec As tPulseRec
                hFile = FreeFile
                Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(TMPpulserec)
                LastRecNum = LOF(hFile) \ Len(TMPpulserec)
                    If Not LastRecNum = 0 Then
                        Seek hFile, 1   'переместиться в начало
                        Select Case Mode    'Выбрать режим проверки
                            Case 0:
                                If Not WUID = -1 Then   'Задан ли номер блока?
                                    For i = 1 To LastRecNum     'Проверить все записи
                                        Get hFile, , TMPpulserec    'Читать запись
                                        If TMPpulserec.WUID = WUID Then 'Да, есть запись для этого блока
                                            CheckRecord = True
                                            i = LastRecNum      'Остановить проверку
                                        End If
                                    Next i
                                End If
                        End Select
                    End If
            Case 5:
                Dim TMPtripletrec As tTripletRec
                hFile = FreeFile
                Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(TMPtripletrec)
                LastRecNum = LOF(hFile) \ Len(TMPtripletrec)
                    If Not LastRecNum = 0 Then
                        Seek hFile, 1   'переместиться в начало
                        Select Case Mode    'Выбрать режим проверки
                            Case 0:
                                If Not WUID = -1 Then   'Задан ли номер блока?
                                    For i = 1 To LastRecNum     'Проверить все записи
                                        Get hFile, , TMPtripletrec    'Читать запись
                                        If TMPtripletrec.WUID = WUID Then 'Да, есть запись для этого блока
                                            CheckRecord = True
                                            i = LastRecNum      'Остановить проверку
                                        End If
                                    Next i
                                End If
                        End Select
                    End If
        End Select
        Close hFile
    End If
        
    Exit Function
CheckRecordErr:
    Call RaiseError(MyObjectError1000, "cOutResult:CheckRecord Method")
End Function

'LINUX compatible
'**********************************************************
'*             Получить имя файла данных                  *
'*    Mode=0: Работа с клиентом для Windows               *
'*    Mode=1: Работа с клиентом для Linux                 *
'*    Mode=2: Работа с журналами                          *
'* Target=0: Получить расположение файла типа Outfile.sah *
'* Target=1: Получить расположение файла типа Result.sah  *
'* Target=2: Получить расположение файла типа SETIgaus.dat*
'* Target=3: Получить расположение файла типа SETIspik.dat*
'* Target=4: Получить расположение файла типа SETIpuls.dat*
'* Target=5: Получить расположение файла типа SETItrip.dat*
'* Имена файлов не фиксированы, что можно использовать    *
'* для обращения к сохраненным ранее файлам результатов,  *
'* задавая необязательный параметр path или изменив       *
'* глобальные константы стандартных имен файлов в Module1 *
'**********************************************************
Public Function GetFileName(ByVal Mode As Long, ByVal Target As Long, Optional path As String = "missing") As String
    On Error GoTo GetFileNameErr
Dim sfile As String     'Временная переменная для хранения результата
Dim Success As Boolean  'Флаг успешного выполнения операции

    sfile = ""
    Success = False
    'Для начала проверим наличие необязательного параметра
    If Not (path = "missing") Then
        'Параметр path задан - используем его!
        sfile = path
        'Стираем пробелы в конце и слэш (если есть)
        Do While (Right(sfile, 1) Like " ")
            sfile = Left(sfile, Len(sfile) - 1)
        Loop
        If Right(sfile, 1) = Slash Then
            sfile = Left(sfile, Len(sfile) - 1)
        End If
        If Right(sfile, 1) = BackSlash Then
            sfile = Left(sfile, Len(sfile) - 1)
        End If
        If (Dir(sfile, vbNormal) <> "") Then
            'Да, такой файл существует (его содержание не проверяется)
            Success = True
        End If
    End If
    'Если path не задан или в нем указана неверная информация, то...
    If Not Success Then
        'Будем использовать стандартные имена и расположение файлов
        Select Case Mode
            Case 0: 'Mode=0: Работа с клиентом для Windows
                sfile = SETIpath & BackSlash
            Case 1: 'Mode=1: Работа с клиентом для Linux
                sfile = LinuxPath & BackSlash
            Case 2: 'Mode=2: Работа с журналами
                sfile = App.path & BackSlash
        End Select
        Select Case Target
            Case 0: 'Mode=0: Получить расположение файла типа Outfile.txt
                sfile = sfile & FileOut
            Case 1: 'Mode=1: Получить расположение файла типа Result.txt
                sfile = sfile & FileRes
            Case 2: 'Mode=2: Получить расположение файла типа SETIgaus.txt
                sfile = sfile & GaussFile
            Case 3: 'Mode=3: Получить расположение файла типа SETIspik.txt
                sfile = sfile & SpikeFile
            Case 4: 'Mode=4: Получить расположение файла типа SETIpuls.txt
                sfile = sfile & PulseFile
            Case 5: 'Mode=5: Получить расположение файла типа SETItrip.txt
                sfile = sfile & TripletFile
        End Select
    End If
    GetFileName = sfile 'Выдать полученный результат вызывающей функции
    
    Exit Function
GetFileNameErr:
    Call RaiseError(MyUnhandledError, "cOutResult:GetFileName Method")
End Function

Public Function GetLastNum(Mode As Long, path As String) As Long
    On Error GoTo GetLastNumErr


    'your code goes here...


    Exit Function
GetLastNumErr:
    Call RaiseError(MyUnhandledError, "cOutResult:GetLastNum Method")
End Function

'LINUX compatible
'**********************************************************
'*        Определяет имя блока данных по записи в         *
'*                    файле Result.sah                    *
'**********************************************************
Public Function DetectWU(sfile As String) As String
Dim stopper As String
    On Error GoTo DetectWUErr
    
    'Определить разделитель
    stopper = "*"
    'Определить имя блока, для которого записан этот файл
    DetectWU = GetToken("name=", sfile, stopper)
    
    Exit Function
DetectWUErr:
    Call RaiseError(MyUnhandledError, "cOutResult:DetectWU Method")
End Function

'LINUX compatible
'**********************************************************
'*      Разделяет информацию, полученную от ReadFile      *
'*     и поочередно передает их на обработку в Decode     *
'*                     sfile - строка                     *
'*  Mode=0: клиент для Windows                            *
'*  Mode=1: клиент для Linux                              *
'*  ProcMode=0: не производить запись в журнал            *
'*  ProcMode=1: записывать результаты в журнал            *
'**********************************************************
Public Function Splitter(Mode As Long, sfile As String, ProcMode As Long, WUID As Long) As Boolean
Dim stopper As String
Dim psfile As String    'Часть строки, передаваемая декодеру
Dim StartPos As Long, EndPos As Long
Dim Seeker As Long      'Содержит позицию, с которой идет обработка строки
    On Error GoTo SplitterErr
    
    Splitter = True
    'Разделитель параметров - звездочка
    stopper = "*"
    'Заголовок не нужен - отрезаем его, чтобы не гонять туда-сюда лишние строки
    StartPos = InStr(1, sfile, "end_seti_header", vbTextCompare)  'Найти положение "end_seti_header" в строке
    If StartPos <> 0 Then
        StartPos = StartPos - 1  'Шаг назад, чтобы не отрезать букву "e" в слове "end_seti_header"
        sfile = Right(sfile, Len(sfile) - StartPos)
    End If
    If Not ProcMode Then    'Подготовка к записи
    'Организовать проверку индекса: если данные по этому блоку уже есть,
    'то организовать проверку каждой новой записи, иначе простой цикл записи
        If CheckRecord(0, 2, WUID) Then
            VerifyMode = True    'Нужна проверка при записи
        ElseIf CheckRecord(0, 3, WUID) Then
            VerifyMode = True    'Нужна проверка при записи
        End If
    End If
    StartPos = 1
    'Начинаем разделение строки
    Seeker = InStr(StartPos, sfile, "spike:", vbTextCompare)    'Инициализация
    Do While Seeker <> 0
        'Получить одну spike-запись
        psfile = Right(sfile, Len(sfile) - Seeker + 1) 'Отрезать все, что ДО текущей позиции
        psfile = GetToken("spike:", psfile, stopper) & " "   'Закрывающий пробел
        StartPos = Seeker + Len(psfile)
        Seeker = InStr(StartPos, sfile, "spike:", vbTextCompare)
        If Not (Decode(Mode, 3, psfile, WUID)) Then
            'Ошибка декодера - выходим из функции
            StartPos = Len(sfile)
            Splitter = False
        Else
            If ProcMode Then
                If Not (WriteHistory(SplitterOverwr, 3, WUID)) Then 'Первый параметр был 0
                    Splitter = False
                End If
            End If
        End If
    Loop
    If Splitter Then
        'Возвращаемся к началу
        StartPos = 1
        Seeker = InStr(StartPos, sfile, "gaussian:", vbTextCompare)
        Do While Seeker <> 0
            'Получить одну gaussian-запись
            psfile = Right(sfile, Len(sfile) - Seeker + 1) 'Отрезать все, что ДО текущей позиции
            psfile = GetToken("gaussian:", psfile, stopper) & " "    'Закрывающий пробел
            StartPos = Seeker + Len(psfile) 'учесть пробел
            Seeker = InStr(StartPos, sfile, "gaussian:", vbTextCompare)
            If Not (Decode(Mode, 2, psfile, WUID)) Then
                'Ошибка декодера - выходим из функции
                StartPos = Len(sfile)
                Splitter = False
            Else
                If ProcMode Then
                    If Not (WriteHistory(SplitterOverwr, 2, WUID)) Then 'Первый параметр был 0
                        Splitter = False
                    End If
                End If
            End If
        Loop
    End If
    If Splitter Then
        'Возвращаемся к началу
        StartPos = 1
        Seeker = InStr(StartPos, sfile, "pulse:", vbTextCompare)
        Do While Seeker <> 0
            'Получить одну pulse-запись
            psfile = Right(sfile, Len(sfile) - Seeker + 1) 'Отрезать все, что ДО текущей позиции
            psfile = GetToken("pulse:", psfile, stopper) & " "    'Закрывающий пробел
            StartPos = Seeker + Len(psfile) 'учесть пробел
            Seeker = InStr(StartPos, sfile, "pulse:", vbTextCompare)
            If Not (Decode(Mode, 4, psfile, WUID)) Then
                'Ошибка декодера - выходим из функции
                StartPos = Len(sfile)
                Splitter = False
            Else
                If ProcMode Then
                    If Not (WriteHistory(SplitterOverwr, 4, WUID)) Then 'Первый параметр был 0
                        Splitter = False
                    End If
                End If
            End If
        Loop
    End If
    If Splitter Then
        'Возвращаемся к началу
        StartPos = 1
        Seeker = InStr(StartPos, sfile, "triplet:", vbTextCompare)
        Do While Seeker <> 0
            'Получить одну triplet-запись
            psfile = Right(sfile, Len(sfile) - Seeker + 1) 'Отрезать все, что ДО текущей позиции
            psfile = GetToken("triplet:", psfile, stopper) & " "    'Закрывающий пробел
            StartPos = Seeker + Len(psfile) 'учесть пробел
            Seeker = InStr(StartPos, sfile, "triplet:", vbTextCompare)
            If Not (Decode(Mode, 5, psfile, WUID)) Then
                'Ошибка декодера - выходим из функции
                StartPos = Len(sfile)
                Splitter = False
            Else
                If ProcMode Then
                    If Not (WriteHistory(SplitterOverwr, 5, WUID)) Then 'Первый параметр был 0
                        Splitter = False
                    End If
                End If
            End If
        Loop
    End If
    
    Exit Function
SplitterErr:
    Call RaiseError(MyUnhandledError, "cOutResult:Splitter Method")
End Function

'LINUX compatible
'**********************************************************
'*   Дешифрует текстовую строку, полученную от Splitter   *
'*                     sfile - строка                     *
'*  Mode=0: клиент для Windows                            *
'*  Mode=1: клиент для Linux                              *
'*     Target=2: Для записи в файл SETIgaus.dat           *
'*     Target=3: Для записи в файл SETIspik.dat           *
'*     Target=4: Для записи в файл SETIpuls.dat           *
'*     Target=5: Для записи в файл SETItrip.dat           *
'**********************************************************
Public Function Decode(ByVal Mode As Long, ByVal Target As Long, sfile As String, WUID As Long) As Boolean
Dim stopper As String
    On Error GoTo DecodeErr
    
    'Разделитель параметров - пробел
    'stopper = Chr(10)
    stopper = "space"
    Select Case Target
        Case 2: 'Gauss-информация
            peak = CSng(Val(GetToken("peak=", sfile, stopper)))
            mean = CSng(Val(GetToken("mean=", sfile, stopper)))
            ra = CSng(Val(GetToken("ra=", sfile, stopper)))
            dec = CSng(Val(GetToken("dec=", sfile, stopper)))
            time = CDbl(Val(GetToken("time=", sfile, stopper)))
            freq = CDbl(Val(GetToken("freq=", sfile, stopper)))
            sigma = CSng(Val(GetToken("sigma=", sfile, stopper)))
            chisqr = CSng(Val(GetToken("chisqr=", sfile, stopper)))
            fft_len = CLng(Val(GetToken("fft_len=", sfile, stopper)))
            chirp_rate = CSng(Val(GetToken("chirprate=", sfile, stopper)))
            maxpow = CSng(Val(GetToken("maxpow=", sfile, stopper)))
            If EncodeHistory(Mode, Target, WUID) Then
                Decode = True
            Else
                Decode = False
            End If
        Case 3: 'Spike-информация
            power = CSng(Val(GetToken("power=", sfile, stopper)))
            ra = CSng(Val(GetToken("ra=", sfile, stopper)))
            dec = CSng(Val(GetToken("dec=", sfile, stopper)))
            time = CDbl(Val(GetToken("time=", sfile, stopper)))
            freq = CDbl(Val(GetToken("freq=", sfile, stopper)))
            fft_len = CLng(Val(GetToken("fft_len=", sfile, stopper)))
            chirp_rate = CSng(Val(GetToken("chirp_rate=", sfile, stopper)))
            If EncodeHistory(Mode, Target, WUID) Then
                Decode = True
            Else
                Decode = False
            End If
        Case 4: 'Pulse-информация
            power = CSng(Val(GetToken("power=", sfile, stopper)))
            mean = CSng(Val(GetToken("mean=", sfile, stopper)))
            period = CSng(Val(GetToken("period=", sfile, stopper)))
            ra = CSng(Val(GetToken("ra=", sfile, stopper)))
            dec = CSng(Val(GetToken("dec=", sfile, stopper)))
            time = CDbl(Val(GetToken("time=", sfile, stopper)))
            freq = CDbl(Val(GetToken("freq=", sfile, stopper)))
            fft_len = CLng(Val(GetToken("fft_len=", sfile, stopper)))
            chirp_rate = CSng(Val(GetToken("chirp_rate=", sfile, stopper)))
            snr = CSng(Val(GetToken("snr=", sfile, stopper)))
            thresh = CSng(Val(GetToken("thresh=", sfile, stopper)))
            len_prof = CLng(Val(GetToken("len_prof=", sfile, stopper)))
            If EncodeHistory(Mode, Target, WUID) Then
                Decode = True
            Else
                Decode = False
            End If
        Case 5: 'Triplet-информация
            power = CSng(Val(GetToken("power=", sfile, stopper)))
            mean = CSng(Val(GetToken("mean=", sfile, stopper)))
            period = CSng(Val(GetToken("period=", sfile, stopper)))
            ra = CSng(Val(GetToken("ra=", sfile, stopper)))
            dec = CSng(Val(GetToken("dec=", sfile, stopper)))
            time = CDbl(Val(GetToken("time=", sfile, stopper)))
            freq = CDbl(Val(GetToken("freq=", sfile, stopper)))
            fft_len = CLng(Val(GetToken("fft_len=", sfile, stopper)))
            chirp_rate = CSng(Val(GetToken("chirp_rate=", sfile, stopper)))
            If EncodeHistory(Mode, Target, WUID) Then
                Decode = True
            Else
                Decode = False
            End If
    End Select
    
    Exit Function
DecodeErr:
    Decode = False
    Call RaiseError(MyUnhandledError, "cOutResult:Decode Method")
End Function

'LINUX compatible
'**********************************************************
'*        Заполнить поля записи переменной журнала        *
'*  Mode=0: клиент для Windows                            *
'*  Mode=1: клиент для Linux                              *
'*     Target=2: Для записи в файл SETIgaus.dat           *
'*     Target=3: Для записи в файл SETIspik.dat           *
'*     Target=4: Для записи в файл SETIpuls.dat           *
'*     Target=5: Для записи в файл SETItrip.dat           *
'**********************************************************
Public Function EncodeHistory(ByVal Mode As Long, ByVal Target As Long, WUID As Long) As Boolean
    On Error GoTo EncodeHistoryErr

    EncodeHistory = False
    Select Case Target
        Case 2:     'Для записи в файл SETIgaus.dat
            With GaussRec
                Select Case Mode
                    Case 0:     'клиент для Windows
                        .Label = 1  'Идентификация источника информации
                    Case 1:     'клиент для Linux
                        .Label = 2  'Идентификация источника информации
                End Select
                .WUID = WUID
                '.ID = RecordID
                .peak = peak
                .mean = mean
                .ra = ra
                .dec = dec
                .time = time
                .freq = freq
                .sigma = sigma
                .chisqr = chisqr
                .fft_len = fft_len
                .chirp_rate = chirp_rate
                .maxpow = maxpow
                .Reserve1 = -123    'Признак отсутствия информации
                .Reserve2 = -123
            End With
            EncodeHistory = True
        Case 3:     'Для записи в файл SETIspik.dat
            With SpikeRec
                Select Case Mode
                    Case 0:     'клиент для Windows
                        .Label = 1  'Идентификация источника информации
                    Case 1:     'клиент для Linux
                        .Label = 2  'Идентификация источника информации
                End Select
                .WUID = WUID
                '.ID = RecordID
                .power = power
                .ra = ra
                .dec = dec
                .time = time
                .freq = freq
                .fft_len = fft_len
                .chirp_rate = chirp_rate
                .Reserve1 = -123    'Признак отсутствия информации
                .Reserve2 = -123
            End With
            EncodeHistory = True
        Case 4:     'Для записи в файл SETIpuls.dat
            With PulseRec
                Select Case Mode
                    Case 0:     'клиент для Windows
                        .Label = 1  'Идентификация источника информации
                    Case 1:     'клиент для Linux
                        .Label = 2  'Идентификация источника информации
                End Select
                .WUID = WUID
                '.ID = RecordID
                .power = power
                .mean = mean
                .period = period
                .ra = ra
                .dec = dec
                .time = time
                .freq = freq
                .fft_len = fft_len
                .chirp_rate = chirp_rate
                .snr = snr
                .thresh = thresh
                .len_prof = len_prof
                .Reserve1 = -123    'Признак отсутствия информации
                .Reserve2 = -123
            End With
            EncodeHistory = True
        Case 5:     'Для записи в файл SETItrip.dat
            With TripletRec
                Select Case Mode
                    Case 0:     'клиент для Windows
                        .Label = 1  'Идентификация источника информации
                    Case 1:     'клиент для Linux
                        .Label = 2  'Идентификация источника информации
                End Select
                .WUID = WUID
                '.ID = RecordID
                .power = power
                .mean = mean
                .period = period
                .ra = ra
                .dec = dec
                .time = time
                .freq = freq
                .fft_len = fft_len
                .chirp_rate = chirp_rate
                .Reserve1 = -123    'Признак отсутствия информации
                .Reserve2 = -123
            End With
            EncodeHistory = True
   End Select
    
    Exit Function
EncodeHistoryErr:
    EncodeHistory = False
    Call RaiseError(MyUnhandledError, "cOutResult:EncodeHistory Method")
End Function

'**********************************************************
'*         Произвести запись в файл данных SETImap        *
'*    Mode=0: Не производить перезаписи совпадений        *
'*    Mode=1: Производить перезапись совпадений           *
'*    Target=2: Произвести запись в файл SETIgaus.dat     *
'*    Target=3: Произвести запись в файл SETIspik.dat     *
'*    Target=4: Произвести запись в файл SETIpuls.dat     *
'*    Target=5: Произвести запись в файл SETItrip.dat     *
'**********************************************************
'ВНИМАНИЕ!!! Если существует несколько записей на одной частоте, то в журнал
'записывается ПЕРВАЯ из них. Считается, что последняя будет записана в
'State.sah (как более сильный сигнал) и сохранена в журнале
Public Function WriteHistory(ByVal Mode As Long, ByVal Target As Long, WUID As Long) As Boolean
Dim LastRecNum As Long
Dim i As Long
Dim Match As Boolean   'Эта запись была найдена при проверке (уже существует)
    On Error GoTo WriteHistoryErr

    WriteHistory = False
    Select Case Target
        Case 2:
            hFile = FreeFile
            Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(GaussRec)
            Seek hFile, 1
            Match = False
            LastRecNum = LOF(hFile) \ Len(GaussRec)
            If ((Not (LastRecNum = 0)) Or (VerifyMode)) Then
                Dim TMPgaussrec As tGaussRec
                For i = 1 To LastRecNum     'Проверить все записи
                    Get hFile, , TMPgaussrec    'Читать запись
                    If TMPgaussrec.WUID = WUID Then 'Текущий блок - начать проверку
                        ''If CLng(Trunc(CDbl(TMPgaussrec.chirp_rate * 10000))) = CLng(Trunc(CDbl(GaussRec.chirp_rate * 10000))) Then
                        'Изменение критерия проверки
                        If CLng(Trunc(CDbl(TMPgaussrec.peak * 10000))) = CLng(Trunc(CDbl(GaussRec.peak * 10000))) Then
                            Debug.Print "Gauss record has been founded in Gauss datafile."
                            Match = True
                            If Mode Then    'Разрешена перезапись
                                Seek hFile, i
                                Put hFile, , GaussRec
                                Debug.Print "Overwrite mode: OutResultHistory -> Gauss record has been overwritten"
                            End If
                            i = LastRecNum      'Остановить проверку, позиция совпала
                            WriteHistory = True
                        End If
                    End If
                Next i
            End If
            'Следующая ветка выполняется только в следующих случаях
            '   -Файл пуст
            '   -Не производится проверка существования записи
            '   -В процессе проверки запись не была найдена
            'Во всех этих случаях (Match=False) производится запись в конец файла
            If Not Match Then
                Seek hFile, LastRecNum + 1
                Put hFile, , GaussRec
                WriteHistory = True
                Debug.Print "Adding record to Gauss data file"
            End If
        Case 3:
            hFile = FreeFile
            Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(SpikeRec)
            Seek hFile, 1
            Match = False
            LastRecNum = LOF(hFile) \ Len(SpikeRec)
            If ((Not (LastRecNum = 0)) Or (VerifyMode)) Then
                Dim TMPspikerec As tSpikeRec
                For i = 1 To LastRecNum     'Проверить все записи
                    Get hFile, , TMPspikerec    'Читать запись
                    If TMPspikerec.WUID = WUID Then 'Текущий блок - начать проверку
                        ''If CLng(Trunc(CDbl(TMPspikerec.chirp_rate * 10000))) = CLng(Trunc(CDbl(SpikeRec.chirp_rate * 10000))) Then
                        'Изменение критерия проверки
                        If CLng(Trunc(CDbl(TMPspikerec.power * 10000))) = CLng(Trunc(CDbl(SpikeRec.power * 10000))) Then
                            Debug.Print "Spike record has been founded in Spike datafile."
                            Match = True
                            If Mode Then    'Разрешена перезапись
                                Seek hFile, i
                                Put hFile, , SpikeRec
                                Debug.Print "Overwrite mode: OutResultHistory -> Spike record has been overwritten"
                            End If
                            i = LastRecNum      'Остановить проверку, позиция совпала
                            WriteHistory = True
                        End If
                    End If
                Next i
            End If
            'Следующая ветка выполняется только в следующих случаях
            '   -Файл пуст
            '   -Не производится проверка существования записи
            '   -В процессе проверки запись не была найдена
            'Во всех этих случаях (Match=False) производится запись в конец файла
            If Not Match Then
                Seek hFile, LastRecNum + 1
                Put hFile, , SpikeRec
                WriteHistory = True
                Debug.Print "Adding record to Spike data file"
            End If
        Case 4:
            hFile = FreeFile
            Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(PulseRec)
            Seek hFile, 1
            Match = False
            LastRecNum = LOF(hFile) \ Len(PulseRec)
            If ((Not (LastRecNum = 0)) Or (VerifyMode)) Then
                Dim TMPpulserec As tPulseRec
                For i = 1 To LastRecNum     'Проверить все записи
                    Get hFile, , TMPpulserec    'Читать запись
                    If TMPpulserec.WUID = WUID Then 'Текущий блок - начать проверку
                        ''If CLng(Trunc(CDbl(TMPpulserec.chirp_rate * 10000))) = CLng(Trunc(CDbl(PulseRec.chirp_rate * 10000))) Then
                        'Изменение критерия проверки
                        If CLng(Trunc(CDbl(TMPpulserec.power * 10000))) = CLng(Trunc(CDbl(PulseRec.power * 10000))) Then
                            Debug.Print "Pulse record has been founded in Pulses datafile."
                            Match = True
                            If Mode Then    'Разрешена перезапись
                                Seek hFile, i
                                Put hFile, , PulseRec
                                Debug.Print "Overwrite mode: OutResultHistory -> Pulse record has been overwritten"
                            End If
                            i = LastRecNum      'Остановить проверку, позиция совпала
                            WriteHistory = True
                        End If
                    End If
                Next i
            End If
            'Следующая ветка выполняется только в следующих случаях
            '   -Файл пуст
            '   -Не производится проверка существования записи
            '   -В процессе проверки запись не была найдена
            'Во всех этих случаях (Match=False) производится запись в конец файла
            If Not Match Then
                Seek hFile, LastRecNum + 1
                Put hFile, , PulseRec
                WriteHistory = True
                Debug.Print "Adding record to Pulse data file"
            End If
        Case 5:
            hFile = FreeFile
            Open GetFileName(HistoryMode, Target) For Random As hFile Len = Len(TripletRec)
            Seek hFile, 1
            Match = False
            LastRecNum = LOF(hFile) \ Len(TripletRec)
            If ((Not (LastRecNum = 0)) Or (VerifyMode)) Then
                Dim TMPtripletrec As tTripletRec
                For i = 1 To LastRecNum     'Проверить все записи
                    Get hFile, , TMPtripletrec    'Читать запись
                    If TMPtripletrec.WUID = WUID Then 'Текущий блок - начать проверку
                        ''If CLng(Trunc(CDbl(TMPtripletrec.chirp_rate * 10000))) = CLng(Trunc(CDbl(TripletRec.chirp_rate * 10000))) Then
                        'Изменение критерия проверки
                        If CLng(Trunc(CDbl(TMPtripletrec.power * 10000))) = CLng(Trunc(CDbl(TripletRec.power * 10000))) Then
                            Debug.Print "Triplet record has been founded in Triplet datafile."
                            Match = True
                            If Mode Then    'Разрешена перезапись
                                Seek hFile, i
                                Put hFile, , TripletRec
                                Debug.Print "Overwrite mode: OutResultHistory -> Triplet record has been overwritten"
                            End If
                            i = LastRecNum      'Остановить проверку, позиция совпала
                            WriteHistory = True
                        End If
                    End If
                Next i
            End If
            'Следующая ветка выполняется только в следующих случаях
            '   -Файл пуст
            '   -Не производится проверка существования записи
            '   -В процессе проверки запись не была найдена
            'Во всех этих случаях (Match=False) производится запись в конец файла
            If Not Match Then
                Seek hFile, LastRecNum + 1
                Put hFile, , TripletRec
                WriteHistory = True
                Debug.Print "Adding record to Triplet data file"
            End If

    End Select
    Close hFile
    
    Exit Function
WriteHistoryErr:
    Call RaiseError(MyUnhandledError, "cOutResult:WriteHistory Method")
End Function

'**********************************************************
'*             Прочитать файл данных SETImap              *
'*           Target=2: Читать файл SETIgaus.dat           *
'*           Target=3: Читать файл SETIspik.dat           *
'*           Target=4: Читать файл SETIpuls.dat           *
'*           Target=5: Читать файл SETItrip.dat           *
'**********************************************************
Public Function ReadDataFile(Target As Long, Optional ID As Long = 0) As Boolean
    On Error GoTo ReadDataFileErr


    'your code goes here...


    Exit Function
ReadDataFileErr:
    Call RaiseError(MyUnhandledError, "cOutResult:ReadDataFile Method")
End Function

'LINUX compatible
'**********************************************************
'*                    Прочитать файл                      *
'*    Mode=0: Работа с клиентом для Windows               *
'*    Mode=1: Работа с клиентом для Linux                 *
'* Source=0: Получить расположение файла типа Outfile.sah *
'* Source=1: Получить расположение файла типа Result.sah  *
'**********************************************************
Public Function ReadFile(ByVal Mode As Long, ByVal Source As Long) As String
Attribute ReadFile.VB_Description = "This method read OUTFILE  (mode=0) or RESULT (mode=1) files. Name of file returned by GetFileName method"
Dim sfile As String
Dim nextchar As String
    On Error GoTo ReadFileErr

    sfile = ""
    If CheckFile(Mode, Source) Then   'OK, file exist
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFileName(Mode, Source) For Input As hFile
        Debug.Print "Success file open (OutResult.ReadFile Method)"
        Do Until EOF(hFile)
            nextchar = Input(1, #1)
            If nextchar = Chr(10) Or nextchar = Chr(13) Then
                nextchar = "*"  'Заменить перевод строки звездочкой
            End If
            sfile = sfile + nextchar
        Loop
        sfile = sfile & "*"     'Закрыть строку звездочкой
        Close hFile
        ReadFile = sfile
    End If
    Exit Function
    
ReadFileErr:
    Call RaiseError(MyUnhandledError, "cOutResult:ReadFile Method")
End Function

Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub

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
