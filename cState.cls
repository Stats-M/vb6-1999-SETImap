VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class working with current progress info file"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Private Type tInterCH   'Запись для сохранения результатов State.sah, перезаписываемых
    sValue1 As Single   'при появлении лучших сигналов (работает в пределах одного
    sValue2 As Single   'блока, затем обнуляется). Обнуление происходит автоматически
    sValue3 As Single   'при обработке Result.sah в OutResult (стираются только те
    sValue4 As Single   'записи, которые имеют отношение к обработанному блоку)
    sValue5 As Single
    sValue6 As Single
    sValue7 As Single
    sValue8 As Single
    lValue1 As Long
    lValue2 As Long
    lValue3 As Long
    Status As Long      'Определяет тип сигнала (пики, импульсы и т.д.)
    WUID As Long
End Type

'Private Type tStateRecEx
'    chirp_rate As Single
'    fftl As Long
'    s_power As Single
'    s_rate As Single
'    s_score As Single
'    s_fft_len As Long
'    s_bin As Long
'    g_power As Single
'    g_rate As Single
'    g_score As Single
'    g_fft_len As Long
'    g_sigma As Single
'    g_chisq As Single
'    g_bin As Single
'    g_true_mean As Single
'    Status As Long '=0 если запись считается пустой (даже если поля данных заполнены)
'End Type

Private Type tStateRec
    chirp_rate As Single
    fftl As Long
    s_power As Single
    s_rate As Single
    s_score As Single
    s_fft_len As Long
    s_bin As Long
    g_power As Single
    g_rate As Single
    g_score As Single
    g_fft_len As Long
    g_sigma As Single
    g_chisq As Single
    g_bin As Single
    g_true_mean As Single
    Status As Long '=0 если запись считается пустой (даже если поля данных заполнены)
    
    potfreq As Long
    potactivity As Long
    outfilepos As Long
    p_score As Single
    p_power As Single
    p_mean As Single
    p_period As Single
    p_freq_bin As Long
    p_time_bin As Long
    p_chirp_rate As Single
    p_fft_len As Long
    
    t_score As Single
    t_power As Single
    t_mean As Single
    t_period As Single
    t_bperiod As Single
    t_freq_bin As Long
    t_time_bin As Single
    t_chirp_rate As Single
    t_scale As Single
    t_fft_len As Long
End Type

'local variable(s) to hold property value(s)
Private mvarbs_power As Single 'local copy
Private mvarbs_rate As Single 'local copy
Private mvarbs_score As Single 'local copy
Private mvarbg_power As Single 'local copy
Private mvarbg_rate As Single 'local copy
Private mvarbg_score As Single 'local copy
Private mvarbs_fft_len As Long 'local copy
Private mvarbg_fft_len As Long 'local copy
Private mvarbg_sigma As Single 'local copy
Private mvarbs_bin As Long 'local copy
Private mvarbg_chisq As Single 'local copy
Private mvarbg_bin As Single 'local copy
Private mvarbg_true_mean As Single 'local copy
Private mvarcr As Single 'local copy
Private mvarfl As Long 'local copy
Private mvarStatus As Long 'local copy
Private mvarhFile As Long 'local copy
'local variable(s) to hold property value(s) (Ver 3.00 extensions)
Private mvarpotfreq As Long 'local copy
Private mvarpotactivity As Long 'local copy
Private mvaroutfilepos As Long 'local copy
Private mvarbp_score As Single 'local copy
Private mvarbp_power As Single 'local copy
Private mvarbp_mean As Single 'local copy
Private mvarbp_period As Single 'local copy
Private mvarbp_freq_bin As Long 'local copy
Private mvarbp_time_bin As Long 'local copy
Private mvarbp_chirp_rate As Single 'local copy
Private mvarbp_fft_len As Long 'local copy
Private mvarbt_score As Single 'local copy
Private mvarbt_power As Single 'local copy
Private mvarbt_mean As Single 'local copy
Private mvarbt_period As Single 'local copy
Private mvarbt_bperiod As Single 'local copy
Private mvarbt_freq_bin As Long 'local copy
Private mvarbt_time_bin As Single 'local copy
Private mvarbt_chirp_rate As Single 'local copy
Private mvarbt_scale As Single 'local copy
Private mvarbt_fft_len As Long 'local copy

'TO DO Events...
'Public Event PercentDone(ByVal Percent As Single, ByRef Cancel As Boolean)

Private StateRec As tStateRec   'Запись для работы с файлом данных (SETIstat.dat)
Private InterCH As tInterCH     'Запись для работы с кэш-файлом (SETIcach.dat)
Private WithEvents rsShistory As ADODB.Recordset
Attribute rsShistory.VB_VarHelpID = -1

Public Property Let bt_fft_len(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_fft_len = 5
    mvarbt_fft_len = vData
End Property

Public Property Get bt_fft_len() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_fft_len
    bt_fft_len = mvarbt_fft_len
End Property

Public Property Let bt_scale(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_scale = 5
    mvarbt_scale = vData
End Property

Public Property Get bt_scale() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_scale
    bt_scale = mvarbt_scale
End Property

Public Property Let bt_chirp_rate(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_chirp_rate = 5
    mvarbt_chirp_rate = vData
End Property

Public Property Get bt_chirp_rate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_chirp_rate
    bt_chirp_rate = mvarbt_chirp_rate
End Property

Public Property Let bt_time_bin(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_time_bin = 5
    mvarbt_time_bin = vData
End Property

Public Property Get bt_time_bin() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_time_bin
    bt_time_bin = mvarbt_time_bin
End Property

Public Property Let bt_freq_bin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_freq_bin = 5
    mvarbt_freq_bin = vData
End Property

Public Property Get bt_freq_bin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_freq_bin
    bt_freq_bin = mvarbt_freq_bin
End Property

Public Property Let bt_bperiod(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_bperiod = 5
    mvarbt_bperiod = vData
End Property

Public Property Get bt_bperiod() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_bperiod
    bt_bperiod = mvarbt_bperiod
End Property

Public Property Let bt_period(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_period = 5
    mvarbt_period = vData
End Property

Public Property Get bt_period() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_period
    bt_period = mvarbt_period
End Property

Public Property Let bt_mean(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_mean = 5
    mvarbt_mean = vData
End Property

Public Property Get bt_mean() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_mean
    bt_mean = mvarbt_mean
End Property

Public Property Let bt_power(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_power = 5
    mvarbt_power = vData
End Property

Public Property Get bt_power() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_power
    bt_power = mvarbt_power
End Property

Public Property Let bt_score(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bt_score = 5
    mvarbt_score = vData
End Property

Public Property Get bt_score() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bt_score
    bt_score = mvarbt_score
End Property

Public Property Let bp_fft_len(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_fft_len = 5
    mvarbp_fft_len = vData
End Property

Public Property Get bp_fft_len() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_fft_len
    bp_fft_len = mvarbp_fft_len
End Property

Public Property Let bp_chirp_rate(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_chirp_rate = 5
    mvarbp_chirp_rate = vData
End Property

Public Property Get bp_chirp_rate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_chirp_rate
    bp_chirp_rate = mvarbp_chirp_rate
End Property

Public Property Let bp_time_bin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_time_bin = 5
    mvarbp_time_bin = vData
End Property

Public Property Get bp_time_bin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_time_bin
    bp_time_bin = mvarbp_time_bin
End Property

Public Property Let bp_freq_bin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_freq_bin = 5
    mvarbp_freq_bin = vData
End Property

Public Property Get bp_freq_bin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_freq_bin
    bp_freq_bin = mvarbp_freq_bin
End Property

Public Property Let bp_period(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_period = 5
    mvarbp_period = vData
End Property

Public Property Get bp_period() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_period
    bp_period = mvarbp_period
End Property

Public Property Let bp_mean(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_mean = 5
    mvarbp_mean = vData
End Property

Public Property Get bp_mean() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_mean
    bp_mean = mvarbp_mean
End Property

Public Property Let bp_power(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_power = 5
    mvarbp_power = vData
End Property

Public Property Get bp_power() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_power
    bp_power = mvarbp_power
End Property

Public Property Let bp_score(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bp_score = 5
    mvarbp_score = vData
End Property

Public Property Get bp_score() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bp_score
    bp_score = mvarbp_score
End Property

Public Property Let outfilepos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.outfilepos = 5
    mvaroutfilepos = vData
End Property

Public Property Get outfilepos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.outfilepos
    outfilepos = mvaroutfilepos
End Property

Public Property Let potactivity(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.potactivity = 5
    mvarpotactivity = vData
End Property

Public Property Get potactivity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.potactivity
    potactivity = mvarpotactivity
End Property

Public Property Let potfreq(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.potfreq = 5
    mvarpotfreq = vData
End Property

Public Property Get potfreq() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.potfreq
    potfreq = mvarpotfreq
End Property

Public Property Let Status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property

Public Property Get Status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property

Public Property Let fl(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fl = 5
    mvarfl = vData
End Property

Public Property Get fl() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fl
    fl = mvarfl
End Property

Public Property Let cr(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cr = 5
    mvarcr = vData
End Property

Public Property Get cr() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cr
    cr = mvarcr
End Property

Public Property Let bg_true_mean(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_true_mean = 5
    mvarbg_true_mean = vData
End Property

Public Property Get bg_true_mean() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_true_mean
    bg_true_mean = mvarbg_true_mean
End Property

Public Property Let bg_bin(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_bin = 5
    mvarbg_bin = vData
End Property

Public Property Get bg_bin() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_bin
    bg_bin = mvarbg_bin
End Property

Public Property Let bg_chisq(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_chisq = 5
    mvarbg_chisq = vData
End Property

Public Property Get bg_chisq() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_chisq
    bg_chisq = mvarbg_chisq
End Property

Public Property Let bs_bin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bs_bin = 5
    mvarbs_bin = vData
End Property

Public Property Get bs_bin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bs_bin
    bs_bin = mvarbs_bin
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

Public Property Let bg_sigma(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_sigma = 5
    mvarbg_sigma = vData
End Property

Public Property Get bg_sigma() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_sigma
    bg_sigma = mvarbg_sigma
End Property

Public Property Let bg_fft_len(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_fft_len = 5
    mvarbg_fft_len = vData
End Property

Public Property Get bg_fft_len() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_fft_len
    bg_fft_len = mvarbg_fft_len
End Property

Public Property Let bs_fft_len(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bs_fft_len = 5
    mvarbs_fft_len = vData
End Property

Public Property Get bs_fft_len() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bs_fft_len
    bs_fft_len = mvarbs_fft_len
End Property

Public Property Let bg_score(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_score = 5
    mvarbg_score = vData
End Property

Public Property Get bg_score() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_score
    bg_score = mvarbg_score
End Property

Public Property Let bg_rate(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_rate = 5
    mvarbg_rate = vData
End Property

Public Property Get bg_rate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_rate
    bg_rate = mvarbg_rate
End Property

Public Property Let bg_power(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bg_power = 5
    mvarbg_power = vData
End Property

Public Property Get bg_power() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bg_power
    bg_power = mvarbg_power
End Property

Public Property Let bs_score(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bs_score = 5
    mvarbs_score = vData
End Property

Public Property Get bs_score() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bs_score
    bs_score = mvarbs_score
End Property

Public Property Let bs_rate(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bs_rate = 5
    mvarbs_rate = vData
End Property

Public Property Get bs_rate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bs_rate
    bs_rate = mvarbs_rate
End Property

Public Property Let bs_power(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bs_power = 5
    mvarbs_power = vData
End Property

Public Property Get bs_power() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bs_power
    bs_power = mvarbs_power
End Property
'END PROPERTIES SECTION
'**********************

'LINUX compatible
'**********************************************************
'*     Определяет, изменились ли параметры в State.sah    *
'*     с момента последнего запуска программы и, при      *
'*     необходимости, обновляет кэш и реестр              *
'*     При старте нового блока инициализация реестра      *
'*     проводится БЕЗ ПРОВЕРКИ в UpdateRegistry           *
'*  Mode=0: работа с клиентом для Windows                 *
'*  Mode=1: работа с клиентом для Linux                   *
'**********************************************************
Public Sub Interchange(ByVal Mode As Long)
Dim TMPvalue As Long
Dim NewResults As Boolean
    On Error GoTo InterchangeErr

    NewResults = False  'Флаг необходимости обновления реестра
    Select Case Mode
        Case 0: 'работа с клиентом для Windows
            TMPvalue = GetSetting(App.Title, "InterCH", "WinSpikeCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bs_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bs_power * 100) = TMPvalue) Then
                        'Изменение! -> Надо записать в файл...
                        EncodeCache Mode, 1
                        If FlushCache Then     'Обновить spike-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "WinGaussianCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bg_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bg_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 2
                        If FlushCache Then     'Обновить gaussian-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "WinPulseCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bp_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bp_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 3
                        If FlushCache Then     'Обновить pulse-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "WinTripletCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bt_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bt_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 4
                        If FlushCache Then     'Обновить triplet-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            If NewResults Then          'Пишем только если есть изменения
                UpdateRegistry (Mode)   'Зафиксировать факт обновления
            End If
        Case 1: 'работа с клиентом для Linux
            TMPvalue = GetSetting(App.Title, "InterCH", "LinSpikeCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bs_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bs_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 1
                        If FlushCache Then     'Обновить spike-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "LinGaussianCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bg_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bg_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 2
                        If FlushCache Then     'Обновить gaussian-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "LinPulseCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bp_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bp_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 3
                        If FlushCache Then     'Обновить pulse-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            TMPvalue = GetSetting(App.Title, "InterCH", "LinTripletCR", -1234)
            If Not (TMPvalue = -1234) Then
            'Инициализация реестра была проведена, можно двигаться дальше
            'При инициализации большинство значений нулевые (state.sah)
                If Not (bt_power = 0) Then
                'Значение НЕНУЛЕВОЕ, т.е. клиент нашел что-то интересное...
                    If Not (Trunc(bt_power * 100) = TMPvalue) Then
                        EncodeCache Mode, 4
                        If FlushCache Then     'Обновить triplet-запись в кэше
                            NewResults = True
                        End If
                    End If
                End If
            End If
            If NewResults Then          'Пишем только если есть изменения
                UpdateRegistry (Mode)   'Зафиксировать факт обновления
            End If
    End Select
Exit Sub

InterchangeErr:
    Call RaiseError(MyUnhandledError, "cState:Interchange Method")
End Sub

'LINUX compatible
'**********************************************************
'*     Записать данные в кэш-файл SETIcach.dat            *
'**********************************************************
Private Function FlushCache() As Boolean
Dim LastRecNum As Long
    On Error GoTo FlushCacheErr

    hFile = FreeFile
    Open GetFileName(5) For Random As hFile Len = Len(InterCH)
    'Файл открываем независимо от того, существует он или нет
    LastRecNum = LOF(hFile) \ Len(InterCH)  'Сколько записей всего
    Seek hFile, LastRecNum + 1              'Перевести головку в конец файла
    Put hFile, , InterCH                    'Записываем новую запись
    Close hFile
    FlushCache = True
Exit Function

FlushCacheErr:
    FlushCache = False
    Call RaiseError(MyUnhandledError, "cState:FlushCache Method")
End Function

'LINUX compatible
'**********************************************************
'*     Закодировать информацию для последующей записи     *
'*     в кэш-файл SETIcach.dat                            *
'*  Mode=0: работа с клиентом для Windows                 *
'*  Mode=1: работа с клиентом для Linux                   *
'*      Target=1: spike-запись                            *
'*      Target=2: gaussian-запись                         *
'*      Target=3: pulse-запись                            *
'*      Target=4: triplet-запись                          *
'**********************************************************
Private Sub EncodeCache(ByVal Mode As Long, ByVal Target As Long)
    On Error GoTo EncodeCacheErr

    ClearCacheRec
    Select Case Target
        Case 1:
            With InterCH
                .sValue1 = bs_power
                .sValue2 = bs_score
                .sValue3 = bs_rate
                .lValue1 = bs_bin
                .lValue2 = bs_fft_len
                .Status = 1
            End With
        Case 2:
            With InterCH
                .sValue1 = bg_power
                .sValue2 = bg_score
                .sValue3 = bg_rate
                .sValue4 = bg_sigma
                .sValue5 = bg_chisq
                .sValue6 = bg_bin
                .sValue7 = bg_true_mean
                .lValue1 = bg_fft_len
                .Status = 2
            End With
        Case 3:
            With InterCH
                .sValue1 = bp_power
                .sValue2 = bp_score
                .sValue3 = bp_mean
                .sValue4 = bp_period
                .sValue5 = bp_chirp_rate
                .lValue1 = bp_freq_bin
                .lValue2 = bp_time_bin
                .lValue3 = bp_fft_len
                .Status = 3
            End With
        Case 4:
            With InterCH
                .sValue1 = bt_power
                .sValue2 = bt_score
                .sValue3 = bt_mean
                .sValue4 = bt_period
                .sValue5 = bt_bperiod
                .sValue6 = bt_time_bin
                .sValue7 = bt_chirp_rate
                .sValue8 = bt_scale
                .lValue1 = bt_freq_bin
                .lValue2 = bt_fft_len
                .Status = 4
            End With
    End Select
    Select Case Mode
        Case 0: 'работа с клиентом для Windows
            InterCH.WUID = WinID
        Case 1: 'работа с клиентом для Linux
            InterCH.WUID = LinID
            InterCH.Status = InterCH.Status + 4
    End Select
    
Exit Sub

EncodeCacheErr:
    Call RaiseError(MyUnhandledError, "cState:EncodeCache Method")
End Sub

'LINUX compatible
'**********************************************************
'*     Записать данные в кэш-файл SETIcach.dat            *
'**********************************************************
Private Sub ClearCacheRec()
Dim LastRecNum As Long
    On Error GoTo ClearCacheRecErr

    With InterCH
        .lValue1 = 0
        .lValue2 = 0
        .lValue3 = 0
        .Status = 0
        .sValue1 = 0
        .sValue2 = 0
        .sValue3 = 0
        .sValue4 = 0
        .sValue5 = 0
        .sValue6 = 0
        .sValue7 = 0
        .sValue8 = 0
        .WUID = 0
    End With
Exit Sub

ClearCacheRecErr:
    Call RaiseError(MyUnhandledError, "cState:ClearCacheRec Method")
End Sub

'LINUX compatible
'**********************************************************
'*                 Прочитать файл                         *
'*  Mode=0: файл типа State.sah (клиент для Windows)      *
'*  Mode=1: файл типа State.sah (клиент для Linux)        *
'**********************************************************
Public Function ReadFile(ByVal Mode As Long) As String
Dim sfile As String
Dim nextchar As String
    On Error GoTo ReadFileErr

    sfile = ""
    If CheckFile(Mode) Then    'OK, file exist
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFileName(Mode) For Input As hFile
        Debug.Print "Success file open (State.ReadFile Method)"
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
    Call RaiseError(MyUnhandledError, "cState:ReadFile Method")
End Function

'LINUX compatible
'**********************************************************
'*           Проверить существование файла                *
'*  Mode=0: файла типа State.sah (клиент для Windows)     *
'*  Mode=1: файла типа State.sah (клиент для Linux)       *
'*  Mode=2: файла типа SETIstat.dat                       *
'*  Mode=3: файла типа SETItopg.dat                       *
'*  Mode=4: файла типа SETItops.dat                       *
'*  Mode=5: файла типа SETIcach.dat                       *
'**********************************************************
Public Function CheckFile(ByVal Mode As Long) As Boolean
Dim CheckResult As Boolean
    On Error GoTo CheckFileErr
    
    CheckResult = False
    If (Dir(GetFileName(Mode), vbNormal) <> "") Then
        CheckResult = True
    End If
    CheckFile = CheckResult
    Exit Function
    
CheckFileErr:
    CheckFile = False
    Call RaiseError(MyUnhandledError, "cState:CheckFile Method")
End Function

'LINUX compatible
'**********************************************************
'*             Получить имя файла данных                  *
'*  Mode=0: Получить расположение файла типа State.sah    *
'*          (клиент для Windows)                          *
'*  Mode=1: Получить расположение файла типа State.sah    *
'*          (клиент для Linux)                            *
'*  Mode=2: Получить расположение файла типа SETIstat.dat *
'*          (файл журнала)                                *
'*  Mode=3: Получить расположение файла типа SETItopg.dat *
'*          (файл-индекс лучших гауссиан)                 *
'*  Mode=4: Получить расположение файла типа SETItops.dat *
'*          (файл-индекс лучших пиковых сигналов)         *
'*  Mode=5: Получить расположение файла типа SETIcach.dat *
'*          (кэш-файл результатов из State.sah)           *
'* Имена файлов не фиксированы, что можно использовать    *
'* для обращения к сохраненным ранее файлам результатов,  *
'* задавая необязательный параметр path или изменив       *
'* глобальные константы стандартных имен файлов в Module1 *
'**********************************************************
Public Function GetFileName(ByVal Mode As Long, Optional ByVal path As String = "missing") As String
Dim sfile As String     'Временная переменная для хранения результата
Dim Success As Boolean  'Флаг успешного выполнения операции
    On Error GoTo GetFileNameErr

    sfile = ""
    Success = False
    'Для начала проверим наличие необязательного параметра
    If Not path = "missing" Then
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
        'Проверяем существование файла
        If (Dir(sfile, vbNormal) <> "") Then
            'Да, такой файл существует (его содержание не проверяется)
            Success = True
        End If
    End If
    'Если path не задан или в нем указана неверная информация, то...
    If Not Success Then
        'Будем использовать стандартные имена и расположение файлов
        Select Case Mode
            Case 0: 'Mode=0: Получить расположение файла типа State.sah (Windows)
                sfile = SETIpath & "\" & FileState
            Case 1: 'Mode=1: Получить расположение файла типа State.sah (Linux)
                sfile = LinuxPath & "\" & FileState
            Case 2: 'Mode=2: Получить расположение файла типа SETIstat.dat
                sfile = App.path & "\" & StateFile
            Case 3: 'Mode=3: Получить расположение файла типа SETItopg.dat
                sfile = App.path & "\" & IndexFileG
            Case 4: 'Mode=4: Получить расположение файла типа SETItops.dat
                sfile = App.path & "\" & IndexFileS
            Case 5: 'Mode=5: Получить расположение файла типа SETIcach.dat
                sfile = App.path & "\" & StateCache
        End Select
    End If
    GetFileName = sfile 'Выдать полученный результат вызывающей функции
    Exit Function
    
GetFileNameErr:
    Call RaiseError(MyUnhandledError, "cState:GetFileName Method")
End Function
'
''LINUX compatible
''**********************************************************
''*           Прочитать заданный параметр                  *
''* tokenname: Название параметра                          *
''*    psfile: Строка, в которой производится поиск        *
''**********************************************************
'Public Function GetToken(ByVal tokenname As String, ByVal psfile As String) As String
'Dim res As String
'Dim i As Long, StartPos As Long, EndPos As Long
'    On Error GoTo GetTokenErr
'
'    res = ""
'    i = InStr(1, psfile, tokenname, vbTextCompare)  'Найти положение параметра в строке
'    If i <> 0 Then
'        StartPos = i + Len(tokenname)   'Продвинуться вперед на длину названия параметра
'        EndPos = InStr(StartPos, psfile, "*", vbTextCompare)    'Найти закрывающую звездочку
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
'    Exit Function
'
'GetTokenErr:
'    Call RaiseError(MyUnhandledError, "cState:GetToken Method")
'End Function

'LINUX compatible
'**********************************************************
'*   Дешифрует заголовок блока из строки, полученной при  *
'*               чтении файла State.sah                   *
'*       sfile - строка, полученная от  ReadFile          *
'**********************************************************
Public Function DecodeState(sfile As String) As Boolean
Dim stopper As String
    On Error GoTo DecodeStateErr
            
        stopper = "*"
        bs_power = CSng(Val(GetToken("bs_power=", sfile, stopper)))
        bs_rate = CSng(Val(GetToken("bs_chirp_rate=", sfile, stopper)))
        bs_score = CSng(Val(GetToken("bs_score=", sfile, stopper)))
        bg_power = CSng(Val(GetToken("bg_power=", sfile, stopper)))
        bg_rate = CSng(Val(GetToken("bg_chirp_rate=", sfile, stopper)))
        bg_score = CSng(Val(GetToken("bg_score=", sfile, stopper)))
        bs_fft_len = CLng(Val(GetToken("bs_fft_len=", sfile, stopper)))
        bg_fft_len = CLng(Val(GetToken("bg_fft_len=", sfile, stopper)))
        bg_sigma = CSng(Val(GetToken("bg_sigma=", sfile, stopper)))
        bs_bin = CLng(Val(GetToken("bs_bin=", sfile, stopper)))
        bg_chisq = CSng(Val(GetToken("bg_chisq=", sfile, stopper)))
        bg_bin = CSng(Val(GetToken("bg_bin=", sfile, stopper)))
        bg_true_mean = CSng(Val(GetToken("bg_true_mean=", sfile, stopper)))
        cr = CSng(Val(GetToken("cr=", sfile, stopper)))
        fl = CLng(Val(GetToken("fl=", sfile, stopper)))
        
        potfreq = CLng(Val(GetToken("potfreq=", sfile, stopper)))
        potactivity = CLng(Val(GetToken("potactivity=", sfile, stopper)))
        outfilepos = CLng(Val(GetToken("outfilepos=", sfile, stopper)))
        bp_score = CSng(Val(GetToken("bp_score=", sfile, stopper)))
        bp_power = CSng(Val(GetToken("bp_power=", sfile, stopper)))
        bp_mean = CSng(Val(GetToken("bp_mean=", sfile, stopper)))
        bp_period = CSng(Val(GetToken("bp_period=", sfile, stopper)))
        bp_freq_bin = CLng(Val(GetToken("bp_freq_bin=", sfile, stopper)))
        bp_time_bin = CLng(Val(GetToken("bp_time_bin=", sfile, stopper)))
        bp_chirp_rate = CSng(Val(GetToken("bp_chirp_rate=", sfile, stopper)))
        bp_fft_len = CLng(Val(GetToken("bp_fft_len=", sfile, stopper)))
        
        bt_score = CSng(Val(GetToken("bt_score=", sfile, stopper)))
        bt_power = CSng(Val(GetToken("bt_power=", sfile, stopper)))
        bt_mean = CSng(Val(GetToken("bt_mean=", sfile, stopper)))
        bt_period = CSng(Val(GetToken("bt_period=", sfile, stopper)))
        bt_bperiod = CSng(Val(GetToken("bt_bperiod=", sfile, stopper)))
        bt_freq_bin = CLng(Val(GetToken("bt_freq_bin=", sfile, stopper)))
        bt_time_bin = CSng(Val(GetToken("bt_time_bin=", sfile, stopper)))
        bt_chirp_rate = CSng(Val(GetToken("bt_chirp_rate=", sfile, stopper)))
        bt_scale = CSng(Val(GetToken("bt_scale=", sfile, stopper)))
        bt_fft_len = CLng(Val(GetToken("bt_fft_len=", sfile, stopper)))
        
        Status = 1
        DecodeState = True
    Exit Function
    
DecodeStateErr:
    DecodeState = False
    Call RaiseError(MyUnhandledError, "cState:DecodeState Method")
End Function

'LINUX compatible
'**********************************************************
'*    Дешифрует запись из файла журнала SETIstat.dat      *
'*       StateRec содержит одну запись из журнала         *
'**********************************************************
Public Function DecodeHistory() As Boolean
    On Error GoTo DecodeHistoryErr
    
    bs_power = StateRec.s_power
    bs_rate = StateRec.s_rate
    bs_score = StateRec.s_score
    bg_power = StateRec.g_power
    bg_rate = StateRec.g_rate
    bg_score = StateRec.g_score
    bs_fft_len = StateRec.s_fft_len
    bg_fft_len = StateRec.g_fft_len
    bg_sigma = StateRec.g_sigma
    bs_bin = StateRec.s_bin
    bg_chisq = StateRec.g_chisq
    bg_bin = StateRec.g_bin
    bg_true_mean = StateRec.g_true_mean
    cr = StateRec.chirp_rate
    fl = StateRec.fftl
    Status = StateRec.Status
    
    potfreq = StateRec.potfreq
    potactivity = StateRec.potactivity
    outfilepos = StateRec.outfilepos
    bp_score = StateRec.p_score
    bp_power = StateRec.p_power
    bp_mean = StateRec.p_mean
    bp_period = StateRec.p_period
    bp_freq_bin = StateRec.p_freq_bin
    bp_time_bin = StateRec.p_time_bin
    bp_chirp_rate = StateRec.p_chirp_rate
    bp_fft_len = StateRec.p_fft_len
    bt_score = StateRec.t_score
    bt_power = StateRec.t_power
    bt_mean = StateRec.t_mean
    bt_period = StateRec.t_period
    bt_bperiod = StateRec.t_bperiod
    bt_freq_bin = StateRec.t_freq_bin
    bt_time_bin = StateRec.t_time_bin
    bt_chirp_rate = StateRec.t_chirp_rate
    bt_scale = StateRec.t_scale
    bt_fft_len = StateRec.t_fft_len

    DecodeHistory = True
    Exit Function
    
DecodeHistoryErr:
    DecodeHistory = False
    Call RaiseError(MyUnhandledError, "cState:DecodeHistory Method")
End Function

'Ver 3.00
'LINUX compatible
'**********************************************************
'*    Дешифрует запись из файла журнала SETIstat.dat      *
'*       StateRec содержит одну запись из журнала         *
'**********************************************************
'Public Function DecodeHistoryEx() As Boolean
'    On Error GoTo DecodeHistoryExErr
'
'    bs_power = StateRecEx.s_power
'    bs_rate = StateRecEx.s_rate
'    bs_score = StateRecEx.s_score
'    bg_power = StateRecEx.g_power
'    bg_rate = StateRecEx.g_rate
'    bg_score = StateRecEx.g_score
'    bs_fft_len = StateRecEx.s_fft_len
'    bg_fft_len = StateRecEx.g_fft_len
'    bg_sigma = StateRecEx.g_sigma
'    bs_bin = StateRecEx.s_bin
'    bg_chisq = StateRecEx.g_chisq
'    bg_bin = StateRecEx.g_bin
'    bg_true_mean = StateRecEx.g_true_mean
'    cr = StateRecEx.chirp_rate
'    fl = StateRecEx.fftl
'    Status = StateRecEx.Status
'
'    DecodeHistoryEx = True
'    Exit Function
'
'DecodeHistoryExErr:
'    DecodeHistoryEx = False
'    Call RaiseError(MyUnhandledError, "cState:DecodeHistoryEx Method")
'End Function

'LINUX compatible
'**********************************************************
'*  Зашифровывает запись для файла журнала SETIstat.dat   *
'*       StateRec содержит одну запись для журнала        *
'**********************************************************
Public Function EncodeHistory() As Boolean
    On Error GoTo EncodeHistoryErr
    
    ClearFields
    StateRec.s_power = bs_power
    StateRec.s_rate = bs_rate
    StateRec.s_score = bs_score
    StateRec.g_power = bg_power
    StateRec.g_rate = bg_rate
    StateRec.g_score = bg_score
    StateRec.s_fft_len = bs_fft_len
    StateRec.g_fft_len = bg_fft_len
    StateRec.g_sigma = bg_sigma
    StateRec.s_bin = bs_bin
    StateRec.g_chisq = bg_chisq
    StateRec.g_bin = bg_bin
    StateRec.g_true_mean = bg_true_mean
    StateRec.chirp_rate = cr
    StateRec.fftl = fl
    
    StateRec.potfreq = potfreq
    StateRec.potactivity = potactivity
    StateRec.outfilepos = outfilepos
    StateRec.p_score = bp_score
    StateRec.p_power = bp_power
    StateRec.p_mean = bp_mean
    StateRec.p_period = bp_period
    StateRec.p_freq_bin = bp_freq_bin
    StateRec.p_time_bin = bp_time_bin
    StateRec.p_chirp_rate = bp_chirp_rate
    StateRec.p_fft_len = bp_fft_len
    StateRec.t_score = bt_score
    StateRec.t_power = bt_power
    StateRec.t_mean = bt_mean
    StateRec.t_period = bt_period
    StateRec.t_bperiod = bt_bperiod
    StateRec.t_freq_bin = bt_freq_bin
    StateRec.t_time_bin = bt_time_bin
    StateRec.t_chirp_rate = bt_chirp_rate
    StateRec.t_scale = bt_scale
    StateRec.t_fft_len = bt_fft_len
    
    EncodeHistory = True
    Exit Function
    
EncodeHistoryErr:
    EncodeHistory = False
    Call RaiseError(MyUnhandledError, "cState:EncodeHistory Method")
End Function

'LINUX compatible
'**********************************************************
'*    Записывает запись в файла журнала SETIstat.dat      *
'*       StateRec содержит одну запись для журнала        *
'*       ID содержит индекс записи                        *
'*  Mode=0: добавить запись в конец файла                 *
'*  Mode=1: Добавить запись с указанным номером           *
'**********************************************************
Public Function WriteHistory(ByVal Mode As Long, Optional ByVal ID As Long = -6) As Boolean
Dim LastRecNum As Long
Dim WritePos As Long
Dim i As Long
    On Error GoTo WriteHistoryErr
    
    hFile = FreeFile
    Open GetFileName(2) For Random As hFile Len = Len(StateRec)
    LastRecNum = LOF(hFile) \ Len(StateRec)
    Select Case Mode
        Case 0: 'добавить в конец файла
            WritePos = LastRecNum + 1
            StateRec.Status = 1 'опустить флаг игнорирования информации
            Seek hFile, WritePos
            Put hFile, , StateRec
        Case 1:
            WritePos = LastRecNum + 1
            If Not (ID = -6) Then
            'Параметр ID задан - используем его!
                If (ID > WritePos) Then 'Заполнить недостающие записи в файле пустыми значениями
                'Эта ветка исполняется ТОЛЬКО тогда, когда ID по крайней мере
                'на 2 больше числа записей в файле
                    'поднять флаг игнорирования информации (status=0)
                    '(считать последующие записи ПУСТЫМИ)
                    StateRec.Status = 0
                    For i = WritePos To ID - 1  'заполнить все записи до ID
                        Seek hFile, i
                        Put hFile, , StateRec
                    Next i
                    StateRec.Status = 1 'опустить флаг игнорирования информации
                    Seek hFile, ID      'переместиться на запись ID
                    Put hFile, , StateRec   'записать информацию
                    Debug.Print "WriteHistory Method -> STATE LOG: Adding record number " & ID
                Else    'Просто перезаписать поверх существующей записи
                    StateRec.Status = 1
                    Seek hFile, ID
                    Put hFile, , StateRec
                    Debug.Print "WriteHistory Method -> STATE LOG: Record number " & ID & "of " & LastRecNum
                End If
            End If
    End Select
    Close hFile
Exit Function
    
WriteHistoryErr:
    WriteHistory = False
    Call RaiseError(MyUnhandledError, "cState:WriteHistory Method")
End Function

'LINUX compatible
'**********************************************************
'*      Читает запись из файла журнала SETIstat.dat       *
'*       StateRec содержит одну запись из журнала         *
'*  ID содержит индекс записи                             *
'*       Возвращаемый параметр                            *
'*  ReadHistory=0: успешное чтение                        *
'*  ReadHistory=1: неизвестная ошибка                     *
'*  ReadHistory=2: индекс вне диапазона                   *
'**********************************************************
Public Function ReadHistory(ByVal ID As Long) As Long
Dim LastRecNum As Long
    On Error GoTo ReadHistoryErr
    
    ReadHistory = 1
    If CheckFile(2) Then
        hFile = FreeFile
        Open GetFileName(2) For Random As hFile Len = Len(StateRec)
        LastRecNum = LOF(hFile) \ Len(StateRec)
        If LastRecNum < ID Then 'ID больше, чем есть записей в файле!
            ReadHistory = 2
            Call RaiseErrMsg(1206, StandartErrHeader)
        Else    'читать запись из файла
            Get hFile, ID, StateRec
            ReadHistory = 0
        End If
        Close hFile
    Else
        Call RaiseErrMsg(1205, StandartErrHeader)
    End If
    Exit Function
    
ReadHistoryErr:
    ReadHistory = False
    Call RaiseError(MyUnhandledError, "cState:ReadHistory Method")
End Function

'Ver 3.00
'LINUX compatible
'**********************************************************
'*      Читает запись из файла журнала SETIstat.dat       *
'*       StateRec содержит одну запись из журнала         *
'*  ID содержит индекс записи                             *
'*       Возвращаемый параметр                            *
'*  ReadHistory=0: успешное чтение                        *
'*  ReadHistory=1: неизвестная ошибка                     *
'*  ReadHistory=2: индекс вне диапазона                   *
'**********************************************************
'Public Function ReadHistoryEx(ByVal ID As Long) As Long
'Dim LastRecNum As Long
'Dim hFILE As Long
'    On Error GoTo ReadHistoryExErr
'
'    ReadHistoryEx = 1
'        hFILE = FreeFile
'        Open App.path & "\SETIex.dat" For Random As hFILE Len = Len(StateRecEx)
'        LastRecNum = LOF(hFILE) \ Len(StateRecEx)
'        If LastRecNum < ID Then 'ID больше, чем есть записей в файле!
'            ReadHistoryEx = 2
'            Call RaiseErrMsg(1206, StandartErrHeader)
'        Else    'читать запись из файла
'            Get hFILE, ID, StateRecEx
'            ReadHistoryEx = 0
'        End If
'        Close hFILE
'    Exit Function
'
'ReadHistoryExErr:
'    ReadHistoryEx = False
'    Call RaiseError(MyUnhandledError, "cState:ReadHistoryEx Method")
'End Function

'LINUX compatible
'**********************************************************
'*      Читает запись из файла индекса SETItopg.dat       *
'*              ID содержит индекс записи                 *
'*  Target=0: Работа с индексом пиков                     *
'*  Target=1: Работа с индексом гауссиан                  *
'**********************************************************
Public Function ReadIndex(ByVal Target As Long, ByVal ID As Long) As Boolean
Dim LastRecNum As Long
    On Error GoTo ReadIndexErr
    
    ReadIndex = False
    Select Case Target
    Case 0:
        If CheckFile(4) Then
            hFile = FreeFile
            Open GetFileName(4) For Random As hFile Len = Len(TopS)
            LastRecNum = LOF(hFile) \ Len(TopS)
            If LastRecNum < ID Then 'ID больше, чем есть записей в файле!
                ReadIndex = False
            Else    'читать запись из файла
                Get hFile, ID, TopS
                ReadIndex = True
            End If
            Close hFile
        End If
    Case 1:
        If CheckFile(3) Then
            hFile = FreeFile
            Open GetFileName(3) For Random As hFile Len = Len(TopG)
            LastRecNum = LOF(hFile) \ Len(TopG)
            If LastRecNum < ID Then 'ID больше, чем есть записей в файле!
                ReadIndex = False
            Else    'читать запись из файла
                Get hFile, ID, TopG
                ReadIndex = True
            End If
            Close hFile
        End If
    End Select
    Exit Function
    
ReadIndexErr:
    ReadIndex = False
    Call RaiseError(MyUnhandledError, "cState:ReadIndex Method")
End Function

'LINUX compatible
'**********************************************************
'*          Возвращает длину файла-индекса                *
'* GetLastRecNum = -1 -файл не найден                     *
'*  Target=2: Работа с файлом журнала                     *
'*  Target=3: Работа с индексом гауссиан                  *
'*  Target=4: Работа с индексом пиков                     *
'*  Target=5: Работа с кэшем                              *
'**********************************************************
Public Function GetLastRecNum(ByVal Target As Long) As Long
Dim LastRecNum As Long
    On Error GoTo GetLastRecNumErr
    
    Select Case Target
        Case 2:
            If CheckFile(2) Then
                hFile = FreeFile
                Open GetFileName(2) For Random As hFile Len = Len(StateRec)
                LastRecNum = LOF(hFile) \ Len(StateRec)
                GetLastRecNum = LastRecNum
                Close hFile
            Else
                GetLastRecNum = -1
            End If
        Case 3:
            If CheckFile(3) Then
                hFile = FreeFile
                Open GetFileName(3) For Random As hFile Len = Len(TopG)
                LastRecNum = LOF(hFile) \ Len(TopG)
                GetLastRecNum = LastRecNum
                Close hFile
            Else
                GetLastRecNum = -1
            End If
        Case 4:
            If CheckFile(4) Then
                hFile = FreeFile
                Open GetFileName(4) For Random As hFile Len = Len(TopS)
                LastRecNum = LOF(hFile) \ Len(TopS)
                GetLastRecNum = LastRecNum
                Close hFile
            Else
                GetLastRecNum = -1
            End If
        Case 5:
            If CheckFile(5) Then
                hFile = FreeFile
                Open GetFileName(5) For Random As hFile Len = Len(InterCH)
                LastRecNum = LOF(hFile) \ Len(InterCH)
                GetLastRecNum = LastRecNum
                Close hFile
            Else
                GetLastRecNum = -1
            End If
    End Select
    Exit Function
    
GetLastRecNumErr:
    Call RaiseError(MyUnhandledError, "cState:GetLastRecNum Method")
End Function

'LINUX compatible
'**********************************************************
'* Строит заново файл индексов SETItopg.dat,              *
'* содержащий отсортированные по интегральному показателю *
'* значения гауссиан. Строится по данным из State.sah     *
'*  Mode=0: Обновить только одну запись в индексе         *
'*          (после этого необходима сортировка)           *
'*          ID-номер записи в файле журнала (SETIstat.sah)*
'*  Mode=1: Построить индекс заново                       *
'*  Target=0: Работа с индексом пиков                     *
'*  Target=1: Работа с индексом гауссиан                  *
'**********************************************************
Public Function RebuildIndex(ByVal Mode As Long, ByVal Target As Long, Optional ByVal ID As Long = -6) As Boolean
Dim LastRecNum As Long
Dim BestPos As Long     'Текущий кандидат на перемещение (в сортировке), а также
                        'используется при поиске записи в индексе при Mode=0
Dim hFileInd As Long     'указатель на файл SETItopg.dat
Dim i As Long, j As Long
Dim MaxAverage As Single    'максимальное интегральное значение (для сортировки TopG)
                            'а также максимальная мощность (для TopS)
Dim MaxPos As Long
Dim TopG1 As tTopG, TopG2 As tTopG      'Временные переменные для сортировки
Dim TopS1 As tTopS, TopS2 As tTopS      'Временные переменные для сортировки
    On Error GoTo RebuildIndexErr
    
    RebuildIndex = False
    Select Case Mode
        Case 0: 'Добавить (если надо) в существующий файл
        ''АЛГОРИТМ
        ''Если ID задан - прочитать запись из журнала, иначе - ошибка, выход из функции
        ''Искать в индексе запись с таким же ID
        ''Если найдена, то Put на ее место новую запись
        ''Если нет, то добавить в конец индекса
        ''Сортировка индекса -> SUCCESS
            Select Case Target
                Case 0: 'Работа с индексом пиков
                    If Not (ID = -6) Then
                        If ReadHistory(ID) = 0 Then 'Чтение журнала успешно
                            If Not (StateRec.Status = 0) Then   'Флаг игнорирования информации опущен
                                'заполнить поля TopS
                                TopS.ID = ID
                                TopS.power = StateRec.s_power
                                TopS.rate = StateRec.s_rate
                                'Вспомогательная переменная
                                BestPos = -1
                                'Открыть файл индекса
                                hFileInd = FreeFile
                                Open GetFileName(4) For Random As hFileInd Len = Len(TopS)
                                LastRecNum = LOF(hFileInd) \ Len(TopS)
                                Seek hFileInd, 1
                                'Есть ли уже запись с таким ID в индексе?
                                For i = 1 To LastRecNum
                                    Get hFileInd, , TopS1
                                    If TopS1.ID = ID Then
                                        'Запомнить положение записи с заданным ID в файле индекса
                                        BestPos = Seek(hFileInd) - 1
                                        'Остановить цикл
                                        i = LastRecNum
                                    End If
                                Next i
                                If Not (BestPos = -1) Then
                                    'Запись с таким ID найдена в индексе
                                    Seek hFileInd, BestPos
                                    Put hFileInd, , TopS
                                Else
                                    'Это новая запись и в индексе ее еще нет: пишем в конец индекса
                                    Seek hFileInd, LastRecNum + 1
                                    Put hFileInd, , TopS
                                End If
                                Close hFileInd
                                RebuildIndex = True
                            End If
                        End If
                    Else
                        'Индекс не задан
                        Call RaiseErrMsg(1204, StandartErrHeader)
                    End If
                Case 1: 'Работа с индексом гауссиан
                    If Not (ID = -6) Then
                        If ReadHistory(ID) = 0 Then 'Чтение журнала успешно
                            If Not (StateRec.Status = 0) Then   'Флаг игнорирования информации опущен
                                'заполнить поля TopG
                                TopG.ID = ID
                                TopG.power = StateRec.g_power
                                TopG.rate = StateRec.g_rate
                                If Not (StateRec.g_power = 0) Then
                                    TopG.average = StateRec.g_power / StateRec.g_chisq
                                Else
                                    TopG.average = 0
                                End If
                                'Вспомогательная переменная
                                BestPos = -1
                                'Открыть файл индекса
                                hFileInd = FreeFile
                                Open GetFileName(3) For Random As hFileInd Len = Len(TopG)
                                LastRecNum = LOF(hFileInd) \ Len(TopG)
                                Seek hFileInd, 1
                                'Есть ли уже запись с таким ID в индексе?
                                For i = 1 To LastRecNum
                                    Get hFileInd, , TopG1
                                    If TopG1.ID = ID Then
                                        'Запомнить положение записи с заданным ID в файле индекса
                                        BestPos = Seek(hFileInd) - 1
                                        'Остановить цикл
                                        i = LastRecNum
                                    End If
                                Next i
                                If Not (BestPos = -1) Then
                                    'Запись с таким ID найдена в индексе
                                    Seek hFileInd, BestPos
                                    Put hFileInd, , TopG
                                Else
                                    'Это новая запись и в индексе ее еще нет: пишем в конец индекса
                                    Seek hFileInd, LastRecNum + 1
                                    Put hFileInd, , TopG
                                End If
                                Close hFileInd
                                RebuildIndex = True
                            End If
                        End If
                    Else
                        'Индекс не задан
                        Call RaiseErrMsg(1204, StandartErrHeader)
                    End If
            End Select
        Case 1: 'Построить индекс заново
            Select Case Target
                Case 0: 'Работа с индексом пиков
                    If CheckFile(4) Then
                        Kill GetFileName(4)
                    End If
                    'Открыть файл индекса
                    hFileInd = FreeFile
                    Open GetFileName(4) For Random As hFileInd Len = Len(TopS)
                    'LastRecordNum - переменная из реестра!
                    For i = 1 To LastRecordNum
                        If ReadHistory(i) = 0 Then
                            If Not (StateRec.Status = 0) Then
                                'Заполнить поля переменной
                                TopS.ID = i
                                TopS.power = StateRec.s_power
                                TopS.rate = StateRec.s_rate
                                'Запись в файл индекса
                                Put hFileInd, , TopS
                            End If
                        End If
                    Next i
                    Close hFileInd
                    RebuildIndex = True
                Case 1: 'Работа с индексом гауссиан
                    If CheckFile(3) Then
                        Kill GetFileName(3)
                    End If
                    'Открыть файл индекса
                    hFileInd = FreeFile
                    Open GetFileName(3) For Random As hFileInd Len = Len(TopG)
                    'LastRecordNum - переменная из реестра!
                    For i = 1 To LastRecordNum
                        If ReadHistory(i) = 0 Then
                            If Not (StateRec.Status = 0) Then
                                'Заполнить поля переменной
                                TopG.ID = i
                                TopG.power = StateRec.g_power
                                TopG.rate = StateRec.g_rate
                                If Not (StateRec.g_power = 0) Then
                                    TopG.average = StateRec.g_power / StateRec.g_chisq
                                Else
                                    TopG.average = 0
                                End If
                                'Запись в файл индекса
                                Put hFileInd, , TopG
                            End If
                        End If
                    Next i
                    Close hFileInd
                    RebuildIndex = True
            End Select
    End Select
    
    If RebuildIndex Then    'Предыдущий этап завершен успешно?
        Select Case Target
            Case 0: 'Работа с индексом пиков
                hFileInd = FreeFile
                Open GetFileName(4) For Random As hFileInd Len = Len(TopS)
                LastRecNum = LOF(hFileInd) \ Len(TopS)   'Сколько записей всего
                MaxAverage = 0
                For i = 1 To LastRecNum
                    BestPos = i         'Если не изменится, значит, переставлять не нужно
                    Seek hFileInd, i     'Переводим указатель
                    For j = i To LastRecNum
                        Get hFileInd, , TopS
                        If MaxAverage < TopS.power Then
                            MaxAverage = TopS.power
                            BestPos = Seek(hFileInd) - 1 'Запомнить положение лучшей записи
                        End If
                    Next j
                    If Not (BestPos = i) Then
                        Get hFileInd, i, TopS1           'Прочитать старую запись
                        Get hFileInd, BestPos, TopS2     'Прочитать новую запись
                        Seek hFileInd, i
                        Put hFileInd, , TopS2            'Записываем новую запись
                        Seek hFileInd, BestPos
                        Put hFileInd, , TopS1            'Записываем старую запись
                    End If
                    MaxAverage = 0  'Обнуляем, чтобы корректно продолжать сравнения
                Next i
                Close hFileInd
            Case 1: 'Работа с индексом гауссиан
                hFileInd = FreeFile
                Open GetFileName(3) For Random As hFileInd Len = Len(TopG)
                LastRecNum = LOF(hFileInd) \ Len(TopG)   'Сколько записей всего
                MaxAverage = 0
                For i = 1 To LastRecNum
                    BestPos = i         'Если не изменится, значит, переставлять не нужно
                    Seek hFileInd, i     'Переводим указатель
                    For j = i To LastRecNum
                        Get hFileInd, , TopG
                        If MaxAverage < TopG.average Then
                            MaxAverage = TopG.average
                            BestPos = Seek(hFileInd) - 1 'Запомнить положение лучшей записи
                        End If
                    Next j
                    If Not (BestPos = i) Then
                        Get hFileInd, i, TopG1           'Прочитать старую запись
                        Get hFileInd, BestPos, TopG2     'Прочитать новую запись
                        Seek hFileInd, i
                        Put hFileInd, , TopG2            'Записываем новую запись
                        Seek hFileInd, BestPos
                        Put hFileInd, , TopG1            'Записываем старую запись
                    End If
                    MaxAverage = 0  'Обнуляем, чтобы корректно продолжать сравнения
                Next i
                Close hFileInd
        End Select
    End If
    Exit Function
    
RebuildIndexErr:
    RebuildIndex = False
    Call RaiseError(MyUnhandledError, "cState:RebuildIndex Method")
End Function

'**********************************************************
'*            Очищает поля переменной StateRec            *
'**********************************************************
Private Sub ClearFields()
    With StateRec
        .chirp_rate = 0
        .fftl = 0
        .s_power = 0
        .s_rate = 0
        .s_score = 0
        .s_fft_len = 0
        .s_bin = 0
        .g_power = 0
        .g_rate = 0
        .g_score = 0
        .g_fft_len = 0
        .g_sigma = 0
        .g_chisq = 0
        .g_bin = 0
        .g_true_mean = 0
        .Status = 0
        
        .potfreq = 0
        .potactivity = 0
        .outfilepos = 0
        .p_score = 0
        .p_power = 0
        .p_mean = 0
        .p_period = 0
        .p_freq_bin = 0
        .p_time_bin = 0
        .p_chirp_rate = 0
        .p_fft_len = 0
        .t_score = 0
        .t_power = 0
        .t_mean = 0
        .t_period = 0
        .t_bperiod = 0
        .t_freq_bin = 0
        .t_time_bin = 0
        .t_chirp_rate = 0
        .t_scale = 0
        .t_fft_len = 0
    End With
End Sub

'**********************************************************
'* Обновляет данные о положении лучших сигналов в реестре *
'*   Mode = 0 работа с клиентом для Windows               *
'*   Mode = 1 работа с клиентом для Linux                 *
'*   Mode = 2 очистить реестр (клиент для Windows)        *
'*   Mode = 3 очистить реестр (клиентом для Linux)        *
'**********************************************************
Public Sub UpdateRegistry(ByVal Mode As Long)
Dim TMPvalue As Long
On Error GoTo UpdateRegistryErr
    
    Select Case Mode
        Case 0: 'работа с клинтом для Windows
            TMPvalue = Trunc(bs_power * 100)
            SaveSetting App.Title, "InterCH", "WinSpikeCR", TMPvalue
            TMPvalue = Trunc(bg_power * 100)
            SaveSetting App.Title, "InterCH", "WinGaussianCR", TMPvalue
            TMPvalue = Trunc(bp_power * 100)
            SaveSetting App.Title, "InterCH", "WinPulseCR", TMPvalue
            TMPvalue = Trunc(bt_power * 100)
            SaveSetting App.Title, "InterCH", "WinTripletCR", TMPvalue
        Case 1: 'работа с клинтом для Linux
            TMPvalue = Trunc(bs_power * 100)
            SaveSetting App.Title, "InterCH", "LinSpikeCR", TMPvalue
            TMPvalue = Trunc(bg_power * 100)
            SaveSetting App.Title, "InterCH", "LinGaussianCR", TMPvalue
            TMPvalue = Trunc(bp_power * 100)
            SaveSetting App.Title, "InterCH", "LinPulseCR", TMPvalue
            TMPvalue = Trunc(bt_power * 100)
            SaveSetting App.Title, "InterCH", "LinTripletCR", TMPvalue
        Case 2: 'очистить реестр (клиент для Windows)
            TMPvalue = -1234
            SaveSetting App.Title, "InterCH", "WinSpikeCR", TMPvalue
            SaveSetting App.Title, "InterCH", "WinGaussianCR", TMPvalue
            SaveSetting App.Title, "InterCH", "WinPulseCR", TMPvalue
            SaveSetting App.Title, "InterCH", "WinTripletCR", TMPvalue
        Case 3: 'очистить реестр (клиент для Linux)
            TMPvalue = -1234
            SaveSetting App.Title, "InterCH", "LinSpikeCR", TMPvalue
            SaveSetting App.Title, "InterCH", "LinGaussianCR", TMPvalue
            SaveSetting App.Title, "InterCH", "LinPulseCR", TMPvalue
            SaveSetting App.Title, "InterCH", "LinTripletCR", TMPvalue
    End Select
Exit Sub

UpdateRegistryErr:
    Call RaiseError(MyUnhandledError, "cState:UpdateRegistry Method")
End Sub


'**********************************************************
'*          Добавляет запись в базу данных для            *
'*          последующей демонстрации в DataGrid           *
'*   Mode = 0 инициализация списка при запуске программы  *
'*   Mode = 1 обновление существующей записи              *
'**********************************************************
Public Sub AddRecord(ByVal Mode As Long, Optional ByVal RecNum As Long)
Dim i As Long
'ЭКВИВАЛЕНТНЫЕ ОПЕРАТОРЫ
'rsWUhistory!ID = ...
'rsWUhistory.Fields.Item("ID") = ...
On Error GoTo AddRecordErr

Select Case Mode
Case 0:     'добавить все записи
    If CheckFile(2) Then                        'файл результатов существует
        For i = 1 To GetLastRecNum(2)         'просмотреть все записи
            If ReadHistory(i) = 0 Then          'успешное чтение записи
                ''If DecodeHistory Then          обойдемся без декодирования, нечего время терять
                    rsShistory.AddNew
                    With rsShistory.Fields
                        .Item("ID") = i
                        If StateRec.Status = 1 Then
                            .Item("Status") = "valid"
                        Else
                            .Item("Status") = "not valid"
                        End If
                        .Item("Spike power") = StateRec.s_power
                        .Item("Spike CR") = StateRec.s_rate
                        .Item("Gaussian power") = StateRec.g_power
                        .Item("fit") = StateRec.g_chisq
                        .Item("Gaussian CR") = StateRec.g_rate
                        .Item("Pulse power") = StateRec.p_power
                        .Item("Pulse score") = StateRec.p_score
                        .Item("Pulse CR") = StateRec.p_chirp_rate
                        .Item("Triplet power") = StateRec.t_power
                        .Item("Triplet score") = StateRec.t_score
                        .Item("Triplet CR") = StateRec.t_chirp_rate
                    End With
                    rsShistory.Update
                ''End If
            End If
        Next i
        rsShistory.MoveFirst
    End If
            
Case 1:
    If CheckFile(2) Then                        'файл результатов существует
        If RecNum <= GetLastRecNum(2) Then     'задано допустимое значение
            For i = 1 To RecNum - 1             'переместиться к нужной записи
                rsShistory.MoveNext
            Next i
            If ReadHistory(RecNum) = 0 Then          'успешное чтение записи
                With rsShistory.Fields
                    .Item("ID") = i
                    .Item("Status") = StateRec.Status
                    .Item("Spike power") = StateRec.s_power
                    .Item("Spike CR") = StateRec.s_rate
                    .Item("Gaussian power") = StateRec.g_power
                    .Item("fit") = StateRec.g_chisq
                    .Item("Gaussian CR") = StateRec.g_rate
                    .Item("Pulse power") = StateRec.p_power
                    .Item("Pulse score") = StateRec.p_score
                    .Item("Pulse CR") = StateRec.p_chirp_rate
                    .Item("Triplet power") = StateRec.t_power
                    .Item("Triplet score") = StateRec.t_score
                    .Item("Triplet CR") = StateRec.t_chirp_rate
                End With
                rsShistory.Update
            End If
        End If
        rsShistory.MoveFirst
    End If
End Select
Exit Sub
AddRecordErr:

    Call RaiseError(MyUnhandledError, "cState:AddRecord Method")
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    Set Data = rsShistory
End Sub

Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
        'Init variable with zero values
    #End If
    
    ClearFields
    
    ' Add the names of the new datamember to the DataMember collection
    ' This allows other objects to see the available DataMembers
    DataMembers.Add "Shistory"

    Set rsShistory = New ADODB.Recordset ' Set the object variable.

    ' Create a recordset with two fields and open the recordset. The
    ' first record has an integer data type and the second is a string,
    ' with a maximum of 256 characters. The CursorType is set to
    ' OpenStatic--an updatable snapshot of a set of records. The
    ' LockType is set to LockOptimistic to allow updates to the
    ' recordset
    With rsShistory
        .Fields.Append "ID", adBSTR
        .Fields.Append "Status", adBSTR
        .Fields.Append "Spike power", adBSTR
        .Fields.Append "Spike CR", adBSTR
        .Fields.Append "Gaussian power", adBSTR
        .Fields.Append "fit", adBSTR
        .Fields.Append "Gaussian CR", adBSTR
        .Fields.Append "Pulse power", adBSTR
        .Fields.Append "Pulse score", adBSTR
        .Fields.Append "Pulse CR", adBSTR
        .Fields.Append "Triplet power", adBSTR
        .Fields.Append "Triplet score", adBSTR
        .Fields.Append "Triplet CR", adBSTR
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

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
