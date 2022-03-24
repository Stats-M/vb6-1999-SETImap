VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGauss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1

#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Type tGaussRec
    ncfft As Long
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
End Type

Private mvartime As Double 'local copy
Private mvarfreq As Double 'local copy
Private mvarfft_len As Long 'local copy
Private mvarra As Single 'local copy
Private mvardec As Single 'local copy
Private mvarpeak As Single 'local copy
Private mvarmean As Single 'local copy
Private mvarchirp_rate As Single 'local copy
Private mvarsigma As Single 'local copy
Private mvarchisqr As Single 'local copy
Private mvarncfft As Long 'local copy
Private mvarhFile As Long 'local copy

GaussRec As tGaussRec   'Запись для работы с файлом данных (SETIgaus.txt)

Public Property Let ncfft(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ncfft = 5
    mvarncfft = vData
End Property

Public Property Get ncfft() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ncfft
    ncfft = mvarncfft
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

Public Property Let chisqr(ByVal vData As Single)
    On Error GoTo chisqrLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chisqr = 5
    mvarchisqr = vData
    Exit Property


chisqrLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:chisqr Property Let")
End Property

Public Property Get chisqr() As Single
    On Error GoTo chisqrGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chisqr
    chisqr = mvarchisqr
    Exit Property


chisqrGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:chisqr Property Get")
End Property

Public Property Let sigma(ByVal vData As Single)
    On Error GoTo sigmaLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sigma = 5
    mvarsigma = vData
    Exit Property


sigmaLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:sigma Property Let")
End Property

Public Property Get sigma() As Single
    On Error GoTo sigmaGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sigma
    sigma = mvarsigma
    Exit Property


sigmaGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:sigma Property Get")
End Property

Public Property Let chirp_rate(ByVal vData As Single)
    On Error GoTo chirp_rateLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chirp_rate = 5
    mvarchirp_rate = vData
    Exit Property


chirp_rateLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:chirp_rate Property Let")
End Property

Public Property Get chirp_rate() As Single
    On Error GoTo chirp_rateGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chirp_rate
    chirp_rate = mvarchirp_rate
    Exit Property


chirp_rateGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:chirp_rate Property Get")
End Property

Public Property Let mean(ByVal vData As Single)
    On Error GoTo meanLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mean = 5
    mvarmean = vData
    Exit Property


meanLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:mean Property Let")
End Property

Public Property Get mean() As Single
    On Error GoTo meanGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mean
    mean = mvarmean
    Exit Property


meanGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:mean Property Get")
End Property

Public Property Let peak(ByVal vData As Single)
    On Error GoTo peakLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.peak = 5
    mvarpeak = vData
    Exit Property


peakLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:peak Property Let")
End Property

Public Property Get peak() As Single
    On Error GoTo peakGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.peak
    peak = mvarpeak
    Exit Property


peakGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:peak Property Get")
End Property

Public Property Let dec(ByVal vData As Single)
    On Error GoTo decLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dec = 5
    mvardec = vData
    Exit Property


decLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:dec Property Let")
End Property

Public Property Get dec() As Single
    On Error GoTo decGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dec
    dec = mvardec
    Exit Property


decGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:dec Property Get")
End Property

Public Property Let ra(ByVal vData As Single)
    On Error GoTo raLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ra = 5
    mvarra = vData
    Exit Property


raLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:ra Property Let")
End Property

Public Property Get ra() As Single
    On Error GoTo raGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ra
    ra = mvarra
    Exit Property


raGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:ra Property Get")
End Property

Public Property Let fft_len(ByVal vData As Long)
    On Error GoTo fft_lenLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fft_len = 5
    mvarfft_len = vData
    Exit Property


fft_lenLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:fft_len Property Let")
End Property

Public Property Get fft_len() As Long
    On Error GoTo fft_lenGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fft_len
    fft_len = mvarfft_len
    Exit Property


fft_lenGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:fft_len Property Get")
End Property

Public Property Let freq(ByVal vData As Double)
    On Error GoTo freqLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.freq = 5
    mvarfreq = vData
    Exit Property


freqLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:freq Property Let")
End Property

Public Property Get freq() As Double
    On Error GoTo freqGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.freq
    freq = mvarfreq
    Exit Property


freqGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:freq Property Get")
End Property

Public Property Let time(ByVal vData As Double)
    On Error GoTo timeLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.time = 5
    mvartime = vData
    Exit Property


timeLetErr:


    Call RaiseError(MyUnhandledError, "cGauss:time Property Let")
End Property

Public Property Get time() As Double
    On Error GoTo timeGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.time
    time = mvartime
    Exit Property


timeGetErr:


    Call RaiseError(MyUnhandledError, "cGauss:time Property Get")
End Property

Public Function CheckRecord(Mode As Long, ID As Long) As Boolean
End Function

Public Function GetFileName(Mode As Long, Optional path As String) As String
End Function

Public Function GetLastNum(path As String) As Long
End Function

Public Function Decode(sfile As String) As Boolean
End Function

Public Function WriteFile(path As String) As Boolean
End Function

Public Function ReadDataFile(Mode As Long, Optional ID As Long) As Boolean
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

