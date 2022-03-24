VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpike"
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

Type tSpikeRec
    ncfft As Long
    power As Single
    ra As Single
    dec As Single
    time As Double
    freq As Double
    fft_len As Long
    chirp_rate As Single
End Type

'local variable(s) to hold property value(s)
Private mvartime As Double 'local copy
Private mvarfreq As Double 'local copy
Private mvarfft_len As Long 'local copy
Private mvarra As Single 'local copy
Private mvardec As Single 'local copy
Private mvarpower As Single 'local copy
Private mvarchirp_rate As Single 'local copy
Private mvarncfft As Long 'local copy
Private mvarhFile As Long 'local copy

SpikeRec As tSpikeRec   'Запись для работы с файлом данных (SETIspik.txt)

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

Public Property Let chirp_rate(ByVal vData As Single)
    On Error GoTo chirp_rateLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chirp_rate = 5
    mvarchirp_rate = vData
    Exit Property


chirp_rateLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:chirp_rate Property Let")
End Property

Public Property Get chirp_rate() As Single
    On Error GoTo chirp_rateGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chirp_rate
    chirp_rate = mvarchirp_rate
    Exit Property


chirp_rateGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:chirp_rate Property Get")
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

Public Property Let dec(ByVal vData As Single)
    On Error GoTo decLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dec = 5
    mvardec = vData
    Exit Property


decLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:dec Property Let")
End Property

Public Property Get dec() As Single
    On Error GoTo decGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dec
    dec = mvardec
    Exit Property


decGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:dec Property Get")
End Property

Public Property Let ra(ByVal vData As Single)
    On Error GoTo raLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ra = 5
    mvarra = vData
    Exit Property


raLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:ra Property Let")
End Property

Public Property Get ra() As Single
    On Error GoTo raGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ra
    ra = mvarra
    Exit Property


raGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:ra Property Get")
End Property

Public Property Let fft_len(ByVal vData As Long)
    On Error GoTo fft_lenLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fft_len = 5
    mvarfft_len = vData
    Exit Property


fft_lenLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:fft_len Property Let")
End Property

Public Property Get fft_len() As Long
    On Error GoTo fft_lenGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fft_len
    fft_len = mvarfft_len
    Exit Property


fft_lenGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:fft_len Property Get")
End Property

Public Property Let freq(ByVal vData As Double)
    On Error GoTo freqLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.freq = 5
    mvarfreq = vData
    Exit Property


freqLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:freq Property Let")
End Property

Public Property Get freq() As Double
    On Error GoTo freqGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.freq
    freq = mvarfreq
    Exit Property


freqGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:freq Property Get")
End Property

Public Property Let time(ByVal vData As Double)
    On Error GoTo timeLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.time = 5
    mvartime = vData
    Exit Property


timeLetErr:


    Call RaiseError(MyUnhandledError, "cSpike:time Property Let")
End Property

Public Property Get time() As Double
    On Error GoTo timeGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.time
    time = mvartime
    Exit Property


timeGetErr:


    Call RaiseError(MyUnhandledError, "cSpike:time Property Get")
End Property

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

