VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class working with user info file"
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
Private mvarID As String 'local copy
Private mvarKey As String 'local copy
Private mvaremail As String 'local copy
Private mvarUserName As String 'local copy
Private mvarURL As String 'local copy
Private mvarcountry As String 'local copy
Private mvarPostalCode As String 'local copy
Private mvarShowName As String 'local copy
Private mvarShowEmail As String 'local copy
Private mvarVenue As String 'local copy
Private mvarRegister As String 'local copy
Private mvarLastWU As String 'local copy
Private mvarLastResult As String 'local copy
Private mvarNwus As String 'local copy
Private mvarNResults As String 'local copy
Private mvartotalCPU As String 'local copy
Private mvarhFile As Long 'local copy

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

Public Property Let totalCPU(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.totalCPU = 5
    mvartotalCPU = vData
End Property

Public Property Get totalCPU() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.totalCPU
    totalCPU = mvartotalCPU
End Property

Public Property Let NResults(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NResults = 5
    mvarNResults = vData
End Property

Public Property Get NResults() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NResults
    NResults = mvarNResults
End Property

Public Property Let Nwus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nwus = 5
    mvarNwus = vData
End Property

Public Property Get Nwus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nwus
    Nwus = mvarNwus
End Property

Public Property Let LastResult(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastResult = 5
    mvarLastResult = vData
End Property

Public Property Get LastResult() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastResult
    LastResult = mvarLastResult
End Property

Public Property Let LastWU(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastWU = 5
    mvarLastWU = vData
End Property

Public Property Get LastWU() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastWU
    LastWU = mvarLastWU
End Property

Public Property Let Register(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Register = 5
    mvarRegister = vData
End Property

Public Property Get Register() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Register
    Register = mvarRegister
End Property

Public Property Let Venue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Venue = 5
    mvarVenue = vData
End Property

Public Property Get Venue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Venue
    Venue = mvarVenue
End Property

Public Property Let ShowEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShowEmail = 5
    mvarShowEmail = vData
End Property

Public Property Get ShowEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShowEmail
    ShowEmail = mvarShowEmail
End Property

Public Property Let ShowName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShowName = 5
    mvarShowName = vData
End Property

Public Property Get ShowName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShowName
    ShowName = mvarShowName
End Property

Public Property Let PostalCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostalCode = 5
    mvarPostalCode = vData
End Property

Public Property Get PostalCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostalCode
    PostalCode = mvarPostalCode
End Property

Public Property Let country(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.country = 5
    mvarcountry = vData
End Property

Public Property Get country() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.country
    country = mvarcountry
End Property

Public Property Let URL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.URL = 5
    mvarURL = vData
End Property

Public Property Get URL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.URL
    URL = mvarURL
End Property

Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property

Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property

Public Property Let email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.email = 5
    mvaremail = vData
End Property

Public Property Get email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.email
    email = mvaremail
End Property

Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property

'LINUX compatible
'**********************************************************
'*                 Прочитать файл                         *
'*  Mode=0: файл типа user_info.sah (клиент для Windows)  *
'*  Mode=1: файл типа user_inf.sah  (клиент для Linux)    *
'**********************************************************
Public Function ReadFile(ByVal Mode As Long) As String
Dim sfile As String
Dim nextchar As String
    On Error GoTo ReadFileErr

    sfile = ""
    If CheckFile(Mode) Then    'OK, file exist
        hFile = FreeFile                        'Получить свободный указатель
        Open GetFileName(Mode) For Input As hFile
        Debug.Print "Success file open (UserInfo.ReadFile Method)"
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
    Call RaiseError(MyUnhandledError, "cUserInfo:ReadFile Method")
End Function

'LINUX compatible
'**********************************************************
'*           Проверить существование файла                *
'*  Mode=0: файла типа user_info.sah (клиент для Windows) *
'*  Mode=1: файла типа user_inf.sah (клиент для Linux)    *
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
    Call RaiseError(MyUnhandledError, "cUserInfo:CheckFile Method")
End Function

'LINUX compatible
'**********************************************************
'*             Получить имя файла данных                  *
'*  Mode=0: Получить расположение файла типа user_info.sah*
'*          (клиент для Windows)                          *
'*  Mode=1: Получить расположение файла типа user_inf.sah *
'*          (клиент для Linux)                            *
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
            Case 0: 'Mode=0: Получить расположение файла типа user_info.sah (Windows)
                sfile = SETIpath & "\" & FileUser
            Case 1: 'Mode=1: Получить расположение файла типа user_inf.sah (Linux)
                sfile = LinuxPath & "\" & FileUserLinux
        End Select
    End If
    GetFileName = sfile 'Выдать полученный результат вызывающей функции
    Exit Function
    
GetFileNameErr:
    Call RaiseError(MyUnhandledError, "cUserInfo:GetFileName Method")
End Function

'LINUX compatible
'**********************************************************
'*   Дешифрует заголовок блока из строки, полученной при  *
'*               чтении файла user_info.sah               *
'*       sfile - строка, полученная от  ReadFile          *
'**********************************************************
Public Function DecodeInfo(sfile As String) As Boolean
Dim stopper As String
    On Error GoTo DecodeInfoErr
            
        stopper = "*"
        ID = GetToken("id=", sfile, stopper)
        Key = GetToken("key=", sfile, stopper)
        email = GetToken("email_addr=", sfile, stopper)
        UserName = GetToken("name=", sfile, stopper)
        URL = GetToken("url=", sfile, stopper)
        country = GetToken("country=", sfile, stopper)
        PostalCode = GetToken("postal_code=", sfile, stopper)
        ShowName = GetToken("show_name=", sfile, stopper)
        ShowEmail = GetToken("show_email=", sfile, stopper)
        Venue = GetToken("venue=", sfile, stopper)
        Register = GetToken("register_time=", sfile, stopper)
        LastWU = GetToken("last_wu_time=", sfile, stopper)
        LastResult = GetToken("last_result_time=", sfile, stopper)
        Nwus = GetToken("nwus=", sfile, stopper)
        NResults = GetToken("nresults=", sfile, stopper)
        totalCPU = GetToken("total_cpu=", sfile, stopper)
        DecodeInfo = True
    Exit Function
    
DecodeInfoErr:
    DecodeInfo = False
    Call RaiseError(MyUnhandledError, "cUserInfo:DecodeInfo Method")
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
