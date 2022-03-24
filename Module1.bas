Attribute VB_Name = "Module1"
Option Explicit

'�������� ������������ �������
'Class cWU          ���������� � �����
'Class cState       ���������� � ������� ��������� �������
'Class cUserInfo    ���������� � ������������
'Class cOutResult   ������ � ��������� ������� SETI@home
'=========================================================

' Reg Key Security Options...
Public Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const BackSlash = "\"
Public Const Slash = "/"
Public Const strSepURLDir = "/"             '����������� URL-�������
Public Const strSepDir = "\"                '����������� ����������

Public Const iMaxSize = 255

Public Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Public Const gREGVALSYSINFO = "PATH"
Public Const gSETIKEYLOC = "SOFTWARE\SETI@Home"
Public Const gSETIKEYVAL = "ClientDir"

Public Const FileWU = "work_unit.sah"           '���� ������� ������ SETI@home
Public Const FileWULinux = "work_uni.sah"       '�� �� (��� Linux)
Public Const FileUser = "user_info.sah"
Public Const FileUserLinux = "user_inf.sah"
Public Const FileState = "state.sah"
Public Const FileOut = "outfile.sah"
Public Const FileRes = "result.sah"
Public Const Datafile = "SETIdata.txt"
Public Const ResultFile = "SETIres.txt"         '? ��� ���? �����-�� ����� �� ������� ����?
Public Const IndexFileG = "SETItopg.dat"        '����-������ ������ ��������
Public Const IndexFileS = "SETItops.dat"        '����-������ ������ �����
Public Const IndexFileW = "SETItopw.dat"        '����-������ ���������� � ������
Public Const GaussFile = "SETIgaus.dat"         '���� �������� �� result.sah
Public Const SpikeFile = "SETIspik.dat"         '���� ����� �� result.sah
Public Const PulseFile = "SETIpuls.dat"         '���� ��������� �� result.sah
Public Const TripletFile = "SETItrip.dat"       '���� ��������� �� result.sah
Public Const StateFile = "SETIstat.dat"         '���� ������ �������� �� state.sah
Public Const StateCache = "SETIcach.dat"        '���� ����������� �������� �� state.sah
Public Const LinuxPath = "C:\setilin"
Public Const ReportFile = "sreport.txt"   '���� �������� ������ � �����������
Public Const HelpCHMFile = "\SETIhelp.chm"      '���� ������
Public Const strHHelpEXEname = "hh.exe"
Public Const ClientNo = 0
Public Const Client9x = 1
Public Const ClientNT = 2

Public Type tTopG
    ID As Long          '����� �������� �����
    power As Single
    rate As Single
    average As Single
End Type

Public Type tTopS
    ID As Long          '����� �������� �����
    power As Single
    rate As Single
End Type

Public Type tTopW
    ID As Long              '����� �����
    time As String * 24     '���� (� ��������� ���� - ����� 24 �������)
    StartRA As Single       '��������� ����
    StartDEC As Single      '��������� ���������
    freq As Single          '������� ������� (base frequency)
End Type

Public fMainForm As frmMain
Public WU As cWU
Public State As cState
Public UserInfo As cUserInfo
Public OutResult As cOutResult
Public StatusStr As Object          '��������� ��� ��������� ������� � �������
Public SETIpath As String           '������������ ������ SETI@home
Public Result As VbMsgBoxResult     '��� ������ ���������
Public RegRecords As Long           '����� ������� � ������� (�������� �������� � �������)
Public LastRecordNum As Long        '����� ��������� ������ (�������� �������� � �������)
Public bResult As Boolean           '��� �������� �������� ���������� �������
Public EditMode As Boolean          '����������, ������������� �� ������ ����(��. frmInput)
Public NewMode As Boolean           '����������, �������� �� ����� ����
Public EditRowNum As Long           '����� ������� ������� ��� ��������������
Public EditID As Long               'ID ����� �� ������� EditRowNum
Public WinID As Long        'ID �������� ����� (������ ��� Windows)
Public LinID As Long        'ID �������� ����� (������ ��� Linux)
Public showWU As Boolean    '� History ���������� ������ WU ��� State?
Public WUbind As Boolean    '����� AddRecord ������������ ������ ���� ���!
Public Sbind As Boolean     '����� AddRecord ������������ ������ ���� ���!
Public TopG As tTopG            '������ ��� ������ � ������ �������� (SETItopg.dat)
Public TopS As tTopS            '������ ��� ������ � ������ �������� (SETItops.dat)
Public TopW As tTopW            '������ ��� ������ � ������ �������� (SETIwu.dat)
Public MarkerType As Long       '��� �������, ������������� ���� ����
Public MarkerSize As Long       '������ �������
Public RedrawOnStartup As Long  '�������������� �� ����� �������������
Public LastInColor As Long      '�������� ������� ���� ������
Public AutoShowWU As Long       '���������� �� ������� ���������� �������������
Public EnableRegSave As Long    '����� �� ��������� �������� � �������?
Public UpdateOnStartup As Long  '������ �� ��� ������ ��������� state-������? (1=��)
Public AllowAnim As Long        '��������� ��������
Public UseDefaultRF As Long     '������������ ���� �������� ������ �� ���������? (0=���)
Public ReportFileReg As String  '��� ����� �������� ������ (�� �������)
Public AnimTick As Long         '����� ������ ������ ����� (����)
Public SplitterOverwr As Long   '�������������� ��������� ���������� (0=���)
Public DoImport As Long         '������������� �� ������� ��� ���������� �������� ����� (0=���)
Public DoLinux As Long          '������������ �� �������� Linux-������� ��� ������

'��� ���������� ����� ��� ������ �������������� (ViewWU)
Public MaxPscore As Long
''Public MaxPperiod As Long     �� ����� ������
Public MaxPpower As Long
Public MaxTscore As Long
''Public MaxTperiod As Long     �� ����� ������
Public MaxTpower As Long
Public MaxGpower As Long
Public MaxGfit As Long
Public MaxSpower As Long
Public MaxGintegr As Long

'������� DLL-�������
Public Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Sub InitApp()
Dim HistoryWUExist As Boolean           '���� � ���� ������ ��� ���� � ����� �������
    
'INIT MAIN COMPONENTS
    GetRegSettings                      '������ �������� �� �������
    Set WU = New cWU                    '������������� ��������
    Set State = New cState
    Set UserInfo = New cUserInfo
    Set OutResult = New cOutResult
    
    HistoryWUExist = False              '���� ������������� ������ � ������� ����� � �������
    EditMode = False                    '���. ����� ���������� ����� ������� (�� frmHistory)
    Set StatusStr = fMainForm.Label2    '������������� ��������� �� ��������� ������
    WinID = 0       '��������, ����� �� �������� ��������� ���������� ����������
    LinID = 0
    showWU = True   '��-��������� ���������� ������ ������� ������
    WUbind = False  '�������� ������ ��� �� ����
    Sbind = False   '�������� ������ ��� �� ����
    If (Dir(App.path & HelpCHMFile) <> "") Then
        App.HelpFile = App.path & HelpCHMFile
    End If
    
'STAGE 1 - PERFORMING CHECK UP OF THE WORK UNIT FILE
    StatusStr.Caption = "�������� ����� �������..."
    bResult = WU.CheckRegSettings(RegRecords, False)
    If bResult Then
        StatusStr.Caption = "�������� ����� ������� ���������. ������ �� ����������."
    Else
        Result = MsgBox("������ ����� ������� �� �������������" + vbCrLf + "������ � ������� Windows." + vbCrLf + "������ �� �� ����� SETImap ��������� ��� ������?", vbYesNo, "������ �������")
        StatusStr.Caption = "�������� ����� ������� ������� ��������� ��������"
        If Result = vbYes Then
            bResult = WU.CheckRegSettings(RegRecords, True)
            Result = MsgBox("���������� �������� ������� ������� � �������", vbOKOnly, "������ �������")
        Else
            Result = MsgBox("������������ ������ ����� ������� ��������" + vbCrLf + "������ ������ � ������������ ������ ���������." + vbCrLf + "SETImap ��������� ������� �� �������������� �����������.", vbOKOnly + vbExclamation, "������ �������")
            bResult = WU.CheckRegSettings(RegRecords, True)
        End If
        StatusStr.Caption = "�������� ����� ������� ���������: ��� ������ ���������."
    End If
    If LastRecordNum <> WU.GetLastNum Then
        Result = MsgBox("����� ��������� ������ � ������� �� �������������" + vbCrLf + "�������� � ������� Windows. ������ ����� ������������� ����������", vbOKOnly, "������ �������")
        LastRecordNum = WU.GetLastNum
        StatusStr.Caption = StatusStr.Caption + " ����� ��������� ������ � ������� ���������."
        SaveSetting App.Title, "Settings", "LastRecordNum", LastRecordNum
    End If
    
'��������! ���� ������� ����� �������� �� ������ ����� ��� �������� � ������ �������
    'Ver 3.00 ������� ������ � ����� ������
    '�������� ����
    '������ 1 ������ �� SETIstat
    '��������������
    '�������� 1 ������ � SETIex
'''''    Dim i As Long
'''''
'''''    For i = 1 To 197
'''''        Result = State.ReadHistoryEx(i)
'''''        If Result = 0 Then
'''''            bResult = State.DecodeHistoryEx
'''''            bResult = State.EncodeHistory
'''''            bResult = State.WriteHistory(0)
'''''        Else
'''''            Result = MsgBox("Error while reading history record " & Str(i), vbOKOnly, "ERROR")
'''''        End If
'''''
'''''    Next i
'''''    Result = MsgBox("ALL DONE!", vbOKOnly, "SUCCESS")
'''''    Ver 3.00 ��������!!!
'''''    Exit Sub
    
    If WU.existWU Then
        WU.DecodeWU (WU.ReadFile(1))
        Debug.Print WU.Nsamples
        Debug.Print WU.Receiver
        Debug.Print WU.SubbandNum
        Debug.Print WU.UnitName
        If WU.CheckUnit(1, WU.UnitName) Then
            HistoryWUExist = True   '���� ���� ��� ������� � ������
        End If
        If Not (HistoryWUExist) Then    '��� ����� ����� � �������
            WU.NumID = LastRecordNum + 1
            If WU.WriteHistory(WU.EncodeWU, 1) Then
                Result = MsgBox("�������� ������ � ���� �������(WINDOWS)", vbOKOnly, "������ �����(WINDOWS)")
                'WriteHistory ��� ����� ����� �������� �������� RegRecords � LastRecordNum
                ''RegRecords = RegRecords + 1
                ''LastRecordNum = LastRecordNum + 1
                SaveRegSettings '��������� ��������� � �������
                StatusStr.Caption = "�������� � ����� ����� ������ ������� �������� � ������(WINDOWS)."
                State.UpdateRegistry (0)    '�������� ������: ����� ����
            End If
        Else
        '����� ��������� ����� ����� �� �������
            WU.NumID = WU.GetIDbyName(WU.UnitName)
        End If
        WinID = WU.NumID
    Else
        Result = MsgBox("���� ������ �� ���������. ��������, SETI@home" + vbCrLf + "��������� ��������� ���������� � ��������� � ����� � ��������", vbOKOnly, "���� ������ ����������� (WINDOWS client)")
        State.UpdateRegistry (2)    '�������� ������: ������� ��������� ������ ��� ������
        If OutResult.CheckFile(0, 1) Then
            WinID = CLng(Val(WU.GetIDbyName(OutResult.DetectWU(OutResult.ReadFile(0, 1)))))
        End If
        'TO DO - ������ 0 ��������� �������� �� �������
        'If WU.CheckRegSettings(0, False) Then
            'WU.DecodeHistory (WU.ReadHistory(0))
        'End If
        'STAGE 3: UPDATING ALL TIME RESULTS LOG (IF RESULT.SAH EXIST)
        'Windows client
        If Not (WinID = 0) Then '���������� ������ ��� ��������� ������ �����
            If OutResult.CheckFile(0, 1) Then
                bResult = OutResult.Splitter(0, OutResult.ReadFile(0, 1), 1, WinID)
            End If
        Else
            'TO DO ��������� �������� �� ������ ��������
            If DoImport = 1 Then
                
            End If
        End If
    End If
    StatusStr.Caption = "�������� �������� ����� Windows-������� ���������."
    
If DoLinux = 1 Then
    '�������� ����� � ������-�������
    If (Dir(WU.GetFilePath(2), vbNormal) <> "") Then
        WU.DecodeWU (WU.ReadFile(2))
        HistoryWUExist = False
        If WU.CheckUnit(1, WU.UnitName) Then
            HistoryWUExist = True   '���� ���� ��� ������� � ������
        End If
        If Not (HistoryWUExist) Then    '��� ����� ����� � �������
            WU.NumID = LastRecordNum + 1
            If WU.WriteHistory(WU.EncodeWU, 1) Then
                Result = MsgBox("�������� ������ � ���� ������� (LINUX)", vbOKOnly, "������ ����� (LINUX)")
                'WriteHistory ��� ����� ����� �������� �������� RegRecords � LastRecordNum
                ''RegRecords = RegRecords + 1
                ''LastRecordNum = LastRecordNum + 1
                SaveRegSettings '��������� ��������� � �������
                StatusStr.Caption = "�������� � ����� ����� ������ ������� �������� � ������ (LINUX)."
                State.UpdateRegistry (1)    '�������� ������: ����� ����
            End If
        Else
        '����� ��������� ����� ����� �� �������
            WU.NumID = WU.GetIDbyName(WU.UnitName)
        End If
        LinID = WU.NumID
    Else
        Result = MsgBox("���� ������ (������ ��� Linux) �� ���������. ��������, SETI@home" + vbCrLf + "��������� ��������� ���������� � ��������� � �������� �����", vbOKOnly, "���� ������ ����������� (LINUX client)")
        State.UpdateRegistry (3)    '�������� ������: ������� ��������� ������ ��� ������
        If OutResult.CheckFile(1, 1) Then
            LinID = CLng(Val(WU.GetIDbyName(OutResult.DetectWU(OutResult.ReadFile(1, 1)))))
        End If
        'TO DO - ������ 0 ��������� �������� �� �������
        'If WU.CheckRegSettings(0, False) Then
            'WU.DecodeHistory (WU.ReadHistory(0))
        'End If
        'STAGE 3: UPDATING ALL TIME RESULTS LOG (IF RESULT.SAH EXIST)
        'Linux client
        If Not (LinID = 0) Then '���������� ������ ��� ��������� ������ �����
            If OutResult.CheckFile(1, 1) Then
                bResult = OutResult.Splitter(1, OutResult.ReadFile(1, 1), 1, LinID)
            End If
        Else
            'TO DO ��������� �������� �� ������ ��������
            If DoImport = 1 Then
                
            End If
        End If
    End If
    StatusStr.Caption = "�������� �������� ����� Linux-������� ���������."
    
'STAGE 2: CHECKING CURRENT RESULTS
    '���� - ��������� ���� STATE.SAH
    If State.CheckFile(1) Then
        If State.DecodeState(State.ReadFile(1)) Then
            StatusStr.Caption = State.bg_power & "<-gaussian (LINUX) spike ->" & State.bs_power
        End If
        ''TEMPORARY!!!!
        ''LinID = 63
        Debug.Print "LINUX client -> block num " & LinID
        State.Interchange 1
        If State.EncodeHistory Then
            If State.WriteHistory(1, LinID) Then
            End If
        End If
    End If
    StatusStr.Caption = "�������� ������� ����������� Linux-������� ���������."
End If  'Perform Linux client check-up

    If State.CheckFile(0) Then
        If State.DecodeState(State.ReadFile(0)) Then
            StatusStr.Caption = StatusStr.Caption & "  " & State.bg_power & "<-gaussian (WINDOWS) spike ->" & State.bs_power
        End If
        Debug.Print "WINDOWS client -> block num " & WinID
        State.Interchange 0
        If State.EncodeHistory Then
            If State.WriteHistory(1, WinID) Then
            End If
        End If
    End If
    StatusStr.Caption = "�������� ������� ����������� Windows-������� ���������."
    StatusStr.Caption = "�������� ������������� ����������� ������� ���������."
    
    frmMain.RunServices     '��������� �������, ��������������� �� ��������
    
''    ��������� ��� � Load ��� Init ���� ������� ����� �� ���������!
''    WU.AddRecord 0  '��������� ��� �������� �� ������� � ��������� ���������� �� �����
End Sub

Sub Main()
Dim PauseTime, Start, Finish

    frmSplash.Show
    frmSplash.Refresh
    PauseTime = 1   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents   ' Yield to other processes.
    Loop
    Finish = Timer   ' Set end time.
    If Not (GetKeyValue(HKEY_LOCAL_MACHINE, gSETIKEYLOC, gSETIKEYVAL, SETIpath)) Then
        Result = MsgBox("������ ��� ������� ����� ������������ SETI@home", vbOKOnly, "CRITICAL ERROR")
        End
    End If
    Set fMainForm = New frmMain
    Load fMainForm
    ''Unload frmSplash

    fMainForm.Show
    Unload frmSplash
    Call InitApp
End Sub

'**********************************************************
'*     ����������� "����������" ���� �� �� ����������     *
'*                � ������������ ������                   *
'**********************************************************
Function ExtractTime(sTime As String) As String
Dim i As Long
Dim res As String
    i = InStr(1, sTime, "(", vbTextCompare)
    res = ""
    res = Mid(sTime, i + 1, InStr(i + 1, sTime, ")", vbTextCompare) - i - 1)
    ExtractTime = res
End Function

'**********************************************************
'*            ����������� ���� �� ����������              *
'*          ������������� ����� � ������� ����            *
'**********************************************************
Function ExtractDigTime(sTime As String) As String
Dim res As String
    res = ""
    res = Left(sTime, InStr(1, sTime, "(", vbTextCompare) - 1)
    If Right(res, 1) = "" Then
        res = Left(res, Len(res) - 1)
    End If
    ExtractDigTime = res
End Function

'**********************************************************
'*       ��������� ������ � �����. ��� �������������      *
'*    ������ ���������� ������� ������ ����� �� �������   *
'*      ���������� -1 ���� �������� �� �������� �����     *
'**********************************************************
Function MyStrToFloat(s As String) As Double
Dim f As Double

On Error GoTo StrToFloatErr
    If (InStr(1, s, ".", vbTextCompare) <= 0) Then
        If (InStr(1, s, ",", vbTextCompare) <= 0) Then
            'Result = MsgBox("It is not a date value!!!", vbOKOnly, "Invalid string")
            MyStrToFloat = -1
            Exit Function
        End If
    End If
    f = CDbl(Val(s))
    MyStrToFloat = f
    Exit Function
StrToFloatErr:
    If InStr(1, s, ".", vbTextCompare) > 0 Then
        's(InStr(1, s, ".", vbTextCompare)) = ","
        s = Replace(s, ".", ",", 1, , vbTextCompare)
    End If
    f = CDbl(Val(s))
    MyStrToFloat = f
End Function

'**********************************************************
'*     ��������� ������ ����������� ������, ��������,     *
'*      ������ "1 �������" ����� ������ "01 �������"      *
'**********************************************************
Function LeftZero(s As String, i As Long) As String
    If Len(s) = i Then
        LeftZero = "0" + s
    Else
        LeftZero = s
    End If
End Function

'**********************************************************
'*           ����������� ������� ����� ���������          *
'**********************************************************
Function Trunc(dValue As Double) As Double
Dim s As String, tmp As String
Dim i As Long
    s = CStr(dValue)
    If InStr(1, s, "E-", vbTextCompare) > 0 Then
        'Scientific format, negative power
        s = 0
    ElseIf InStr(1, s, "E", vbTextCompare) > 0 Then
        'Scientific format, positive power
        tmp = s
        If InStr(1, s, ".", vbTextCompare) > 0 Then
            s = Left(s, InStr(1, s, ".", vbTextCompare) - 1)
        ElseIf InStr(1, s, ",", vbTextCompare) > 0 Then
            s = Left(s, InStr(1, s, ",", vbTextCompare) - 1)
        End If
        For i = 1 To Val(Right(tmp, Len(s) - InStr(1, tmp, "E", vbTextCompare) - 1))
            s = s & "0"
        Next i
    ElseIf InStr(1, s, ".", vbTextCompare) > 0 Then
        s = Left(s, InStr(1, s, ".", vbTextCompare) - 1)
    ElseIf InStr(1, s, ",", vbTextCompare) > 0 Then
        s = Left(s, InStr(1, s, ",", vbTextCompare) - 1)
    End If
    Trunc = CDbl(s)
End Function

'**********************************************************
'*      ����������� ������� ����� �������� � ��������     *
'*                  ��������� ������                      *
'**********************************************************
Function CRtoPercent(cr As Single) As Long
Dim TMPvalue As Long
Dim Negative As Boolean
Dim fResult As Long
    Negative = False
    TMPvalue = Trunc(cr * 100)
    If TMPvalue < 0 Then
        Negative = True
        TMPvalue = Abs(TMPvalue)
    End If
    If TMPvalue < 500 Then
    '����� �������� ����� 5
        fResult = (TMPvalue * 25) \ 500
    Else
    '����� �������� ����� 5
        fResult = 25 + (TMPvalue - 500) \ 180
    End If
    If Negative Then
        fResult = fResult + 50
    End If
    If fResult > 100 Then
        fResult = 100
    End If
    CRtoPercent = fResult
End Function

'**********************************************************
'*     ���������� ����� � ������� SETI@home � �������     *
'*              ���, ����, ������ � �������               *
'**********************************************************
Function DecodeTime(dTime As Double, bDay As Boolean) As String
    If Not bDay Then
        DecodeTime = LeftZero(CStr(dTime \ 3600), 1) + " ��� " + LeftZero(CStr(Trunc((dTime - ((dTime \ 3600) * 3600)) / 60)), 1) + " ��� " + LeftZero(CStr(Trunc(((dTime * 60) - Trunc(dTime * 60)) * 60)), 1) + " ���"
    Else
        DecodeTime = CStr((dTime \ 86400)) + " ���� " + LeftZero(CStr(Trunc((dTime - ((dTime \ 86400) * 86400)) / 3600)), 1) + " ��� " + LeftZero(CStr(Trunc((dTime - ((dTime \ 3600) * 3600)) / 60)), 1) + " ��� " + LeftZero(CStr(Trunc(((dTime * 60) - Trunc(dTime * 60)) * 60)), 1) + " ���"
    End If
End Function

'**********************************************************
'*      ���������� RA-���������� � ������� SETI@home      *
'*               � ����, ������ � �������                 *
'**********************************************************
Function DecodeRA(ra As Double) As String
    DecodeRA = LeftZero(CStr(Trunc(ra)), 1) + " ��� " + LeftZero(CStr(Trunc((ra - Trunc(ra)) * 60)), 1) + " ��� " + LeftZero(CStr(Trunc(((ra * 60) - Trunc(ra * 60)) * 60)), 1) + " ���"
End Function

'**********************************************************
'*     ���������� DEC-���������� � ������� SETI@home      *
'*             � �������, ������ � �������                *
'**********************************************************
Function DecodeDEC(dec As Double) As String
    DecodeDEC = LeftZero(CStr(Trunc(dec)), 1) + " ���� " + LeftZero(CStr(Trunc((dec - Trunc(dec)) * 60)), 1) + " ��� " + LeftZero(CStr(Trunc(((dec * 60) - Trunc(dec * 60)) * 60)), 1) + " ���"
'//Catching these strange "0 degrees 300 minutes 300 seconds" report - SUCCESS
'//Bug fixed - Trunc function has beed modified in order to handle
'//numbers in scientific format (like 1.2345E-06)
'    If Trunc((DEC - Trunc(DEC)) * 60) = 300 Then
'        Debug.Print "Error reporting!"
'        Debug.Print DEC
'        Debug.Print Trunc((DEC - Trunc(DEC)) * 60)
'        Debug.Print LeftZero(CStr(Trunc((DEC - Trunc(DEC)) * 60)), 1) + " ��� "
'        Debug.Print "-----------------------------------------"
'    End If
End Function

'**********************************************************
'*       ���������� ���� �� ������������� ���������       *
'**********************************************************
Function GetHourStr(dTime As Double) As String
'procedure SplitCoor(time : real; var hr, min, sec :string);
    dTime = Abs(dTime)
    GetHourStr = LeftZero(CStr(Trunc(dTime)), 1)
End Function

'**********************************************************
'*     ���������� ������ �� ������������� ���������       *
'**********************************************************
Function GetMinStr(dTime As Double) As String
    dTime = Abs(dTime)
    GetMinStr = LeftZero(CStr(Trunc((dTime - Trunc(dTime)) * 60)), 1)
End Function

'**********************************************************
'*    ���������� ������� �� ������������� ���������       *
'**********************************************************
Function GetSecStr(dTime As Double) As String
    dTime = Abs(dTime)
    GetSecStr = LeftZero(CStr(Trunc(((dTime * 60) - Trunc(dTime * 60)) * 60)), 1)
End Function

'**********************************************************
'*     ������������� ���������� � ������ SETI@home        *
'**********************************************************
Function EncodeCoor(hr As String, min As String, sec As String) As String
Dim res As String
    If Len(hr) > 1 Then
        If Not (hr = "00") Then
            Do While (Not (Left(hr, 1) Like "[1-9]"))
                hr = Right(hr, Len(hr) - 1)
            Loop
            '������ ������ ����� (�������� ���� �����), ������� ��������� ���� �� �����
            ''Do While (Not (Right(hr, 1) Like "[0-9]"))
                ''hr = Left(hr, Len(hr) - 1)
            ''Loop
        Else
            hr = "0"
        End If
    ElseIf hr = "" Then
        hr = "0"
    End If
    res = hr + "." + LeftZero(CStr(Round((CInt(Val(min)) * 100) / 6 + (CInt(Val(sec)) * 10) / 36)), 2)
    EncodeCoor = res
End Function

'LINUX compatible
'**********************************************************
'*           ��������� �������� ��������                  *
'* tokenname: �������� ���������                          *
'*    psfile: ������, � ������� ������������ �����        *
'*   Stopper: ������, �������� ������������ �������       *
'**********************************************************
Public Function GetToken(ByVal tokenname As String, ByVal psfile As String, ByVal stopper As String) As String
Dim res As String
Dim i As Long, StartPos As Long, EndPos As Long
    On Error GoTo GetTokenErr
    
    res = ""
    If stopper = "space" Then
        stopper = " "
    End If
    i = InStr(1, psfile, tokenname, vbTextCompare)  '����� ��������� ��������� � ������
    If i <> 0 Then
        StartPos = i + Len(tokenname)   '������������ ������ �� ����� �������� ���������
        Do While (Mid(psfile, StartPos, 1) = " ")
            StartPos = StartPos + 1
        Loop
        EndPos = InStr(StartPos, psfile, stopper, vbTextCompare)    '����� ����������� ������
        res = Mid(psfile, StartPos, EndPos - StartPos)
    End If
    'Trim spaces
    Do While (Left(res, 1) Like " ")
        res = Right(res, Len(res) - 1)
    Loop
    Do While (Right(res, 1) Like " ")
        res = Left(res, Len(res) - 1)
    Loop
    GetToken = res
    Exit Function

GetTokenErr:
    Call RaiseError(MyUnhandledError, "cState:GetToken Method")
End Function

'LINUX compatible
'**********************************************************
'*   ��������� ��������, ������� �� ����� ���� ��������   *
'*   �������� GetToken. ��������� ��� ������� �����       *
'*   tokenname � stopper, �� ����������� ��������         *
'*   �������� ������ � (�����������) ��������.            *
'* tokenname: �������� ���������                          *
'*   stopper: ������ ��������, �������� �������������     *
'*    psfile: ������, � ������� ������������ �����        *
'**********************************************************
Public Function GetTokenEx(ByVal tokenname As String, ByVal psfile As String, ByVal stopper As String, ByVal SpacesStay As Boolean) As String
Dim res As String, TMPstr As String
Dim i As Long, StartPos As Long, EndPos As Long
    On Error GoTo GetTokenExErr
    
    res = ""
    i = InStr(1, psfile, tokenname, vbTextCompare)  '����� ��������� ��������� � ������
    If i <> 0 Then
        StartPos = i + Len(tokenname)   '������������ ������ �� ����� �������� ���������
        i = 0
        i = InStr(StartPos, psfile, stopper, vbTextCompare) '������� ������������
        If i <> 0 Then
            '���������� ������ ������ ���� ������ ������������, ����� - �����
            EndPos = i
            For i = 0 To EndPos - StartPos - 1
                TMPstr = Mid(psfile, StartPos + i, 1)
                If Not (TMPstr = Chr(10)) Then
                    If Not (TMPstr = Chr(13)) Then  '������ �������� ������
                        If TMPstr = " " Then
                            If SpacesStay Then      '������ ���������� ������ ���� ��������
                                res = res & TMPstr
                            End If
                        Else
                            res = res & TMPstr
                        End If
                    End If
                End If
            Next i
            'Trim asterisks
            If Right(res, 1) = "*" Then
                res = Left(res, Len(res) - 1)
            End If
        End If
    End If
    GetTokenEx = res
    Exit Function

GetTokenExErr:
    Call RaiseError(MyUnhandledError, "Module1:GetTokenEx Method")
End Function

'**********************************************************
'*        ������ �� ������� ��������� ���������           *
'**********************************************************
Public Sub GetRegSettings()
    '��������� ��������
    RegRecords = GetSetting(App.Title, "Settings", "NumOfHistoryRec", 0)
    LastRecordNum = GetSetting(App.Title, "Settings", "LastRecordNum", 0)
    SplitterOverwr = GetSetting(App.Title, "Settings", "SplitterOverwrite", 0)
    '��������� �����
    MarkerType = GetSetting(App.Title, "Starmap", "MarkerType", 0)
    MarkerSize = GetSetting(App.Title, "Starmap", "MarkerSize", 0)
    RedrawOnStartup = GetSetting(App.Title, "Starmap", "RedrawOnStartup", 0)
    LastInColor = GetSetting(App.Title, "Starmap", "LastInColor", 0)
    '��������� ���������
    AutoShowWU = GetSetting(App.Title, "Settings", "AutoShowWU", 0)
    EnableRegSave = GetSetting(App.Title, "Settings", "EnableRegSave", 1)
    UpdateOnStartup = GetSetting(App.Title, "Settings", "UpdateOnStartup", 1)
    AllowAnim = GetSetting(App.Title, "Settings", "AllowAnim", 1)
    ReportFileReg = GetSetting(App.Title, "Settings", "ReportFile", "")
    UseDefaultRF = GetSetting(App.Title, "Settings", "UseDefaultReportFile", 1)
    AnimTick = GetSetting(App.Title, "Settings", "AnimationTick", 50)
    DoLinux = GetSetting(App.Title, "Settings", "DoLinux", 0)
    '��������� �������������� (ViewWU)
    MaxPscore = GetSetting(App.Title, "AutoRange", "MaxPscore", 0)
    MaxPpower = GetSetting(App.Title, "AutoRange", "MaxPpower", 0)
    MaxTscore = GetSetting(App.Title, "AutoRange", "MaxTscore", 0)
    MaxTpower = GetSetting(App.Title, "AutoRange", "MaxTpower", 0)
    MaxGpower = GetSetting(App.Title, "AutoRange", "MaxGpower", 0)
    MaxGfit = GetSetting(App.Title, "AutoRange", "MaxGfit", 2500000)
    MaxGintegr = GetSetting(App.Title, "AutoRange", "MaxGintegr", 0)
    MaxSpower = GetSetting(App.Title, "AutoRange", "MaxSpower", 0)
    DoImport = GetSetting(App.Title, "Settings", "DoImport", 0)
End Sub

'**********************************************************
'*        ��������� � ������� ��������� ���������         *
'**********************************************************
Public Sub SaveRegSettings()
    SaveSetting App.Title, "Settings", "NumOfHistoryRec", RegRecords
    SaveSetting App.Title, "Settings", "LastRecordNum", LastRecordNum
End Sub

'********************************************
'* ��������� ���� ��������� Windows �����   *
'* Win API                                  *
'* ���������� ���� �������� �����������     *
'* ����������� ���������� \                 *
'********************************************
Function GetWindowsDir() As String
Dim strBuf As String
Dim iZeroPos As Integer

    '��������� ����� ���������
    strBuf = Space(iMaxSize)
    If GetWindowsDirectory(strBuf, iMaxSize) > 0 Then
        '���� ���������� ������
        iZeroPos = InStr(strBuf, Chr$(0))
        '���� ���������� ����, �� ������� ���
        If iZeroPos > 0 Then
            strBuf = Left$(strBuf, iZeroPos - 1)
        End If
        '���� �� ����� ������ ��� ����������� ����������, ��������� ���
        If Right(Trim(strBuf), Len(strSepURLDir)) <> strSepURLDir And _
           Right(Trim(strBuf), Len(strSepDir)) <> strSepDir Then
            strBuf = RTrim$(strBuf) & strSepDir
        End If
        GetWindowsDir = strBuf
    Else
        GetWindowsDir = vbNullString
    End If
End Function

'************************************************************
'* ������ ���������� ������� Windows (������ ������� *.CHM) *
'* ����� ����� hh.exe ����� ������ ������������� �� �����,  *
'* ��������� �� ��, ��� ���� ���� � ����������� �������     *
'* ����� � ����� Windows                                    *
'************************************************************
Public Sub ShowCHMHelp()
Dim RetValue As Double
    '�������� ���� � ����� Windows ����� DLL call
    RetValue = Shell(GetWindowsDir & strHHelpEXEname & Chr(32) & App.path & HelpCHMFile, vbMaximizedFocus)
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

''���� ���� ����, �� ��������� ����� State
''�����
''��������� State (� ������ �������� ID)
''��������� Result
