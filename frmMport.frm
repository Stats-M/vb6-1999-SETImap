VERSION 5.00
Begin VB.Form frmMport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������� ������"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3 
      Caption         =   "�������"
      Height          =   345
      HelpContextID   =   15001
      Left            =   5460
      TabIndex        =   3
      Top             =   2520
      Width           =   1467
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "�������"
      Height          =   345
      HelpContextID   =   15002
      Left            =   5460
      TabIndex        =   2
      Top             =   3045
      Width           =   1467
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      Height          =   3270
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   105
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������ !"
      Default         =   -1  'True
      Height          =   345
      HelpContextID   =   15000
      Left            =   5460
      TabIndex        =   0
      Top             =   210
      Width           =   1467
   End
End
Attribute VB_Name = "frmMport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim STRout As String       '��������� ��������� (������������ ��������� � ������)

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Text1.text = ""
    Call PosInIndex
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Caption = "������������� ����������"
End Sub

Private Sub PosInIndex()
Dim Place As Long   '��������� ���������� ��� �������� ����� ���������� � ����� �������
Dim IndexLen As Long    '���������� ������� � �������
Dim i As Long
Dim Summa As Single     '����� ��������� ��������
Dim NumValid As Long    '���������� ��������� ��������
    IndexLen = State.GetLastRecNum(4)  '������ � �������� �����
    Place = -1
    For i = 1 To IndexLen
        bResult = State.ReadIndex(0, i)
        'TO DO
        '������� ����� �������� Win/Lin...
        If TopS.ID = WinID Then
            Place = i       '��������� �����
            i = IndexLen    '��������� ����
        End If
    Next i
    
    If Place = -1 Then      '������ �� �������!
        STRout = "������ � ������ ������� �������� (������� ���� " & Str(WinID) & ") �� ����������. "
        STRout = STRout & "��������, �� ��� �� ��������� ���������� ���� ����������. "
        STRout = STRout & "�������� ����� ���� ����������/������ ���� ��� ����������/������ ���������. "
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    Else
        STRout = "������� ������ ������� �������� ����� ��������� �� " & Place & " ����� �����" & Str(i - 1) & " ������������ �� ��� �����. " & vbCrLf & vbCrLf
        If TopS.power > 400 Then
            STRout = STRout & "����������� ������� �������� �������. ��������, ��� �������� ������, ������������� � ��������� ������� ����� ��� �������� ������������ ������ ��������. "
        ElseIf TopS.power > 250 Then
            STRout = STRout & "����� ������� �������� �������, ������ �����, �� ����� ������ ����� ����� ������� ���� �������� ����� �� ������������� ������� �����. "
        ElseIf TopS.power > 200 Then
            STRout = STRout & "������� �������� �������, ��������, ��� ��������� ��������� ������-���� ������������ ��������. "
        ElseIf TopS.power > 180 Then
            STRout = STRout & "������� �������� �������, ��������, ��� ����������� ������ ������� �������������. "
        Else
            STRout = STRout & "�������� ������� ��������, ��������, ��� ������, ��������� �� �������. "
        End If
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    End If
    
    IndexLen = State.GetLastRecNum(3)  '������ � �������� ��������
    Place = -1
    For i = 1 To IndexLen
        bResult = State.ReadIndex(1, i)
        'TO DO
        '������� ����� �������� Win/Lin...
        If TopG.ID = WinID Then
            Place = i       '��������� �����
            i = IndexLen    '��������� ����
        End If
    Next i
    
    If Place = -1 Then      '������ �� �������!
        STRout = "������ � ������ ���������� (������� ���� " & Str(WinID) & ") �� ����������. "
        STRout = STRout & "��������, �� ��� �� ��������� ���������� ���� ����������. "
        STRout = STRout & "�������� ����� ���� ����������/������ ���� ��� ����������/������ ���������. "
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    Else
        STRout = "������ ��������� ������� �������� ����� ��������� �� " & Place & " ����� �����" & Str(i - 1) & " ������������ �� ��� �����. " & vbCrLf & vbCrLf
        '������ ������������� ����������
        If TopG.average > 3 Then
            STRout = STRout & "������ ����� �����, ��������� ������� � ���������! "
            STRout = STRout & "���� ���� ������� � ����� ������ �� �������� �������� �������. "
        ElseIf TopG.average > 0.7 Then
            STRout = STRout & "������ ����� �����, ������ �������������� � ���������."
        ElseIf TopG.average > 0.3 Then
            STRout = STRout & "������ ����� �����, ���������� ������� � ���������."
        ElseIf TopG.average > 0.2 Then
            STRout = STRout & "������ ����� �����, ������� ������� �� ���������."
        ElseIf TopG.average > 0.18 Then
            STRout = STRout & "������ ����� �����, ����� ��������������� ���������."
        ElseIf TopG.average = 0 Then
            STRout = STRout & "������� ������������� �������. ����� �������� �� ������������."
        Else
            STRout = STRout & "������ ����� �����, ���� �������������� �� ��������������� ���������."
        End If
        STRout = STRout & " "
        Text1.text = Text1.text & STRout
        '������ �������� �������
        If TopG.power > 2 Then
            STRout = "����� ������� �������� ������� � ������� ����� ��������� ��������������� � ���, "
            STRout = STRout & "��� ��� �������� ��������� ����� ������ �� ����� � "
            STRout = STRout & "������ �� �������� ��������� �� ������� �������� ��� ��������������� � �������."
        ElseIf TopG.power > 1.4 Then
            STRout = "�������� ������� � ������� ����� ��������� �������, "
            STRout = STRout & "�������� ������� ���� ��������� �� ������ ����� (������������� �������������, �������� ��������� ������), "
            STRout = STRout & "���� (������������) ���������� ����� �� ������ ������������� ��� ���������."
        ElseIf TopG.power > 0.75 Then
            STRout = "�������� ������� � ������� ����� ��������� �������, "
            STRout = STRout & "� ������ ����� ����������� ����� ������������ ��� ����������� ��������� ������� �������������, "
            STRout = STRout & "��� � �������� �� ����� ��������� ����� �� ������ ������������� ��� ���������."
        ElseIf TopG.power = 0 Then
            STRout = "�������� ��������� ���������� ����������."
        Else
            STRout = "������ �������� � ������� ����� ���������, "
            STRout = STRout & "��������� �����, ��� ������, ��������� �� �������� ������� (��������������, ��������), ���� "
            STRout = STRout & "���� ������ ����������� ��������� ������� �������������."
        End If
        STRout = STRout & " "
        Text1.text = Text1.text & STRout
        '�������������� ������
        Summa = 0
        NumValid = 0
        For i = 1 To IndexLen                   '������� ������� ��������
            bResult = State.ReadIndex(1, i)
            'TO DO
            '������� ����� �������� Win/Lin...
            If Not (TopG.average = 0) Then
                Summa = Summa + TopG.average    '��������� ��� ��������� ��������
                NumValid = NumValid + 1
            End If
        Next i
        bResult = State.ReadIndex(1, Place)
        If Not (NumValid = 0) Then
            Summa = Summa / NumValid
        End If
        If Not (TopG.average = 0) Then
            If 1.5 * Summa < TopG.average Then
                STRout = "������������ ���������� �������� ����� ����������� ���� ���������������������."
            ElseIf Summa < TopG.average Then
                STRout = "������������ ���������� �������� ����� ���� ���������������������."
            ElseIf Summa > TopG.average Then
                If Summa > 1.5 * TopG.average Then
                    STRout = "������������ ���������� �������� ����� ����������� ���� ���������������������."
                Else    'Summa > TopG.average
                    STRout = "������������ ���������� �������� ����� ���� ���������������������."
                End If
            End If
        End If
        STRout = STRout & vbCrLf & vbCrLf
        Text1.text = Text1.text & STRout
    End If
End Sub
