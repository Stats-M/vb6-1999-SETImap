VERSION 5.00
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "����������:"
      Height          =   1065
      Left            =   7770
      TabIndex        =   5
      Top             =   1155
      Width           =   1485
      Begin VB.OptionButton Option1 
         Caption         =   "���������"
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   630
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      Default         =   -1  'True
      Height          =   345
      Left            =   7770
      TabIndex        =   4
      Top             =   630
      Width           =   1467
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "�������"
      Height          =   345
      Left            =   7770
      TabIndex        =   3
      Top             =   3150
      Width           =   1467
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��������� ������ ������"
      Height          =   540
      Left            =   7770
      TabIndex        =   2
      Top             =   2415
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      HideSelection   =   0   'False
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   7365
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   7365
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
Dim j As Long
Dim ErrTrap As Boolean
Dim OneStr  As String

    ErrTrap = False
    If Option1(0).Value Then
    '�������� ������ �����
        Label1.Caption = "�������� ������� - ����� ����� - �������"
        If Check1.Value = vbChecked Then
            '��������� ������ ����� ������
            bResult = State.RebuildIndex(1, 0)
        Else
            '�������� � ������� ����� ������ ������ ��� �������� �����
            bResult = State.RebuildIndex(0, 0, State.GetLastRecNum(2))
        End If
        If bResult Then
            Text1.text = ""
            For i = 1 To State.GetLastRecNum(4)
                If State.ReadIndex(0, i) Then   'Target=0 - ����
                    OneStr = Format(TopS.power, "####.0000") & "   "
                    '����������� ������ ������
                    j = Len(Str(TopS.ID))
                    For j = j To 4
                        OneStr = OneStr & " "
                    Next j
                    OneStr = OneStr & Format(TopS.ID, "#####")
                    '����������� ���������� ������
                    j = Len(Str(i))
                    For j = j To 5
                        OneStr = OneStr & " "
                    Next j
                    OneStr = OneStr & Format(i, "#####")
                    
                    OneStr = OneStr & vbCrLf
                    Text1.text = Text1.text + OneStr
                Else
                    ErrTrap = True
                End If
            Next i
            If ErrTrap Then
                '������ ��� ������ ������ ����� ������� �����
                Call RaiseErrMsg(1207, StandartErrHeader)
            End If
        Else
            '������ ��� ���������� ������� �����
            Call RaiseErrMsg(1208, StandartErrHeader)
        End If
    Else
    '�������� ������ ��������
        Label1.Caption = "������������ ���������� - ����� ����� - �������"
        If Check1.Value = vbChecked Then
            '��������� ������ �������� ������
            bResult = State.RebuildIndex(1, 1)
        Else
            '�������� � ������� �������� ������ ������ ��� �������� �����
            bResult = State.RebuildIndex(0, 1, State.GetLastRecNum(2))
        End If
        If bResult Then
            Text1.text = ""
            For i = 1 To State.GetLastRecNum(3)
                If State.ReadIndex(1, i) Then   'Target=1 - ���������
                    OneStr = Format(TopG.average, "0.0000000") & "   "
                    '����������� ������ ������
                    j = Len(Str(TopG.ID))
                    For j = j To 4
                        OneStr = OneStr & " "
                    Next j
                    OneStr = OneStr & Format(TopG.ID, "#####")
                    '����������� ���������� ������
                    j = Len(Str(i))
                    For j = j To 5
                        OneStr = OneStr & " "
                    Next j
                    OneStr = OneStr & Format(i, "#####")
                    
                    OneStr = OneStr & vbCrLf
                    Text1.text = Text1.text + OneStr
                Else
                    ErrTrap = True
                End If
            Next i
            If ErrTrap Then
                '������ ��� ������ ������ ����� ������� ��������
                Call RaiseErrMsg(1200, StandartErrHeader)
            End If
        Else
            '������ ��� ���������� ������� ��������
            Call RaiseErrMsg(1201, StandartErrHeader)
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    ''Me.Caption = "���������� - ������ ���������"
    Me.Caption = "����������"
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Me.Caption = "���������� - ������ ����"
            Option1(0).Value = True '���������� ����
            Check1.Value = vbUnchecked
        Case 1:
            Me.Caption = "���������� - ������ ���������"
            Option1(1).Value = True '���������� ���������
            Check1.Value = vbUnchecked
    End Select
End Sub
