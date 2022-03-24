VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������� ����������� ������"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2835
      TabIndex        =   7
      Top             =   1335
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2730
      TabIndex        =   5
      Top             =   2310
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   945
      TabIndex        =   4
      Top             =   2310
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2835
      TabIndex        =   3
      Top             =   1755
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "���������� ����� �����"
      Height          =   225
      Left            =   525
      TabIndex        =   6
      Top             =   1365
      Width           =   2220
   End
   Begin VB.Label Label4 
      Caption         =   "��� �����"
      Height          =   225
      Left            =   525
      TabIndex        =   2
      Top             =   1785
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "��� ������ ��������� ���������� ��� ������������� ����� � �������� �� ������� � ���� ������"
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   4950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����������, ������� ����� ������ � �� ���������� ���"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4950
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'TO DO - much to do here!
    WU.ClearAll (0)
    'Check record for existing else sorting datafile
    WU.NumID = Text1(0).text
    WU.UnitName = Text1(1).text
    EditMode = False
    If WU.NumID = "" Then
        Result = MsgBox("����������������� ����� ����� �� ������." + vbCrLf + "����������, ��������� ��� ����.", vbOKOnly, "����� ������ �� ������")
        Text1(0).SetFocus
        Exit Sub
    End If
    If WU.UnitName = "" Then
        Result = MsgBox("���������� ��� ����� �� �������." + vbCrLf + "����������, ��������� ��� ����.", vbOKOnly, "��� ����� �� �������")
        Text1(1).SetFocus
        Exit Sub
    End If
    '�������� ��������� � ����� ��������������
    If WU.CheckUnit(2, WU.NumID, WU.UnitName) Then
        Result = MsgBox("��� ������ ��� ����������! ������ �� ��" + vbCrLf + "������������� �� ���������?", vbYesNo + vbExclamation, "������ ��� ����������")
        If Result = vbYes Then
            EditMode = True
            Debug.Print EditMode & " Edit mode"
            If Not (WU.DecodeHistory(WU.ReadHistory(WU.NumID, 1))) Then
                Result = MsgBox("������ ��� ������� ��������� ���������� � �����" + vbCrLf + "�������������� ����������.", vbOKOnly, "������ ������ �����")
                EditMode = False
            End If
            Unload Me
            Exit Sub
        Else
            Result = MsgBox("����������, ������� ������ ������.", vbOKOnly + vbExclamation, "������ ��� ����������")
            Text1(1).text = ""
            Text1(0).text = ""
            Text1(0).SetFocus
            Debug.Print "Enter new info, please!"
            Exit Sub
        End If
    ElseIf WU.CheckUnit(0, WU.NumID) Then
        'UnitID already exist
        Result = MsgBox("���� � ���� ������� ��� ����������!" + vbCrLf + "����������, �������� ��� ����������", vbOKOnly, "����� ����� ��� ����������")
        Text1(0).SetFocus
        SendKeys "{Home}+{End}"
        Debug.Print "Block ID already exist!"
        Exit Sub
    ElseIf WU.CheckUnit(1, WU.UnitName) Then
        'Unit already exist
        Result = MsgBox("���� � ���� ������ ��� ����������!" + vbCrLf + "����������, �������� ��� ����������", vbOKOnly, "��� ����� ��� ����������")
        Text1(1).SetFocus
        SendKeys "{Home}+{End}"
        Debug.Print "Block name already exist!"
        Exit Sub
    End If
    Debug.Print "New record accepted"
    NewMode = True
    WU.ClearAll (1)
    Unload Me
End Sub

Private Sub Command2_Click()
    EditMode = False
    NewMode = False
    Unload Me
End Sub
