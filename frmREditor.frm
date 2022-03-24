VERSION 5.00
Begin VB.Form frmREditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Результаты"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text20 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   46
      Top             =   2625
      Width           =   1380
   End
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   45
      Top             =   2940
      Width           =   1380
   End
   Begin VB.TextBox Text18 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   44
      Top             =   3255
      Width           =   1380
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   40
      Top             =   2625
      Width           =   1380
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   39
      Top             =   2940
      Width           =   1380
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   38
      Top             =   3255
      Width           =   1380
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   33
      Top             =   5250
      Width           =   1380
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   32
      Top             =   4935
      Width           =   1380
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   28
      Top             =   4635
      Width           =   1380
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   27
      Top             =   4320
      Width           =   1380
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   26
      Top             =   4020
      Width           =   1380
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   25
      Top             =   4620
      Width           =   1380
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   24
      Top             =   4305
      Width           =   1380
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   23
      Top             =   3990
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Из файла..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   105
      TabIndex        =   19
      Top             =   5880
      Width           =   1467
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   5985
      TabIndex        =   18
      Top             =   5880
      Width           =   1467
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4410
      TabIndex        =   17
      Top             =   5880
      Width           =   1467
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   13
      Top             =   1860
      Width           =   1380
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   12
      Top             =   1545
      Width           =   1380
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5985
      TabIndex        =   11
      Top             =   1230
      Width           =   1380
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   10
      Top             =   1890
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2205
      TabIndex        =   9
      Top             =   1575
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   810
      Width           =   750
   End
   Begin VB.Label Label27 
      Caption         =   "Мощность (power)"
      Height          =   225
      Left            =   3990
      TabIndex        =   49
      Top             =   2655
      Width           =   1800
   End
   Begin VB.Label Label26 
      Caption         =   "Форма сигнала (score)"
      Height          =   225
      Left            =   3990
      TabIndex        =   48
      Top             =   2970
      Width           =   1800
   End
   Begin VB.Label Label25 
      Caption         =   "Смещение (chirp rate)"
      Height          =   225
      Left            =   3990
      TabIndex        =   47
      Top             =   3285
      Width           =   1800
   End
   Begin VB.Label Label24 
      Caption         =   "Мощность (power)"
      Height          =   225
      Left            =   210
      TabIndex        =   43
      Top             =   2655
      Width           =   1800
   End
   Begin VB.Label Label23 
      Caption         =   "Форма сигнала (score)"
      Height          =   225
      Left            =   210
      TabIndex        =   42
      Top             =   2970
      Width           =   1800
   End
   Begin VB.Label Label22 
      Caption         =   "Смещение (chirp rate)"
      Height          =   225
      Left            =   210
      TabIndex        =   41
      Top             =   3285
      Width           =   1800
   End
   Begin VB.Label Label21 
      Caption         =   "Триплеты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3990
      TabIndex        =   37
      Top             =   2310
      Width           =   1485
   End
   Begin VB.Label Label20 
      Caption         =   "Импульсы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   36
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label19 
      Caption         =   "true_mean"
      Height          =   225
      Left            =   3990
      TabIndex        =   35
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Label Label18 
      Caption         =   "bin"
      Height          =   225
      Left            =   3990
      TabIndex        =   34
      Top             =   4965
      Width           =   1800
   End
   Begin VB.Label Label17 
      Caption         =   "sigma"
      Height          =   225
      Left            =   3990
      TabIndex        =   31
      Top             =   4665
      Width           =   1800
   End
   Begin VB.Label Label16 
      Caption         =   "fft_len"
      Height          =   225
      Left            =   3990
      TabIndex        =   30
      Top             =   4350
      Width           =   1800
   End
   Begin VB.Label Label15 
      Caption         =   "score"
      Height          =   225
      Left            =   3990
      TabIndex        =   29
      Top             =   4050
      Width           =   1800
   End
   Begin VB.Label Label14 
      Caption         =   "bin"
      Height          =   225
      Left            =   210
      TabIndex        =   22
      Top             =   4650
      Width           =   1800
   End
   Begin VB.Label Label13 
      Caption         =   "fft_len"
      Height          =   225
      Left            =   210
      TabIndex        =   21
      Top             =   4335
      Width           =   1800
   End
   Begin VB.Label Label12 
      Caption         =   "score"
      Height          =   225
      Left            =   210
      TabIndex        =   20
      Top             =   4020
      Width           =   1800
   End
   Begin VB.Line Line6 
      X1              =   105
      X2              =   7455
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line5 
      X1              =   3780
      X2              =   3780
      Y1              =   3990
      Y2              =   5670
   End
   Begin VB.Label Label11 
      Caption         =   "Гауссианы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3990
      TabIndex        =   16
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label10 
      Caption         =   "Пики"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   15
      Top             =   1260
      Width           =   540
   End
   Begin VB.Line Line4 
      X1              =   4935
      X2              =   7455
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Line Line3 
      X1              =   2310
      X2              =   105
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Label Label9 
      Caption         =   "Необязательные параметры"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2310
      TabIndex        =   14
      Top             =   3675
      Width           =   2640
   End
   Begin VB.Label Label8 
      Caption         =   "Смещение (chirp rate)"
      Height          =   225
      Left            =   3990
      TabIndex        =   7
      Top             =   1890
      Width           =   1800
   End
   Begin VB.Label Label7 
      Caption         =   "Форма сигнала (fit)"
      Height          =   225
      Left            =   3990
      TabIndex        =   6
      Top             =   1575
      Width           =   1800
   End
   Begin VB.Label Label6 
      Caption         =   "Мощность (power)"
      Height          =   225
      Left            =   3990
      TabIndex        =   5
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Смещение (chirp rate)"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   1890
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Мощность (power)"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   1575
      Width           =   1800
   End
   Begin VB.Line Line2 
      X1              =   3780
      X2              =   105
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   3780
      Y1              =   735
      Y2              =   3570
   End
   Begin VB.Label Label3 
      Caption         =   "Номер рабочего блока"
      Height          =   225
      Left            =   525
      TabIndex        =   2
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Если какая-либо величина неизвестна, оставьте соответствующее поле НЕЗАПОЛНЕННЫМ."
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   315
      Width           =   7260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста, среди необязательных параметров укажите ТОЛЬКО ТЕ, которые Вам известны."
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7470
   End
End
Attribute VB_Name = "frmREditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Key2 As Boolean
Dim Key3 As Boolean
Dim Key4 As Boolean
Dim Key5 As Boolean
Dim Key6 As Boolean
Dim Key15 As Boolean
Dim Key16 As Boolean
Dim Key17 As Boolean
Dim Key18 As Boolean
Dim Key19 As Boolean
Dim Key20 As Boolean

Private Sub Command1_Click()
    If Not (Key2 And Key3 And Key4 And Key5 And Key6 And Key15 And Key16 And Key17 And Key18 And Key19 And Key20) Then
        Call RaiseErrMsg(1209, 1101)
        If Not (Key2) Then
            Text2.SetFocus
        ElseIf Not (Key3) Then
            Text3.SetFocus
        ElseIf Not (Key4) Then
            Text4.SetFocus
        ElseIf Not (Key5) Then
            Text5.SetFocus
        ElseIf Not (Key6) Then
            Text6.SetFocus
        ElseIf Not (Key15) Then
            Text15.SetFocus
        ElseIf Not (Key16) Then
            Text16.SetFocus
        ElseIf Not (Key17) Then
            Text17.SetFocus
        ElseIf Not (Key18) Then
            Text18.SetFocus
        ElseIf Not (Key19) Then
            Text19.SetFocus
        ElseIf Not (Key20) Then
            Text20.SetFocus
        End If
    Else
        ScanEntry
        If State.EncodeHistory Then
            If State.WriteHistory(1, EditID) Then
            End If
        End If
        frmHistory.Command3.Enabled = False
        Unload Me
    End If
End Sub

Private Sub ScanEntry()
    State.bs_power = Val(Text2.text)
    State.bs_rate = Val(Text3.text)
    State.bg_power = Val(Text4.text)
    State.bg_chisq = Val(Text5.text)
    State.bg_rate = Val(Text6.text)
    If Text7.text = "" Then
        State.bs_score = 0
    Else
        State.bs_score = Val(Text7.text)
    End If
    If Text8.text = "" Then
        State.bs_fft_len = 654321
    Else
        State.bs_fft_len = Val(Text8.text)
    End If
    If Text9.text = "" Then
        State.bs_bin = 0
    Else
        State.bs_bin = Val(Text9.text)
    End If
    If Text10.text = "" Then
        State.bg_score = 0
    Else
        State.bg_score = Val(Text10.text)
    End If
    If Text11.text = "" Then
        State.bg_fft_len = 0
    Else
        State.bg_fft_len = Val(Text11.text)
    End If
    If Text12.text = "" Then
        State.bg_sigma = 654321
    Else
        State.bg_sigma = Val(Text12.text)
    End If
    If Text13.text = "" Then
        State.bg_bin = 0
    Else
        State.bg_bin = Val(Text13.text)
    End If
    If Text14.text = "" Then
        State.bg_true_mean = 0
    Else
        State.bg_true_mean = Val(Text14.text)
    End If
    State.bp_power = Val(Text17.text)
    State.bp_chirp_rate = Val(Text15.text)
    State.bp_score = Val(Text16.text)
    State.bt_power = Val(Text20.text)
    State.bt_chirp_rate = Val(Text18.text)
    State.bt_score = Val(Text19.text)
End Sub

Private Sub Command2_Click()
    frmHistory.Command3.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Caption = "Редактор результатов"
    Text1.text = Str(EditID)
    If State.ReadHistory(EditID) = 0 Then
        If State.DecodeHistory Then
            If State.Status = 1 Then
            'Флаг игнорирования информации опущен - обрабатываем ее.
                Text2.text = Str(State.bs_power)
                Text3.text = Str(State.bs_rate)
                Text4.text = Str(State.bg_power)
                Text5.text = Str(State.bg_chisq)
                Text6.text = Str(State.bg_rate)
                Text15.text = Str(State.bp_chirp_rate)
                Text16.text = Str(State.bp_score)
                Text17.text = Str(State.bp_power)
                Text18.text = Str(State.bt_chirp_rate)
                Text19.text = Str(State.bt_score)
                Text20.text = Str(State.bt_power)
                If (State.bg_fft_len = 654321) Or (State.bs_fft_len = 654321) Then
                    Text7.text = ""
                    Text8.text = ""
                    Text9.text = ""
                    Text10.text = ""
                    Text11.text = ""
                    Text12.text = ""
                    Text13.text = ""
                    Text14.text = ""
                Else
                    Text7.text = Str(State.bs_score)
                    Text10.text = Str(State.bg_score)
                    Text8.text = Str(State.bs_fft_len)
                    Text11.text = Str(State.bg_fft_len)
                    Text12.text = Str(State.bg_sigma)
                    Text9.text = Str(State.bs_bin)
                    Text13.text = Str(State.bg_bin)
                    Text14.text = Str(State.bg_true_mean)
                End If
            Else
                Text2.text = ""
                Text3.text = ""
                Text7.text = ""
                Text8.text = ""
                Text4.text = ""
                Text6.text = ""
                Text10.text = ""
                Text11.text = ""
                Text12.text = ""
                Text9.text = ""
                Text5.text = ""
                Text13.text = ""
                Text14.text = ""
                Text15.text = ""
                Text16.text = ""
                Text17.text = ""
                Text18.text = ""
                Text19.text = ""
                Text20.text = ""
            End If
        End If
    End If
    Key2 = False
    Key3 = False
    Key4 = False
    Key5 = False
    Key6 = False
    Key15 = False
    Key16 = False
    Key17 = False
    Key18 = False
    Key19 = False
    Key20 = False
End Sub

Private Sub Text2_LostFocus()
    If Text2.text <> "" Then
        Key2 = True
    Else
        Key2 = False
    End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.text <> "" Then
        Key3 = True
    Else
        Key3 = False
    End If
End Sub

Private Sub Text4_LostFocus()
    If Text4.text <> "" Then
        Key4 = True
    Else
        Key4 = False
    End If
End Sub

Private Sub Text5_LostFocus()
    If Text5.text <> "" Then
        Key5 = True
    Else
        Key5 = False
    End If
End Sub

Private Sub Text6_LostFocus()
    If Text6.text <> "" Then
        Key6 = True
    Else
        Key6 = False
    End If
End Sub

Private Sub Text15_LostFocus()
    If Text15.text <> "" Then
        Key15 = True
    Else
        Key15 = False
    End If
End Sub

Private Sub Text16_LostFocus()
    If Text16.text <> "" Then
        Key16 = True
    Else
        Key16 = False
    End If
End Sub

Private Sub Text17_LostFocus()
    If Text17.text <> "" Then
        Key17 = True
    Else
        Key17 = False
    End If
End Sub

Private Sub Text18_LostFocus()
    If Text18.text <> "" Then
        Key18 = True
    Else
        Key18 = False
    End If
End Sub

Private Sub Text19_LostFocus()
    If Text19.text <> "" Then
        Key19 = True
    Else
        Key19 = False
    End If
End Sub

Private Sub Text20_LostFocus()
    If Text20.text <> "" Then
        Key20 = True
    Else
        Key20 = False
    End If
End Sub
