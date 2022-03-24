VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Default         =   -1  'True
      Height          =   345
      Left            =   6615
      TabIndex        =   0
      Top             =   5460
      Width           =   1467
   End
   Begin VB.Label Label15 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   32
      Top             =   4620
      Width           =   3900
   End
   Begin VB.Label Label1 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   31
      Top             =   210
      Width           =   3900
   End
   Begin VB.Label Label16 
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
      Index           =   1
      Left            =   4200
      TabIndex        =   30
      Top             =   5040
      Width           =   3900
   End
   Begin VB.Label Label14 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   29
      Top             =   4305
      Width           =   3900
   End
   Begin VB.Label Label13 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   28
      Top             =   3990
      Width           =   3900
   End
   Begin VB.Label Label12 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   27
      Top             =   3675
      Width           =   3900
   End
   Begin VB.Label Label11 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   26
      Top             =   3360
      Width           =   3900
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   25
      Top             =   3045
      Width           =   3900
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   24
      Top             =   2730
      Width           =   3900
   End
   Begin VB.Label Label8 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   23
      Top             =   2415
      Width           =   3900
   End
   Begin VB.Label Label2 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   22
      Top             =   525
      Width           =   3900
   End
   Begin VB.Label Label3 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   21
      Top             =   840
      Width           =   3900
   End
   Begin VB.Label Label4 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   20
      Top             =   1155
      Width           =   3900
   End
   Begin VB.Label Label5 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   19
      Top             =   1470
      Width           =   3900
   End
   Begin VB.Label Label6 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   18
      Top             =   1785
      Width           =   3900
   End
   Begin VB.Label Label7 
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   17
      Top             =   2100
      Width           =   3900
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Количество отправленных результатов (N Results)"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   16
      Top             =   4620
      Width           =   3900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ID"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   15
      Top             =   210
      Width           =   3900
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Суммарное процессорное время"
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
      Index           =   0
      Left            =   105
      TabIndex        =   14
      Top             =   5040
      Width           =   3900
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "N WUs"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   13
      Top             =   4305
      Width           =   3900
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Дата последней отправки результатов"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   12
      Top             =   3990
      Width           =   3900
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Last WU Time"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   3675
      Width           =   3900
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Дата регистрации"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   3360
      Width           =   3900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Venue"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   3045
      Width           =   3900
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Показывать e-mail"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   2730
      Width           =   3900
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Показывать имя"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   2415
      Width           =   3900
   End
   Begin VB.Line Line1 
      X1              =   4095
      X2              =   4095
      Y1              =   5355
      Y2              =   105
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Почтовый индекс"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   2100
      Width           =   3900
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Страна"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   1785
      Width           =   3900
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "URL"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   1470
      Width           =   3900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Имя участника"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   1155
      Width           =   3900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   3900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ключ"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   3900
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim TMPstr As String
Dim TMPDbl As Double
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Caption = "Информация о пользователе"
    bResult = UserInfo.DecodeInfo(UserInfo.ReadFile(0))
    Label1(1).Caption = UserInfo.ID
    Label2(1).Caption = UserInfo.Key
    Label3(1).Caption = UserInfo.email
    Label4(1).Caption = UserInfo.UserName
    Label5(1).Caption = UserInfo.URL
    Label6(1).Caption = UserInfo.country
    Label7(1).Caption = UserInfo.PostalCode
    Label8(1).Caption = UserInfo.ShowName
    Label9(1).Caption = UserInfo.ShowEmail
    Label10(1).Caption = UserInfo.Venue
    Label11(1).Caption = UserInfo.Register
    Label12(1).Caption = UserInfo.LastWU
    Label13(1).Caption = UserInfo.LastResult
    Label14(1).Caption = UserInfo.Nwus
    Label15(1).Caption = UserInfo.NResults
    TMPstr = UserInfo.totalCPU
    TMPDbl = CDbl(Val(TMPstr))
    TMPstr = DecodeTime(TMPDbl, True)
    Label16(1).Caption = TMPstr
End Sub
Private Sub Form_Click()
   Dim i, OldFontSize   ' Declare variables.
   ''Width = 8640: Height = 5760   ' Set form size in twips.
   Move 100, 100  ' Move form origin.
   AutoRedraw = -1   ' Turn on AutoRedraw.
   OldFontSize = FontSize   ' Save old font size.
   BackColor = QBColor(7)   ' Set background to gray.
   Scale (0, 110)-(130, 0)   ' Set custom coordinate system.
   For i = 100 To 10 Step -10
      Line (0, i)-(2, i)   ' Draw scale marks every 10 units.
      CurrentY = CurrentY + 1.5   ' Move cursor position.
      Print i   ' Print scale mark value on left.
      Line (ScaleWidth - 2, i)-(ScaleWidth, i)
      CurrentY = CurrentY + 1.5   ' Move cursor position.
      CurrentX = ScaleWidth - 9
      Print i   ' Print scale mark value on right.
   Next i
   ' Draw bar chart.
   Line (10, 0)-(20, 45), RGB(0, 0, 255), BF   ' First blue bar.
   Line (20, 0)-(30, 55), RGB(255, 0, 0), BF   ' First red bar.
   Line (40, 0)-(50, 40), RGB(0, 0, 255), BF
   Line (50, 0)-(60, 25), RGB(255, 0, 0), BF
   Line (70, 0)-(80, 35), RGB(0, 0, 255), BF
   Line (80, 0)-(90, 60), RGB(255, 0, 0), BF
   Line (100, 0)-(110, 75), RGB(0, 0, 255), BF
   Line (110, 0)-(120, 90), RGB(255, 0, 0), BF
   CurrentX = 18: CurrentY = 100   ' Move cursor position.
   FontSize = 14   ' Enlarge font for title.
   Print "Widget Quarterly Sales"   ' Print title.
   FontSize = OldFontSize   ' Restore font size.
   CurrentX = 27: CurrentY = 93   ' Move cursor position.
   Print "Planned Vs. Actual"   ' Print subtitle.
   Line (29, 86)-(34, 88), RGB(0, 0, 255), BF   ' Print legend.
   Line (43, 86)-(49, 88), RGB(255, 0, 0), BF
End Sub
