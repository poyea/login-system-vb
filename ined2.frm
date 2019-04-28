VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   2625
   ClientLeft      =   12660
   ClientTop       =   6645
   ClientWidth     =   3900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "計時"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢剩餘時間"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "還有"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "00:00:00"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Private Sub Timer1_Timer()
    Dim de As Date
    de = Now
    Form3.Caption = "時間主控台 - " & de
    Label3.Caption = de
End Sub


Private Sub Form_Load()
    Dim ret As Long
    Dim buff As String
    buff = String(255, 0)
    ret = GetPrivateProfileString("InED", "Box2N", "", buff, 256, "C:\ined.ini")
    Text1.Text = buff
    If buff = "00:00:00" Then
        Label2.Caption = "時間已用完"
        Label1.Visible = False
    End If
    MsgBox "請按使用剩餘時間/計時按鈕", 32, "注意"
End Sub


Private Sub Command2_Click()
    Dim a As Long, y As Double
    a = WritePrivateProfileString("InED", "Box2N", Label1.Caption, "C:\ined.ini")
    y = WritePrivateProfileString("InED", "Box3N", Label3.Caption, "C:\ined.ini")
    MsgBox "管理員會查看記錄，一但發現作弊，將會被凍結帳號", 32, "注意"
    End
End Sub


Private Sub Command3_Click()
    Me.Top = 15195
    Me.Left = 19200
    Form4.Show
    Command2.Visible = True
    Command1.Visible = False
    Command3.Visible = False
End Sub

Private Sub Command1_Click()
    Label1.Visible = True
    Label2.Visible = False
    Form4.Show
    Me.Top = 15195
    Me.Left = 19200
    Command1.Visible = False
    Command2.Visible = True
    Command3.Visible = False
    Dim BeforeTime As Single
    Static InDo As Boolean
    InDo = Not InDo
    If InDo Then
        If Not IsDate(Text1) Then
        Else
            BeforeTime = Timer
            Text1.Enabled = False
            Label1 = Format(Text1, "hh:mm:ss")
            Do While Label1 <> "00:00:00"
                If Not InDo Then Exit Do
                If Timer >= BeforeTime + 1 Then
                    BeforeTime = Timer
                    Label1 = Format(DateAdd("s", -1, Label1), "hh:mm:ss")
                End If
                DoEvents
            Loop
        End If
        InDo = False
    End If
    If Label1.Caption = "00:00:00" Then
        Me.Top = 1350
        Me.Left = 1350
        Dim k As Long
        k = MsgBox("剩餘時間已經完結，請與管理員聯絡購買", 4096, "特別注意")
        If k = 1 Then
            Dim r As Long
            r = WritePrivateProfileString("InED", "Box2N", "00:00:00", "C:\ined.ini")
            'Shell "shutdown -s -f -t 120 "
            End
        End If
    End If
End Sub




