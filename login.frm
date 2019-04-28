VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "登入系統 - JOHN"
   ClientHeight    =   1620
   ClientLeft      =   11520
   ClientTop       =   7605
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "記住帳號"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  '暫止
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登入"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "密碼﹕"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "登入者名稱﹕"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
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

Private Sub Command1_Click()
    Dim success As Long
    If Check1.Value = 1 Then
        success = WritePrivateProfileString("LOGIN", "USERNAME", Text1.Text, "C:\SaveD.ini")
    End If
    If Text1.Text = "admin" Then
        ok1 = True
    Else
        ok1 = False
    End If
    If Text2.Text = "pass" Then
        ok2 = True
    Else
        ok2 = False
    End If
    If ok1 = False Or ok2 = False Then
        MsgBox "帳號或密碼錯誤", 0, "提示"
    ElseIf ok1 = True And ok2 = True Then
        MsgBox "成功登入!", 0, "提示"
        Form2.Show
    End If
End Sub

Private Sub Form_Load()
    Text2.PasswordChar = "*"
    Dim ret As Long
    Dim buff As String
    buff = String(255, 0)
    ret = GetPrivateProfileString("Myapp", "Box1N", "", buff, 256, "C:\SaveD.ini")
    Text1.Text = buff
End Sub





