VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '單線固定
   Caption         =   "請輸入驗證碼"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   960
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '透明
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   135
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "換一個"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Delay(ASecond As Integer)
    Dim t
    t = Timer
    Do
        DoEvents
    Loop Until (Int(Timer - t) = ASecond)
End Sub

Private Sub Command1_Click()
    If Label8.Caption = "輸入正確" And Label8.ForeColor = &H8000& Then
        Form1.Hide
        Me.Hide
        Form3.Show
    ElseIf Label9.Caption = 5 Then
        MsgBox "錯誤次數過多"
        End
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    Timer1.Interval = 10
    
End Sub

Private Sub Label7_Click()
    Timer1.Enabled = True
End Sub



Private Sub Text1_Change()
    Dim x As Long
    x = Label9.Caption
    ans = Label1.Caption & Label2.Caption & Label3.Caption & Label4.Caption & Label5.Caption & Label6.Caption
    If Text1.Text = ans Then
        Label8.Caption = "輸入正確"
        Label8.ForeColor = &H8000&
    Else
        Label8.Caption = "輸入錯誤"
        Label8.ForeColor = &HFF&
        Label9.Caption = x + 1
    End If
End Sub


Private Sub Timer1_Timer()
    
    Randomize
    
    For i = 1 To 150
        Label1.Caption = Int((9 * Rnd) + 1)
        Label1.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label1.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    
        Label2.Caption = Int((9 * Rnd) + 1)
        Label2.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label2.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
        
        Label3.Caption = Int((9 * Rnd) + 1)
        Label3.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label3.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
        
        Label4.Caption = Int((9 * Rnd) + 1)
        Label4.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label4.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
        
        Label5.Caption = Int((9 * Rnd) + 1)
        Label5.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label5.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
        
        Label6.Caption = Int((9 * Rnd) + 1)
        Label6.FontSize = Int((40 - 15 + 1) * Rnd + 15)
        Label6.ForeColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
        
        Delay (0.5)
        
    Next
    
    Timer1.Enabled = False
End Sub


