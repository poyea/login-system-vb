VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "��ܥD���x"
   ClientHeight    =   735
   ClientLeft      =   11715
   ClientTop       =   7995
   ClientWidth     =   4725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�L�n������"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Caption         =   "�Y�T�ݥΥ�����"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Top = 1350
Form3.Left = 1350
End Sub

Private Sub Command2_Click()
Form3.Top = 15195
Form3.Left = 19200
End Sub

Private Sub Command3_Click()
Text1.Visible = True
Text2.Visible = True
End Sub

