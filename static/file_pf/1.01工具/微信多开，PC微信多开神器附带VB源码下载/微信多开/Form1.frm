VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "΢�Ŷ࿪"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   1935
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "�������̲���"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�࿪"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "2"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "�ᰮ�ƽ� LaoDan"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "΢�Ŷ࿪��"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '��������API�е�ShellExecute��

Private Sub Command1_Click()
Dim ·�� As String
·�� = GetStringValue(HKEY_CURRENT_USER, "Software\Tencent\WeChat", "InstallPath")
ChDir ·��
For i = 1 To Text1.Text
ShellExecute Me.hwnd, "open", ·�� & "\WeChat.exe", vbNullString, vbNullString, vbNormalFocus
Next
End Sub

Private Sub Command2_Click()
Shell "cmd.exe /c taskkill.exe /f /im WeChat.exe"
End Sub
