VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ע�����"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3480
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      Caption         =   "ע�����ɾ����ʱ�ļ�"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ע��"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "�������򿪺���ʾSysTray.ocx�ؼ�û��ע�ᣬ�밴���桰ע�ᡱ���ע��"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim temp() As Byte
temp = LoadResData(106, "CUSTOM")
Open App.Path & "\Register.bat" For Binary As #1
Put #1, 1, temp()
Close #1
temp = ""
temp = LoadResData(103, "CUSTOM")
Open App.Path & "\MSCOMCTL.OCX" For Binary As #1
Put #1, 1, temp()
Close #1
temp = ""
temp = LoadResData(104, "CUSTOM")
Open App.Path & "\TABCTL32.OCX" For Binary As #1
Put #1, 1, temp()
Close #1
temp = ""
temp = LoadResData(105, "CUSTOM")
Open App.Path & "\mscomct2.ocx" For Binary As #1
Put #1, 1, temp()
Close #1
Shell "cmd /c " & App.Path & "\Register.bat"
Command1.Enabled = False
MsgBox "ע�����,Ҫ�鿴������벻Ҫɾ����ʱ�ļ������������е�Register.bat�ļ�"
Command1.Enabled = True
If Check1.Value = Checked Then
If Dir(App.Path & "\Register.bat") <> "" Then
Kill App.Path & "\Register.bat"
End If
If Dir(App.Path & "\SysTray.ocx") <> "" Then
Kill App.Path & "\SysTray.ocx"
End If
If Dir(App.Path & "\MSCOMCTL.OCX") <> "" Then
Kill App.Path & "\MSCOMCTL.OCX"
End If
If Dir(App.Path & "\TABCTL32.OCX") <> "" Then
Kill App.Path & "\TABCTL32.OCX"
End If
If Dir(App.Path & "\mscomct2.ocx") <> "" Then
Kill App.Path & "\mscomct2.ocx"
End If
End If
End Sub
