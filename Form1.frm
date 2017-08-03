VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "快速打字"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5175
   StartUpPosition =   3  '窗口缺省
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4125
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "就绪"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7011
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "开始"
      TabPicture(0)   =   "Form1.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Timer1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "预览"
      TabPicture(1)   =   "Form1.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "CommonDialog1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Drive1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "File1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Dir1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "速度控制"
      TabPicture(2)   =   "Form1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton Command6 
         Caption         =   "取消打字"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -72480
         TabIndex        =   23
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "速度控制"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   3015
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   1200
            TabIndex        =   21
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "Text2"
            BuddyDispid     =   196631
            OrigLeft        =   1440
            OrigTop         =   960
            OrigRight       =   1695
            OrigBottom      =   1215
            Max             =   2000
            Min             =   10
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   270
            Left            =   600
            TabIndex        =   20
            Text            =   "100"
            Top             =   1200
            Width           =   600
         End
         Begin VB.OptionButton Option4 
            Caption         =   "使用限制速度"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "使用极限速度"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "毫秒打打一个字"
            Height          =   255
            Left            =   1440
            TabIndex        =   22
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "每隔"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1200
            Width           =   375
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "使用快捷键"
         Height          =   255
         Left            =   -74040
         TabIndex        =   15
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "使用倒计时"
         Height          =   255
         Left            =   -74040
         TabIndex        =   14
         Top             =   2640
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "开始打字"
         Height          =   615
         Left            =   -73920
         TabIndex        =   12
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -74760
         Top             =   1800
      End
      Begin VB.CommandButton Command4 
         Caption         =   "外部打开"
         Height          =   495
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "路径"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   4575
         Begin VB.Label Label4 
            Caption         =   "打字文件路径："
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   4335
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   615
         Left            =   -71520
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   615
         Left            =   3720
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "预览"
         Height          =   615
         Left            =   3720
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2640
         Width           =   4575
      End
      Begin VB.DirListBox Dir1 
         Height          =   1560
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.FileListBox File1 
         Height          =   1890
         Left            =   1920
         Pattern         =   "*.txt"
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "文本 (*.txt)|*.txt"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim start
Dim result As String
Dim counttime
Dim status
Dim path As String
Dim Lines As String
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Const Hwndx = -1
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Command6_Click()
status = 6
End Sub

Private Sub Form_Load()
Dim XX As Long
XX = SetWindowPos(Me.Hwnd, Hwndx, 0, 0, 0, 0, 3)
End Sub
Private Sub readtxt()
Open path For Input As #1
Dim NextLine As String
Dim i As Integer
Do While Not EOF(1)
On Error Resume Next
Line Input #1, NextLine
Lines = Lines & NextLine & Chr(13) & Chr(10)
Loop
Close #1
End Sub
Private Sub Command1_Click()
If Right(Dir1 + "\" + File1, 4) = ".txt" Then
Open Dir1 + "\" + File1 For Input As #1
Dim Lines As String
Dim NextLine As String
Dim i As Integer
Do While Not EOF(1)
On Error Resume Next
Line Input #1, NextLine
Lines = Lines & NextLine & Chr(13) & Chr(10)
Loop
Close #1
Text1.Text = Lines
Else
If MsgBox("不能选中文件夹！", , "出错") = vbOK Then
End If
End If
End Sub
Private Sub Command2_Click()
If Right(Dir1 + "\" + File1, 4) = ".txt" Then
Label1.Caption = Dir1 + "\" + File1
path = Label1.Caption
readtxt
SSTab1.Tab = 0
Else
If MsgBox("不能选中文件夹！", , "出错") = vbOK Then
End If
End If
End Sub
Private Sub Command3_Click()
path = ""
Lines = ""
SSTab1.Tab = 1
Label1.Caption = ""
End Sub
Private Sub Command4_Click()
CommonDialog1.ShowOpen
Label1.Caption = CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then
SSTab1.Tab = 0
End If
End Sub
Private Sub Command5_Click()
counttime = "6"
start = 1
If Label1.Caption <> "" Then
    Timer1.Enabled = True
    Command6.Enabled = True
    Command3.Enabled = False
    Command5.Enabled = False
    If Option1.Value = True Then
        If Option3.Value = True Then
        status = 1
        End If
        If Option4.Value = True Then
           Timer1.Interval = Text2.Text
        status = 2
        End If
    End If
    If Option2.Value = True Then
        If Option3.Value = True Then
        status = 3
        End If
        If Option4.Value = True Then
           Timer1.Interval = Text2.Text
        status = 4
        End If
    End If
Else
MsgBox "路径不能为空！"
End If
End Sub
Private Sub Dir1_Change()
 File1.path = Dir1
End Sub
Private Sub Drive1_Change()
Dir1 = Left(Drive1, 2)
File1.path = Left(Drive1, 2)
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text2.Enabled = False
UpDown1.Enabled = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Text2.Enabled = True
UpDown1.Enabled = True
End If
End Sub

Private Sub Text2_Change()
If Text1.Text > "2000" Then
MsgBox "数字超过上限"
Text1.Text = "100"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Select Case status
Case 1
    Timer1.Interval = 1000
    counttime = counttime - 1
    StatusBar1.SimpleText = "距离打字开始还有" & counttime & "秒"
    If counttime = "0" Then
    SendKeys Lines
    status = 6
    End If
Case 2
    Timer1.Interval = 1000
    counttime = counttime - 1
    StatusBar1.SimpleText = "距离打字开始还有" & counttime & "秒"
    If counttime = "0" Then
    StatusBar1.SimpleText = "开始打字"
    status = 5
    End If
Case 3
    StatusBar1.SimpleText = "请按下Ctrl+Alt+F启动打字"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc("F")) Then
    StatusBar1.SimpleText = "开始打字"
    SendKeys Lines
    status = 6
    End If
Case 4
    StatusBar1.SimpleText = "请按下Ctrl+Alt+F启动打字"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc("F")) Then
    StatusBar1.SimpleText = "开始打字"
    status = 5
    End If
Case 5
    result = Mid(Lines, start, 1)
    start = start + 1
    SendKeys result
    If result = "" Then
    status = 6
    End If
Case 6
    Command5.Enabled = True
    Command3.Enabled = True
    Command6.Enabled = False
    Timer1.Enabled = False
    Sleep 3000
    StatusBar1.SimpleText = "就绪"
End Select
End Sub
