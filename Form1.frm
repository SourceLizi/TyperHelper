VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "���ٴ���"
   ClientHeight    =   4905
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4905
   ScaleWidth      =   5175
   StartUpPosition =   3  '����ȱʡ
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4650
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "����"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "��ʼ"
      TabPicture(0)   =   "Form1.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
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
      Tab(0).Control(7)=   "Timer3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Timer5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Ԥ��"
      TabPicture(1)   =   "Form1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command4"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(4)=   "Dir1"
      Tab(1).Control(5)=   "File1"
      Tab(1).Control(6)=   "Drive1"
      Tab(1).Control(7)=   "CommonDialog1"
      Tab(1).Control(8)=   "Label6"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "����"
      TabPicture(2)   =   "Form1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command7 
         Caption         =   "��ͣ����"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2520
         TabIndex        =   28
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   3240
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   2760
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   2280
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ȡ������"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2520
         TabIndex        =   25
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "��������"
         Height          =   735
         Left            =   -74760
         TabIndex        =   24
         Top             =   2520
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "�Զ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2520
            Top             =   240
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�ٶȿ���"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   3015
         Begin VB.CheckBox Check1 
            Caption         =   "�����ڴ���ʱ�����ٶ�"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   2295
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   1200
            TabIndex        =   21
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "Text2"
            BuddyDispid     =   196619
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
            Top             =   1080
            Width           =   600
         End
         Begin VB.OptionButton Option4 
            Caption         =   "ʹ�������ٶ�"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "ʹ�ü����ٶ�"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "������һ����"
            Height          =   255
            Left            =   1560
            TabIndex        =   22
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "ÿ��"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "ʹ�ÿ�ݼ�"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ʹ�õ���ʱ"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   2640
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "��ʼ����"
         Height          =   615
         Left            =   1080
         TabIndex        =   12
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   1800
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�ⲿ��"
         Height          =   495
         Left            =   -71280
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "·��"
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4575
         Begin VB.Label Label4 
            Caption         =   "�����ļ�·����"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   495
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   9
            Top             =   600
            Width           =   4335
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ȡ��"
         Height          =   615
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ȷ��"
         Height          =   615
         Left            =   -71280
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ԥ��"
         Height          =   615
         Left            =   -71280
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2640
         Width           =   4575
      End
      Begin VB.DirListBox Dir1 
         Height          =   1560
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.FileListBox File1 
         Height          =   1890
         Left            =   -73080
         Pattern         =   "*.txt"
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -74880
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "�ı� (*.txt)|*.txt"
      End
      Begin VB.Label Label6 
         Caption         =   "��ʾ��֧���ļ��ϷŹ���"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   2400
         Width           =   2895
      End
   End
   Begin VB.Label Label5 
      Caption         =   "   ���ٴ���V2.10"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1320
      Picture         =   "Form1.frx":091E
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3840
      Picture         =   "Form1.frx":11E8
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   4320
      Picture         =   "Form1.frx":14E0
      Top             =   0
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim time As Long
Dim start
Dim Result As String
Dim counttime
Dim status
Dim path, Lines As String
Dim mouse_x, mouse_y As Single
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Const Hwndx = -1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Sub Check2_Click()
Timer2.Enabled = Check2.Value
End Sub

Private Sub Command6_Click()
status = 6
StatusBar1.SimpleText = "������ȡ��"
End Sub

Private Sub Command7_Click()
Timer1.Enabled = False
Command5.Enabled = True
StatusBar1.SimpleText = "��������ͣ"
End Sub

Private Sub Form_Load()
Me.Hide
Dim XX As Long
XX = SetWindowPos(Me.hWnd, Hwndx, 0, 0, 0, 0, 3)
Me.Show
time = 1
SSTab1.Tab = 0
Sleep 100
Timer3.Enabled = True
End Sub

Private Sub readtxt(txtpath As String)
Open txtpath For Input As #1
Dim NextLine As String
Dim i As Integer
Do While Not EOF(1)
On Error Resume Next
Line Input #1, NextLine
Lines = Lines & NextLine & Chr(13) & Chr(10)
Loop
Close #1
End Sub
Private Sub Form1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(102, 0)
If Button = 1 Then
    Form1.Left = Form1.Left + (X - mouse_x)
    Form1.Top = Form1.Top + (Y - mouse_y)
End If
End Sub

Private Sub Form1_Resize()
       If Me.WindowState = vbMinimized Then Me.Hide
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
    If MsgBox("����ѡ���ļ��У�", , "����") = vbOK Then
    End If
End If
End Sub

Private Sub Command2_Click()
If Right(Dir1 + "\" + File1, 4) = ".txt" Then
    Label1.Caption = Dir1 + "\" + File1
    readtxt (Label1.Caption)
    SSTab1.Tab = 0
Else
    If MsgBox("����ѡ���ļ��У�", , "����") = vbOK Then
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

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 1 And LCase(Right(Data.Files(1), 4)) = ".txt" Then
    Label1.Caption = Data.Files(1)
    SSTab1.Tab = 0
End If
End Sub

Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim fext As String
'����ƹ��б����ǲ����ļ����Լ��Ƿ�ֻ�Ϸ�һ���ļ�
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'��ʾ���Է��µ�ͼ�꣬�Ǵ�С�Ӻŵ�����
    fext = LCase(Right(Data.Files(1), 4))
'�Ƿ�ָ�����ļ�����
    If fext = ".txt" Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
Else
    '������ʾ���ɷ��µ�ͼ�꣬��ԲȦ��б������
    Effect = vbDropEffectNone
End If
End Sub

Private Sub SSTab1_OLEDragDrop(Data As TabDlg.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 1 And LCase(Right(Data.Files(1), 4)) = ".txt" Then
    Label1.Caption = Data.Files(1)
    SSTab1.Tab = 0
End If
End Sub

Private Sub Command5_Click()
counttime = "6"
start = 1
If Label1.Caption <> "" Then
    Command6.Enabled = True
    Command3.Enabled = False
    Command5.Enabled = False
    Command7.Enabled = True
    Option1.Enabled = False
    Option2.Enabled = False
    Option4.Enabled = False
    Option3.Enabled = False
    Text2.Enabled = False
    UpDown1.Enabled = False
    Check1.Enabled = False
    If Option1.Value = True Then
        If Option3.Value = True Then
        status = 1
        End If
        If Option4.Value = True Then
        status = 2
        End If
    End If
    If Option2.Value = True Then
        If Option3.Value = True Then
        status = 3
        End If
        If Option4.Value = True Then
        status = 4
        End If
    End If
    Timer1.Enabled = True
Else
MsgBox "·������Ϊ�գ�"
End If
End Sub
Private Sub Dir1_Change()
 File1.path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1 = Left(Drive1, 2)
File1.path = Left(Drive1, 2)
End Sub

Private Sub Image1_Click()
Timer4.Enabled = True
If Check2.Enabled = True Then
    Timer2.Enabled = False
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(102, 0)
End Sub

Private Sub Image2_Click()
Me.Hide
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(101, 0)
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse_x = 0
 mouse_y = 0
 If Button = 1 Then
    mouse_x = X
    mouse_y = Y
 End If
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Form1.Left = Form1.Left + (X - mouse_x)
    Form1.Top = Form1.Top + (Y - mouse_y)
 End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse_x = 0
 mouse_y = 0
 If Button = 1 Then
    mouse_x = X
    mouse_y = Y
 End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 Form1.Left = Form1.Left + (X - mouse_x)
 Form1.Top = Form1.Top + (Y - mouse_y)
 End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    Text2.Enabled = False
    UpDown1.Enabled = False
    Check1.Enabled = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    Text2.Enabled = True
    UpDown1.Enabled = True
    Check1.Enabled = True
End If
End Sub

Private Sub SSTab1_OLEDragOver(Data As TabDlg.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim fext As String
'����ƹ��б����ǲ����ļ����Լ��Ƿ�ֻ�Ϸ�һ���ļ�
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'��ʾ���Է��µ�ͼ�꣬�Ǵ�С�Ӻŵ�����
    fext = LCase(Right(Data.Files(1), 4))
'�Ƿ�ָ�����ļ�����
    If fext = ".txt" Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
Else
'������ʾ���ɷ��µ�ͼ�꣬��ԲȦ��б������
    Effect = vbDropEffectNone
End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 1 And LCase(Right(Data.Files(1), 4)) = ".txt" Then
    readtxt (Data.Files(1))
    Text1.Text = Lines
End If
End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim fext As String
'����ƹ��б����ǲ����ļ����Լ��Ƿ�ֻ�Ϸ�һ���ļ�
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'��ʾ���Է��µ�ͼ�꣬�Ǵ�С�Ӻŵ�����
    fext = LCase(Right(Data.Files(1), 4))
'�Ƿ�ָ�����ļ�����
    If fext = ".txt" Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
Else
'������ʾ���ɷ��µ�ͼ�꣬��ԲȦ��б������
    Effect = vbDropEffectNone
End If
End Sub

Private Sub Text2_Change()
If Text1.Text > "2000" Then
    MsgBox "���ֳ�������"
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
    StatusBar1.SimpleText = "������ֿ�ʼ����" & counttime & "��"
    If counttime = "0" Then
    StatusBar1.SimpleText = "��ʼ����"
    SendKeys Lines
    status = 6
    End If
Case 2
    Timer1.Interval = 1000
    counttime = counttime - 1
    StatusBar1.SimpleText = "������ֿ�ʼ����" & counttime & "��"
    If counttime = "0" Then
    StatusBar1.SimpleText = "��ʼ����"
    If Check1.Value = Checked Then
    Sleep 1000
    StatusBar1.SimpleText = "��Pageup��Pagedown���ϵ����µ��ٶ�"
    status = 7
    Timer1.Interval = Text2.Text
    Else
    status = 5
    End If
    End If
Case 3
    StatusBar1.SimpleText = "�밴��Ctrl+Alt+F��������"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc("F")) Then
    StatusBar1.SimpleText = "��ʼ����"
    SendKeys Lines
    status = 6
    End If
Case 4
    StatusBar1.SimpleText = "�밴��Ctrl+Alt+F��������"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc("F")) Then
    StatusBar1.SimpleText = "��ʼ����"
    If Check1.Value = Checked Then
    Sleep 1000
    StatusBar1.SimpleText = "��Pageup��Pagedown���ϵ����µ��ٶ�"
    status = 7
    Timer1.Interval = Text2.Text
    Else
    status = 5
    End If
    End If
Case 5
    Timer1.Interval = Text2.Text
    Result = Mid(Lines, start, 1)
    start = start + 1
    SendKeys Result
    If Result = "" Then
    status = 6
    End If
Case 7
    If Timer1.Interval > "10" Then
    If GetAsyncKeyState(vbKeyPageDown) Then
    Timer1.Interval = Timer1.Interval - 10
    End If
    End If
    If GetAsyncKeyState(vbKeyPageUp) Then
    Timer1.Interval = Timer1.Interval + 10
    End If
    Result = Mid(Lines, start, 1)
    start = start + 1
    SendKeys Result
    If Result = "" Then
    status = 6
    End If
Case 6
    Command5.Enabled = True
    Command3.Enabled = True
    Command6.Enabled = False
    Command7.Enabled = False
    Option1.Enabled = True
    Option2.Enabled = True
    Option4.Enabled = True
    Option3.Enabled = True
    If Option4.Value = True Then
    Text2.Enabled = True
    UpDown1.Enabled = True
    Check1.Enabled = True
    End If
    Timer1.Enabled = False
    Sleep 3000
    StatusBar1.SimpleText = "����"
End Select
End Sub


Private Sub Timer2_Timer()
Dim P As POINTAPI, R As RECT
GetCursorPos P
GetWindowRect Me.hWnd, R
If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then
    Timer5.Enabled = True
    Timer4.Enabled = False
Else
    Timer3.Enabled = True
    Timer4.Enabled = False
    Timer5.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If time < 255 Then
    Me.Visible = True
    time = time + 10
Else
    Timer3.Enabled = False
End If
SetFormToAlpha Me.hWnd, time
End Sub

Private Sub Timer4_Timer()
If time > 0 Then
    time = time - 10
End If
SetFormToAlpha Me.hWnd, time
If time = 1 Then
    End
End If
End Sub

 Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse_x = 0
mouse_y = 0
 If Button = 1 Then
    mouse_x = X
    mouse_y = Y
 End If
 End Sub
 
 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
    Form1.Left = Form1.Left + (X - mouse_x)
    Form1.Top = Form1.Top + (Y - mouse_y)
 End If
Image2.Picture = LoadResPicture(102, 0)
 End Sub
 
Private Sub Timer5_Timer()
If time >= 5 Then
    time = time - 5
Else
    Timer5.Enabled = False
    Me.Visible = False
End If
SetFormToAlpha Me.hWnd, time
End Sub

