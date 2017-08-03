VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "快速打字"
   ClientHeight    =   5490
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5490
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   5235
      Width           =   5700
      _ExtentX        =   10054
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
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "开始"
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
      Tab(0).Control(5)=   "Command6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Timer3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Timer4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Timer6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tmrCheck"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tmrMove"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Option2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Timer7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Timer8"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "预览"
      TabPicture(1)   =   "Form1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Dir1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "File1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Drive1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CommonDialog1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "输出设置"
      TabPicture(2)   =   "Form1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "程序控制"
      TabPicture(3)   =   "Form1.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).ControlCount=   2
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   2640
      End
      Begin VB.Frame Frame4 
         Caption         =   "自定义快捷键"
         Height          =   975
         Left            =   -74760
         TabIndex        =   34
         Top             =   1920
         Width           =   3135
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   1920
            TabIndex        =   35
            Text            =   "F"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "快捷键：Ctrl+Alt+"
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   3120
      End
      Begin VB.CheckBox Check7 
         Caption         =   "自动保存文章"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Caption         =   "自动模式"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "定向发送"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   28
         Top             =   2520
         Width           =   3135
         Begin VB.CheckBox Check5 
            Caption         =   "同时暂停打字"
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox Check4 
            Caption         =   "失去焦点时自动转到定向发送"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.CheckBox Check3 
            Caption         =   "定向发送"
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "窗体属性"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   3135
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2520
            Top             =   240
         End
         Begin VB.CheckBox Check2 
            Caption         =   "自动隐藏"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "使用快捷键"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Timer tmrMove 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   720
         Top             =   4020
      End
      Begin VB.Timer tmrCheck 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   3600
      End
      Begin VB.Timer Timer6 
         Interval        =   100
         Left            =   240
         Top             =   4020
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   3540
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   3060
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   240
         Top             =   2580
      End
      Begin VB.CommandButton Command6 
         Caption         =   "取消打字"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2760
         TabIndex        =   22
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "速度控制"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   3135
         Begin VB.CheckBox Check1 
            Caption         =   "允许在打字时控制速度"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   2295
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   1200
            TabIndex        =   19
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "Text2"
            BuddyDispid     =   196634
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
            TabIndex        =   18
            Text            =   "100"
            Top             =   1080
            Width           =   600
         End
         Begin VB.OptionButton Option4 
            Caption         =   "使用限制速度"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "使用极限速度"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "毫秒打打一个字"
            Height          =   255
            Left            =   1560
            TabIndex        =   20
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "每隔"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "使用倒计时"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   2760
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "开始打字"
         Height          =   615
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   2100
      End
      Begin VB.CommandButton Command4 
         Caption         =   "外部打开"
         Height          =   975
         Left            =   -70800
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "路径"
         Height          =   1095
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   4935
         Begin VB.Label Label4 
            Caption         =   "打字文件路径："
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   495
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   8
            Top             =   480
            Width           =   4695
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   615
         Left            =   3840
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   855
         Left            =   -70800
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2760
         Width           =   5175
      End
      Begin VB.DirListBox Dir1 
         Height          =   1560
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   2070
         Left            =   -72600
         Pattern         =   "*.txt"
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -74760
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "文本 (*.txt)|*.txt"
      End
      Begin VB.Label Label6 
         Caption         =   "提示：支持文件拖放功能"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   2520
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      Caption         =   "   快速打字V2.31"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1320
      Picture         =   "Form1.frx":093A
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4320
      Picture         =   "Form1.frx":1204
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   4800
      Picture         =   "Form1.frx":14FC
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
Dim myhwnd As Long
Dim time As Long
Dim start, length As Integer
Dim path As String
Dim Result As String
Dim counttime As Integer
Dim status As Integer
Dim Lines As String
Dim mouse_x As Single
Dim mouse_y As Single
Private Const WM_SETTEXT = &HC
Private Const WM_CHAR = &H102
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const Hwndx = -1
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
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
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_EXITSIZEMOVE = &H232
Private Const WM_MOVING = &H216
Private WithEvents c_Subclass   As iSubClass
Attribute c_Subclass.VB_VarHelpID = -1

Private Const SIZE_SHOW         As Long = 60    '隐藏后留出来的宽度或高度,单位缇
Private Const SHOWHIDE_SPEED    As Long = 170    '(自动显示隐藏速度，单位缇)
'显示标识
'0  自动隐藏
'1  自动显示
Private m_ShowFlag              As Long
'显示方向
'0  向左
'1  向右
'2  向上
Private m_ShowOrient            As Long
'显示速度
Private m_ShowSpeed             As Long
'是否已经启动自动隐藏(为了防止WM_MOVING调整窗口位置)
Private m_MoveEnabled           As Boolean

'//下面是把窗口移动Top=0且Left=0或Right=Screen.Width的时候让窗口高度=屏幕高度
'是否自动调整了大小
Private m_AutoSize              As Boolean
Private m_OldHeight             As Long


Const WM_SYSCOMMAND = &H112
Const SC_MOVE = &HF012
Private Const HTCAPTION = 2

Private Sub Check3_Click()
Select Case Check3.Value
Case Checked
Option4.Value = True
Option3.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Case Unchecked
Option3.Enabled = True
Check6.Enabled = False
Check7.Enabled = False
Option1.Enabled = True
Option2.Enabled = True
End Select
End Sub

Private Sub Check6_Click()
Select Case Check6.Value
Case Checked
Command5.Enabled = Not Command5.Enabled
Command3.Enabled = Not Command3.Enabled
Option1.Enabled = Not Option1.Enabled
Option2.Enabled = Not Option2.Enabled
StatusBar1.SimpleText = "检测文章中"
Timer7.Enabled = True
Check7.Enabled = True
Case Unchecked
Command5.Enabled = Not Command5.Enabled
Command3.Enabled = Not Command3.Enabled
Option1.Enabled = Not Option1.Enabled
Option2.Enabled = Not Option2.Enabled
StatusBar1.SimpleText = "检测文章中"
Timer7.Enabled = False
Check7.Enabled = False
End Select
End Sub

Private Sub File1_Click()
If File1.FileName <> "" Then
Readtxt (Replace(Dir1.path & "\" & File1.FileName, "\\", "\"))
Text1.Text = Lines
Lines = ""
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ReleaseCapture
     SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub c_Subclass_GetWindowMessage(Result As Long, ByVal cHwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case Message
        Case WM_NCLBUTTONDOWN
            Const HTCAPTION = 2
            If wParam = HTCAPTION Then
                '点击标题栏让所有Timer停止工作
                m_MoveEnabled = True
                tmrCheck.Enabled = False
                tmrMove.Enabled = False
            End If
            
        Case WM_MOVING
            If m_MoveEnabled = False Then Exit Sub
            '这里仅仅是为了不让窗口移出屏幕，可以忽略
            Dim rcMov   As RECT
            Dim rcWnd   As RECT
            Dim lScrW   As Long
            '获取窗口矩形
            Call GetWindowRect(cHwnd, rcWnd)
            '//屏幕宽度
            lScrW = Screen.Width / Screen.TwipsPerPixelX
            '获取移动目标位置矩形
            Call CopyMemory(rcMov, ByVal lParam, Len(rcMov))
            With rcMov
                If .Left < 0 Then
                    .Left = 0
                    .Right = rcWnd.Right - rcWnd.Left
                End If
                If .Top < 0 Then
                    .Top = 0
                    .Bottom = rcWnd.Bottom - rcWnd.Top
                End If
                If .Right > lScrW Then
                    .Left = lScrW - (rcWnd.Right - rcWnd.Left)
                    .Right = .Left + (rcWnd.Right - rcWnd.Left)
                End If
            End With
            '//如果窗口的靠在右上角或左上角，则把高度设置为屏幕高度
            If rcMov.Top = 0 And (rcMov.Left = 0 Or rcMov.Right = Screen.Width / Screen.TwipsPerPixelX) Then
                'If m_AutoSize = False Then
                    'm_AutoSize = True
                    '保存旧的高度
                    'm_OldHeight = rcMov.Bottom - rcMov.Top
                    'rcMov.Bottom = Screen.Height / Screen.TwipsPerPixelY
                'End If
            Else
                If m_AutoSize Then
                    m_AutoSize = False
                    '设置旧的高度
                    rcMov.Bottom = rcMov.Top + m_OldHeight
                End If
            End If
            Call CopyMemory(ByVal lParam, rcMov, Len(rcMov))
            
        Case WM_EXITSIZEMOVE
            m_MoveEnabled = False
            Call GetWindowRect(cHwnd, rcWnd)
            If rcWnd.Left <= 0 Or rcWnd.Top <= 0 Or _
                rcWnd.Right >= Screen.Width / Screen.TwipsPerPixelX Then
                '如果窗口停靠在屏幕边缘
                '让检查鼠标位置的Timer工作
                
                '设置显示方向
                If rcWnd.Left = 0 Then
                    m_ShowOrient = 0
                ElseIf rcWnd.Right >= Screen.Width / Screen.TwipsPerPixelX Then
                    m_ShowOrient = 1
                ElseIf rcWnd.Top = 0 Then
                    m_ShowOrient = 2
                End If
                tmrCheck.Enabled = True
            End If
    End Select
    Result = c_Subclass.CallDefaultWindowProc(cHwnd, Message, wParam, lParam)
End Sub

Private Sub Form_Unload(Cancel As Integer)
c_Subclass.SetMsgUnHook
Set c_Subclass = Nothing
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ReleaseCapture
     SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ReleaseCapture
     SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Timer7_Timer()
Dim pth As String
Select Case Check7.Value
Case Checked
    pth = GetActicle(True)
    DoEvents
    If pth <> "" Then
        Label1.Caption = pth
        Readtxt (Label1.Caption)
        myhwnd = GetFirstEditHwnd()
        StatusBar1.SimpleText = "检测到文章，开始打字"
        DoEvents
    End If
Case Unchecked
    pth = GetActicle(False)
    DoEvents
    If pth <> "" Then
        Label1.Caption = pth
        Readtxt (Label1.Caption)
        myhwnd = GetFirstEditHwnd()
        StatusBar1.SimpleText = "检测到文章，开始打字"
        DoEvents
    End If
End Select
If Label1.Caption <> "" Then
        If Option3.Value = True Then
        myhwnd = GetFirstEditHwnd()
        status = 5
        End If
        If Option4.Value = True Then
        myhwnd = GetFirstEditHwnd()
           Timer1.Interval = Text2.Text
            If Check1.Value = Checked Then
            StatusBar1.SimpleText = "用Pageup和Pagedown键上调或下调速度"
            status = 8
            Else
            status = 6
            End If
        End If
Timer1.Enabled = True
End If
End Sub

Private Sub Timer8_Timer()
If Label1.Caption <> "" Then
    If GetFirstEditHwnd <> 0 Then
        If Check1.Value = Checked Then
            status = 10
        Else
            status = 9
        End If
        Timer1.Interval = Text2.Text
        length = 1
        start = 1
        myhwnd = GetFirstEditHwnd
        Timer1.Enabled = True
        StatusBar1.SimpleText = "开始打字"
        Timer8.Enabled = False
    End If
End If
End Sub

Private Sub tmrCheck_Timer()
    Dim pt As POINTAPI
    Dim rc As RECT
    Call GetCursorPos(pt)
    Call GetWindowRect(Me.hwnd, rc)
    If PtInRect(rc, pt.X, pt.Y) Then
        '鼠标停留在窗口上
        If m_ShowFlag = 1 Then Exit Sub
        m_ShowSpeed = SHOWHIDE_SPEED
        m_ShowFlag = 1
        tmrMove.Enabled = True
    Else
        '鼠标不再窗口上
        If m_ShowFlag = 0 Then Exit Sub
        m_ShowSpeed = SHOWHIDE_SPEED
        m_ShowFlag = 0
        tmrMove.Enabled = True
    End If
End Sub

Private Sub tmrMove_Timer()
    Dim nTop    As Long
    Dim nLeft   As Long
    m_ShowSpeed = m_ShowSpeed + SHOWHIDE_SPEED
    '如果大于300T则加快速度
    'If m_ShowSpeed > 300 Then m_ShowSpeed = m_ShowSpeed + m_ShowSpeed * 0.2
    Select Case m_ShowOrient
        Case 0  '0  向左
            If m_ShowFlag = 0 Then
                nLeft = Me.Left - m_ShowSpeed
                If nLeft < -Me.Width + SIZE_SHOW Then nLeft = -Me.Width + SIZE_SHOW: tmrMove.Enabled = False
            Else
                nLeft = Me.Left + m_ShowSpeed
                If nLeft > -SIZE_SHOW Then nLeft = -SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.Left = nLeft
            
        Case 1  '1  向右
            If m_ShowFlag = 0 Then
                nLeft = Me.Left + m_ShowSpeed
                If nLeft > Screen.Width - SIZE_SHOW Then nLeft = Screen.Width - SIZE_SHOW: tmrMove.Enabled = False
            Else
                nLeft = Me.Left - m_ShowSpeed
                If nLeft < Screen.Width - Me.Width + SIZE_SHOW Then nLeft = Screen.Width - Me.Width + SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.Left = nLeft
            
        Case 2  '2  向上
            If m_ShowFlag = 0 Then
                nTop = Me.Top - m_ShowSpeed
                If nTop < -Me.Height + SIZE_SHOW Then nTop = -Me.Height + SIZE_SHOW: tmrMove.Enabled = False
            Else
                nTop = Me.Top + m_ShowSpeed
                If nTop > -SIZE_SHOW Then nTop = -SIZE_SHOW: tmrMove.Enabled = False
            End If
            Me.Top = nTop
            
    End Select
End Sub



Private Sub Check2_Click()
Timer2.Enabled = Check2.Value
End Sub

Private Sub Combo1_Click()
Clipboard.Clear
End Sub

Private Sub Command6_Click()
status = 7
Timer7.Enabled = False
Timer8.Enabled = False
StatusBar1.SimpleText = "打字已取消"
Command6.Enabled = False
Command5.Enabled = True
Command3.Enabled = False
If Check3.Value = Unchecked Then
Option2.Enabled = True
Option3.Enabled = True
End If
End Sub

Private Sub Form_Load()
Me.Hide
Dim XX As Long
XX = SetWindowPos(Me.hwnd, Hwndx, 0, 0, 0, 0, 3)
start = 1
time = 1
length = 1
SSTab1.Tab = 0
DoEvents
Me.Show
Timer3.Enabled = True
Dim i As Integer
For i = 1 To 26
Combo1.AddItem Chr(i + 64)
Next i
   Set c_Subclass = New iSubClass
   c_Subclass.SetMsgHook Me.hwnd
End Sub

Private Sub Readtxt(FileName As String)
Dim i As Integer, s As String, BB() As Byte
If Dir(FileName) = "" Then Exit Sub
i = FreeFile
ReDim BB(FileLen(FileName) - 1)
Open FileName For Binary As #i
Get #i, , BB
Close #i
s = StrConv(BB, vbUnicode)
Lines = s
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

Private Sub Command2_Click()
If Dir1.path <> "" And File1.FileName <> "" Then
path = Dir1.path & "\" & File1.FileName
path = Replace(path, "\\", "\")
Label1.Caption = path
Readtxt (Label1.Caption)
SSTab1.Tab = 0
Else
MsgBox "请先确定路径"
End If
End Sub

Private Sub Command3_Click()
Lines = ""
SSTab1.Tab = 1
Label1.Caption = ""
path = ""
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
Readtxt (Label1.Caption)
SSTab1.Tab = 0
End If
End Sub

Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim fext As String
'检查移过列表框的是不是文件，以及是否只拖放一个文件
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'显示可以放下的图标，是带小加号的那种
fext = LCase(Right(Data.Files(1), 4))
'是否指定的文件类型
If fext = ".txt" Then
Effect = vbDropEffectCopy And Effect
Else
Effect = vbDropEffectNone
End If
Else
'否则显示不可放下的图标，是圆圈加斜线那种
Effect = vbDropEffectNone
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
Text1.Refresh
Command4.Refresh
Command2.Refresh
End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub
Private Sub SSTab1_OLEDragDrop(Data As TabDlg.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 1 And LCase(Right(Data.Files(1), 4)) = ".txt" Then
Label1.Caption = Data.Files(1)
Readtxt (Label1.Caption)
SSTab1.Tab = 0
SSTab1.Tab = 0
End If
End Sub

Private Sub Command5_Click()
If Label1.Caption <> "" Then
    Readtxt (Label1.Caption)
    Lines = Replace(Lines, vbCrLf, "")
    Lines = Replace(Lines, Chr(10) & Chr(32), "")
    Command6.Enabled = True
    Command3.Enabled = False
    Command5.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Option4.Enabled = False
    Option3.Enabled = False
    Text2.Enabled = False
    UpDown1.Enabled = False
    Check1.Enabled = False
    If Check3.Value = Checked Then
        StatusBar1.SimpleText = "等待打开窗口中"
        Timer8.Enabled = True
        Exit Sub
    End If
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
Private Sub Image1_Click()
Timer4.Enabled = True
If Check2.Enabled = True Then
Timer2.Enabled = False
End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(104, 0)
Image1.Picture = LoadResPicture(102, 0)
End Sub
Private Sub Image2_Click()
Me.Hide
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(101, 0)
Image1.Picture = LoadResPicture(103, 0)
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
'检查移过列表框的是不是文件，以及是否只拖放一个文件
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'显示可以放下的图标，是带小加号的那种
fext = LCase(Right(Data.Files(1), 4))
'是否指定的文件类型
If fext = ".txt" Then
Effect = vbDropEffectCopy And Effect
Else
Effect = vbDropEffectNone
End If
Else
'否则显示不可放下的图标，是圆圈加斜线那种
Effect = vbDropEffectNone
End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 1 And LCase(Right(Data.Files(1), 4)) = ".txt" Then
path = Data.Files(1)
Readtxt (Data.Files(1))
Text1.Text = Lines
Lines = ""
End If
End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim fext As String
'检查移过列表框的是不是文件，以及是否只拖放一个文件
If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
'显示可以放下的图标，是带小加号的那种
fext = LCase(Right(Data.Files(1), 4))
'是否指定的文件类型
If fext = ".txt" Then
Effect = vbDropEffectCopy And Effect
Else
Effect = vbDropEffectNone
End If
Else
'否则显示不可放下的图标，是圆圈加斜线那种
Effect = vbDropEffectNone
End If
End Sub

Private Sub Text2_Change()
If Val(Text1.Text) > 2000 Then
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
    StatusBar1.SimpleText = "开始打字"
    status = 5
    End If
Case 2
    Timer1.Interval = 1000
    counttime = counttime - 1
    StatusBar1.SimpleText = "距离打字开始还有" & counttime & "秒"
    If counttime = "0" Then
    StatusBar1.SimpleText = "开始打字"
    If Check1.Value = Checked Then
    DoEvents
    StatusBar1.SimpleText = "用Pageup和Pagedown键上调或下调速度"
    status = 8
    Timer1.Interval = Text2.Text
    Else
    Timer1.Interval = Text2.Text
    status = 6
    End If
    End If
Case 3
    StatusBar1.SimpleText = "请按下Ctrl+Alt+" & Combo1.Text & "启动打字"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc(Combo1.Text)) Then
    StatusBar1.SimpleText = "开始打字"
    status = 5
    End If
Case 4
    StatusBar1.SimpleText = "请按下Ctrl+Alt+" & Combo1.Text & "启动打字"
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(Asc(Combo1.Text)) Then
    StatusBar1.SimpleText = "开始打字"
    If Check1.Value = Checked Then
    DoEvents
    StatusBar1.SimpleText = "用Pageup和Pagedown键上调或下调速度"
    status = 8
    Timer1.Interval = Text2.Text
    Else
    Timer1.Interval = Text2.Text
    status = 6
    End If
    End If
Case 5
    SendKeys Lines, True
    status = 7
Case 6
    Result = Mid(Lines, start, 1)
    start = start + 1
    SendKeys Result, True
    If Result = "" Then
    status = 7
    End If
Case 8
    If Timer1.Interval > 10 Then
    If GetAsyncKeyState(vbKeyPageUp) Then
    Timer1.Interval = Timer1.Interval - 5
    End If
    End If
    If GetAsyncKeyState(vbKeyPageDown) Then
    Timer1.Interval = Timer1.Interval + 5
    End If
    Result = Mid(Lines, start, 1)
    start = start + 1
    SendKeys Result, True
    If Result = "" Then
    status = 7
    End If
Case 9
If myhwnd <> 0 Then
    Result = Mid(Lines, start, length)
    DoEvents
    SendMessage myhwnd, WM_SETTEXT, 0, ByVal (Result)
    length = length + 1
    If IsWindowEnabled(GetNextWindow(myhwnd, GW_HWNDPREV)) <> 0 Then
            myhwnd = GetNextWindow(myhwnd, GW_HWNDPREV)
            start = start + length - 1
            length = 1
            DoEvents
    Else
        If GetNextWindow(myhwnd, GW_HWNDPREV) = 0 Then
            status = 7
        End If
    End If
Else
    status = 7
End If
Case 10
    If Timer1.Interval > 10 Then
    If GetAsyncKeyState(vbKeyPageUp) Then
    Timer1.Interval = Timer1.Interval - 5
    End If
    End If
    If GetAsyncKeyState(vbKeyPageDown) Then
    Timer1.Interval = Timer1.Interval + 5
    End If
If myhwnd <> 0 Then
    Result = Mid(Lines, start, length)
    DoEvents
    SendMessage myhwnd, WM_SETTEXT, 0, ByVal (Result)
    length = length + 1
    If IsWindowEnabled(GetNextWindow(myhwnd, GW_HWNDPREV)) <> 0 Then
            myhwnd = GetNextWindow(myhwnd, GW_HWNDPREV)
            start = start + length - 1
            length = 1
            DoEvents
    Else
        If GetNextWindow(myhwnd, GW_HWNDPREV) = 0 Then
            status = 7
        End If
    End If
Else
    status = 7
End If
Case 7
    StatusBar1.SimpleText = "打字结束"
    Command5.Enabled = True
    Command3.Enabled = True
    Command6.Enabled = False
    Option4.Enabled = True
    Option3.Enabled = True
    If Option4.Value = True Then
    Text2.Enabled = True
    UpDown1.Enabled = True
    Check1.Enabled = True
    End If
    If Check3.Value = Unchecked Then
    Option1.Enabled = True
    Option2.Enabled = True
    End If
    length = 1
    start = 1
    counttime = 6
    Timer1.Enabled = False
    DoEvents
    StatusBar1.SimpleText = "就绪"
End Select
End Sub

Private Sub Timer2_Timer()
Dim p As POINTAPI, r As RECT
GetCursorPos p
GetWindowRect Me.hwnd, r
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then
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
time = time + 5
Else
Timer3.Enabled = False
End If
SetFormToAlpha Me.hwnd, time
End Sub
Private Sub Timer4_Timer()
If time > 0 Then
time = time - 5
End If
SetFormToAlpha Me.hwnd, time
If time = 1 Then
Unload Form1
End If
End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadResPicture(104, 0)
 Image1.Picture = LoadResPicture(103, 0)
 End Sub
 
Private Sub Timer5_Timer()
If time >= 5 Then
time = time - 5
Else
Timer5.Enabled = False
Me.Visible = False
End If
SetFormToAlpha Me.hwnd, time
End Sub

Private Sub Timer6_Timer()
Dim p As POINTAPI, r As RECT
GetCursorPos p
GetWindowRect Me.hwnd, r
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then
Image2.Picture = LoadResPicture(104, 0)
 Image1.Picture = LoadResPicture(103, 0)
End If
End Sub

