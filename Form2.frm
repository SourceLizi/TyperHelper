VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "句柄查找工具"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Text            =   "hwnd1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   1350
      Left            =   360
      Pattern         =   "*.hwnd"
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "拖拽靶标以查找窗体"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "句柄工具"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "保存名称"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2640
      Picture         =   "Form2.frx":030A
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "当前句柄"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsDragging As Boolean
Dim SnapHwnd&
Dim DeskHwnd&, DeskDC&
Dim oldRop2&
Dim rc As RECT
Dim i As Integer

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim CurrentClassname(), Titlename() As String * 255
Dim Classname As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long



Private Sub Command1_Click()
Dim CurrentHwnd(), t As Long
ReDim Preserve CurrentHwnd(1)
ReDim Preserve CurrentClassname(1)
ReDim Preserve Titlename(1)
CurrentHwnd(0) = Val(Text1.Text)
CurrentClassname(0) = Space(255)
Titlename(0) = Space(255)
t = 0
If Text1.Text = "" Or Text1.Text = 0 Then Exit Sub
If Text2.Text = "" Then
MsgBox "名称不能为空", , "错误"
Exit Sub
End If
Open App.path & "\" & Text2.Text & ".hwnd" For Output As #1
    Do
    CurrentClassname(t) = Space(255)
    GetClassName CurrentHwnd(t), CurrentClassname(t), 255
    GetWindowText CurrentHwnd(t), Titlename(t), 255
    Print #1, Replace(CurrentClassname(t), " ", "") & "|" & Replace(Titlename(t), " ", "")
    t = t + 1
    ReDim Preserve CurrentHwnd(t)
    ReDim Preserve CurrentClassname(t)
    ReDim Preserve Titlename(t)
    CurrentHwnd(t) = GetParent(CurrentHwnd(t - 1))
    CurrentClassname(t) = Space(255)
        If CurrentHwnd(t) <> 0 Then
            Print #1, "<"
        Else
            Exit Do
        End If
    Loop
Close #1
MsgBox "保存成功", , "成功"
End Sub

Private Sub Form_Load()
IsDragging = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadResPicture(103, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ReleaseCapture
     SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



Private Sub Image1_Click()
Unload Form2
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadResPicture(102, 0)
End Sub



Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ReleaseCapture
     SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If IsDragging = False Then    '判断是否为拖动状态
        IsDragging = True
        Screen.MousePointer = vbCustom
        Screen.MouseIcon = LoadResPicture(104, 1)  '用鼠标指针变为靶状
        Picture1.Picture = LoadResPicture(102, 1)    '此时图片框加载另一无靶图标
        '将以后的鼠标输入消息都发送到本程序窗口
        SetCapture (Picture1.hwnd)
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False   '让闪烁的矩形消失
    If IsDragging = True Then
        Screen.MousePointer = vbDefault
        IsDragging = False
        ReleaseCapture
        If SnapHwnd& = 0 Then Exit Sub
        Text1.Text = Val(SnapHwnd)
        Classname = Space(255)
        GetClassName Val(SnapHwnd), Classname, 255
        Text2.Text = Classname
        Picture1.Picture = LoadResPicture(103, 1)
    End If
End Sub

Private Sub Timer1_Timer()
Dim pnt As POINTAPI
    Dim newPen&, oldPen&
    'Dim SnapHwnd&
    Dim DeskHwnd&, DeskDC&
    Dim oldRop2&

    DeskHwnd& = GetDesktopWindow()    '取得桌面句柄
    DeskDC& = GetWindowDC(DeskHwnd&)     '取得桌面设备场景
    '
    oldRop2& = SetROP2(DeskDC&, 10)
    GetCursorPos pnt                '取得鼠标坐标

    SnapHwnd = WindowFromPoint(pnt.X, pnt.Y)      '取得鼠标指针处窗口句柄
    GetWindowRect SnapHwnd, rc        '获得窗口矩形
    If rc.Left < 0 Then rc.Left = 0
    If rc.Top < 0 Then rc.Top = 0
    If rc.Right > Screen.Width / 15 Then rc.Right = Screen.Width / 15
    If rc.Bottom > Screen.Height / 15 Then rc.Bottom = Screen.Height / 15
    newPen& = CreatePen(0, 3, &H0)       '建立新画笔,载入DeskDC
    oldPen& = SelectObject(DeskDC, newPen)
    Rectangle DeskDC, rc.Left, rc.Top, rc.Right, rc.Bottom     '在指示窗口周围显示闪烁矩形
    Sleep 200 'Timer1.Interval    '设置闪烁时间间隔
    
    Rectangle DeskDC, rc.Left, rc.Top, rc.Right, rc.Bottom

    SetROP2 DeskDC, oldRop2
    SelectObject DeskDC, oldPen
    DeleteObject newPen
    ReleaseDC DeskHwnd, DeskDC: DeskDC = 0
End Sub
