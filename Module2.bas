Attribute VB_Name = "Module2"
 Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
 Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) _
                As Long
 Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Long, _
                ByVal dwFlags As Long) _
                As Long
 Const GWL_EXSTYLE = (-20)
 Const LWA_ALPHA As Long = &H2
 Const WS_EX_LAYERED As Long = &H80000
 Const GW_HWNDFIRST = 0
 Const GW_HWNDLAST = 1
 Const GW_HWNDNEXT = 2
 Const GW_HWNDPREV = 3
 Declare Function SendMessage _
 Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, _
 ByVal wMsg As Long, _
 ByVal wParam As Long, _
 lParam As Any) As Long
 Declare Function FindWindow _
 Lib "user32" _
 Alias "FindWindowA" (ByVal lpClassName As String, _
 ByVal lpWindowName As String) As Long
 Declare Function GetDlgItem _
 Lib "user32" (ByVal hDlg As Long, _
 ByVal nIDDlgItem As Long) As Long
 Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
 Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
 Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
 Declare Function GetDesktopWindow Lib "user32" () As Long

'---------------------------------------------------------API------------------------------------------------------------------------------------------------------------------------

Public Function SetFormToAlpha(hwnd As Long, lngAlpha As Long)
    Dim tmpLog As Long
    If hwnd = 0 Then Exit Function
    If lngAlpha >= 0 And lngAlpha <= 255 Then
        tmpLog = GetWindowLong(hwnd, GWL_EXSTYLE) '窗口属性
        Call SetWindowLong(hwnd, GWL_EXSTYLE, tmpLog Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hwnd, 0, lngAlpha, LWA_ALPHA)
    End If
End Function

 '-----------------------------------------------------------------Ico-----------------------------------------------------------------------------------------------
Function GetFirstEditHwnd() As Long
Dim MainForm_hwnd, ControlPanel_hwnd, TypeFrm_hwnd, MainControl_hWnd, Edit_hwnd As Long
MainForm_hwnd = FindWindow("TFrmTester", vbNullString)
ControlPanel_hwnd = FindWindowEx(MainForm_hwnd, 0, "TPanel", vbNullString)
 TypeFrm_hwnd = FindWindowEx(ControlPanel_hwnd, 0, "TFrmTyping", vbNullString)
 MainControl_hWnd = FindWindowEx(TypeFrm_hwnd, 0, "TPanel", vbNullString)
Edit_hwnd = FindWindowEx(MainControl_hWnd, 0, "TEdit", vbNullString)
GetFirstEditHwnd = GetNextWindow(Edit_hwnd, GW_HWNDLAST) '在学生端中的打字框，顺序是颠倒的，即第一个就是最后一个，下一个就是上一个
End Function

'------------------------------------------------------------------Hwnd------------------------------------------------------------------------------------------------------

'Function ReadRARPath() As String
'On Error Resume Next
'Dim ws As Object
    'Set ws = CreateObject("WScript.Shell")
     'ReadRARPath = ws.RegRead("HKEY_CURRENT_USER\Software\WinRAR\Paths\TempFolder")
'unable module
'End Function

Function GetActicle(savess As Boolean) As String
Dim Id_e As Integer
Dim Id_c As Integer
Dim path As String
Dim mypath As String
path = ReadRARPath
    For Id_e = 1 To 255
        If Dir(path & "e" & Id_e & ".txt") <> "" Then
            If Dir(App.path & "\article", vbDirectory) = "" Then
            MkDir (App.path & "\article")
            End If
        If savess = True Then
            Call FileCopy(path & "e" & Id_e & ".txt", App.path & "\article\" & "e" & Id_e & ".txt")
        End If
        'GetActicle = path & "e" & Id_e & ".txt"
        Exit Function
        End If
    Next Id_e
    For Id_c = 1 To 255
        If Dir(path & "c" & Id_c & ".txt") <> "" Then
            If Dir(App.path & "\article", vbDirectory) = "" Then
            MkDir (App.path & "\article\" & "c" & Id_c & ".txt")
            End If
        If savess = True Then
            Call FileCopy(path & "c" & Id_c & ".txt", App.path & "\article\" & "c" & Id_c & ".txt")
        End If
        'GetActicle = path & "c" & Id_c & ".txt"
        Exit Function
        End If
    Next Id_c
GetActicle = ""
End Function


Function ReadtxtLines(txtpaths As String) As Long
Dim s
Dim Freenum As Integer
Freenum = FreeFile
Open txtpaths For Binary As Freenum
 s = Split(Input$(LOF(1), 1), vbCrLf)
 Close #1
ReadtxtLines = UBound(s) + 1
End Function

Function GetFile(FileName As String) As String
Dim i As Integer, s As String, BB() As Byte
If Dir(FileName) = "" Then Exit Function
i = FreeFile
ReDim BB(FileLen(FileName) - 1)
Open FileName For Binary As #i
Get #i, , BB
Close #i
s = StrConv(BB, vbUnicode)
GetFile = s
End Function


Function Gethwnddata(datapath As String) As Long
Dim Classname() As String
End Function

