Attribute VB_Name = "Module2"
Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) _
                As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Long, _
                ByVal dwFlags As Long) _
                As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000


Public Function SetFormToAlpha(hWnd As Long, lngAlpha As Long)
    Dim tmpLog As Long
    If hWnd = 0 Then Exit Function
    If lngAlpha >= 0 And lngAlpha <= 255 Then
        tmpLog = GetWindowLong(hWnd, GWL_EXSTYLE) '´°¿ÚÊôÐÔ
        Call SetWindowLong(hWnd, GWL_EXSTYLE, tmpLog Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hWnd, 0, lngAlpha, LWA_ALPHA)
    End If
End Function


