Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long

Sub Main()
On Error GoTo Err
Dim i As Long, s() As String, x As Boolean
'MsgBox Command
s = Split(Trim(Command), ",")
'MsgBox UBound(s)
If UBound(s) = 0 Then
    i = CLng(Trim(Command))
    x = True
Else
    i = CLng(Trim(s(1)))
    x = True
    If s(0) = "D" Then x = False
End If
'MsgBox get_path_from_reg("", "winbuild")
If CLng(get_path_from_reg("", "winbuild")) >= 19041 Then
    If x Then DwmSetWindowAttribute i, 20, True, 4
    If Not x Then DwmSetWindowAttribute i, 20, False, 4
Else
    If CLng(get_path_from_reg("", "winbuild")) >= 16299 And CLng(get_path_from_reg("", "winbuild")) < 19041 Then
        If x Then DwmSetWindowAttribute i, 19, True, 4
        If Not x Then DwmSetWindowAttribute i, 19, False, 4
    End If
End If
Err:
End
End Sub


