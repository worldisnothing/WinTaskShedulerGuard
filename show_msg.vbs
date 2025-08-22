Option Explicit

Dim WshShell, exec, lines, lastLine, i, pos2, pos1, result
Set WshShell = CreateObject("WScript.Shell")

' ���� ��� ������� ID 106 (������� ������) �� ������� TaskScheduler
Set exec = WshShell.Exec("wevtutil qe Microsoft-Windows-TaskScheduler/Operational /q:*[System[(EventID=106)]] /f:text")

' ���������� ��� ������ ���������� � ����������
lines = Split(exec.StdOut.ReadAll, vbCrLf)

' ���������� ��������� �� ������ ������, �.�. ������ - ������, ����� - ����� �������.
lastLine = ""
For i = UBound(lines) To 0 Step -1
    If Trim(lines(i)) <> "" Then
        lastLine = lines(i)
        Exit For
    End If
Next

' �����������. ��������� �� ������ �������� ����� �������� ���������
pos2 = InStrRev(lastLine, """")
pos1 = InStrRev(lastLine, """", pos2 - 1)
If pos1 > 0 And pos2 > pos1 Then
' +2 �.�. ����� �������� ������� ����������� ����
    result = Mid(lastLine, pos1 + 2, pos2 - pos1 - 2)
Else
    result = ""
End If

' ��������� ��������� �� �����. 
Dim message, title
If result = "" Then
    message = "�� ������� ���������� ��������� ����������� ������. �� ��� ����!"
Else
    message = "��������� ����� ������ � ������������:" & vbCrLf & result
End If
title = "Sheduler Monitor"

MsgBox message, 64, title
