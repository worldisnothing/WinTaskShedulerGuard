Option Explicit

Dim WshShell, exec, lines, lastLine, i, pos2, pos1, result
Set WshShell = CreateObject("WScript.Shell")

' Берём все события ID 106 (создать задачу) из журнала TaskScheduler
Set exec = WshShell.Exec("wevtutil qe Microsoft-Windows-TaskScheduler/Operational /q:*[System[(EventID=106)]] /f:text")

' Записываем все строки выполнения в переменную
lines = Split(exec.StdOut.ReadAll, vbCrLf)

' Определяем последнюю не пустую строку, т.к. сверху - старые, снизу - новые события.
lastLine = ""
For i = UBound(lines) To 0 Step -1
    If Trim(lines(i)) <> "" Then
        lastLine = lines(i)
        Exit For
    End If
Next

' Форматируем. Извлекаем из строки название между двойными кавычками
pos2 = InStrRev(lastLine, """")
pos1 = InStrRev(lastLine, """", pos2 - 1)
If pos1 > 0 And pos2 > pos1 Then
' +2 т.к. после открытой кавычки добавляется слэш
    result = Mid(lastLine, pos1 + 2, pos2 - pos1 - 2)
Else
    result = ""
End If

' Формируем сообщение на вывод. 
Dim message, title
If result = "" Then
    message = "Не удалось определить последнюю добавленную задачу. Но она есть!"
Else
    message = "Добавлена новая задача в Планировщике:" & vbCrLf & result
End If
title = "Sheduler Monitor"

MsgBox message, 64, title
