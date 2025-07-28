Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("MSXML2.XMLHTTP")
Set stream = CreateObject("ADODB.Stream")

ghostDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Ghost"
If Not fso.FolderExists(ghostDir) Then fso.CreateFolder(ghostDir)

url = "https://raw.githubusercontent.com/dakj0326/ghost/main/s1.mp3"
path = ghostDir & "\ghost_message.txt"

x.Open "GET", url, False
x.Send

If x.Status = 200 Then
    stream.Type = 2  ' text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText x.ResponseText
    stream.SaveToFile path, 2  ' 2 = overwrite
    stream.Close
    MsgBox "Download complete: " & path
Else
    MsgBox "Download failed. HTTP status: " & x.Status
End If
