Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ghostDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Ghost"

MsgBox "ghost opened"

Do
    WScript.Sleep 3000  ' test delay (20 min = 1200000)

    Randomize
    n = Int((6 * Rnd) + 1)
    soundFile = ghostDir & "\s" & n & ".mp3"

    If fso.FileExists(soundFile) Then
        Set player = CreateObject("WMPlayer.OCX")
        Set media = player.newMedia(soundFile)

        ' Add to playlist and play
        player.currentPlaylist.appendItem media
        player.controls.play

        ' Let the sound play (adjust time as needed)
        WScript.Sleep 10000

        Set player = Nothing
        Set media = Nothing
    End If
Loop
