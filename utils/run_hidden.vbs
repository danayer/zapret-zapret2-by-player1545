Option Explicit

If WScript.Arguments.Count < 1 Then
    WScript.Quit 1
End If

Dim targetPath
targetPath = WScript.Arguments(0)

Dim shell
Set shell = CreateObject("WScript.Shell")

Dim command
command = "cmd.exe /c """ & Replace(targetPath, """", """""") & """"

shell.Run command, 0, False
WScript.Quit 0
