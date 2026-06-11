Option Explicit

If WScript.Arguments.Count < 1 Then
    WScript.Quit 1
End If

Dim targetPath
targetPath = WScript.Arguments(0)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(targetPath) Then
    WScript.Quit 2
End If

Dim ext
ext = LCase(fso.GetExtensionName(targetPath))
If ext <> "bat" And ext <> "cmd" Then
    WScript.Quit 3
End If

If InStr(targetPath, "&") > 0 Or InStr(targetPath, "|") > 0 Or InStr(targetPath, "<") > 0 Or InStr(targetPath, ">") > 0 Or InStr(targetPath, "^") > 0 Then
    WScript.Quit 4
End If

Dim shell
Set shell = CreateObject("WScript.Shell")

Dim command
command = "cmd.exe /c call """ & Replace(targetPath, """", """""") & """"

shell.Run command, 0, False
WScript.Quit 0
