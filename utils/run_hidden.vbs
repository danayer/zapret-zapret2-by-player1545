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

Dim shell
Set shell = CreateObject("WScript.Shell")

shell.Run Chr(34) & targetPath & Chr(34), 0, False
WScript.Quit 0
