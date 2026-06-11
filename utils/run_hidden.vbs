Option Explicit

If WScript.Arguments.Count < 1 Then
    WScript.Quit 1
End If

Dim targetPath
targetPath = WScript.Arguments(0)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not IsAbsolutePath(targetPath) Then
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    targetPath = fso.BuildPath(fso.GetParentFolderName(scriptDir), targetPath)
End If

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

Function IsAbsolutePath(path)
    If Len(path) >= 3 Then
        Dim driveLetter
        driveLetter = UCase(Left(path, 1))
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", driveLetter) > 0 And Mid(path, 2, 1) = ":" And (Mid(path, 3, 1) = "\" Or Mid(path, 3, 1) = "/") Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If
    If Len(path) >= 2 Then
        If Left(path, 2) = "\\" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If
    IsAbsolutePath = False
End Function
