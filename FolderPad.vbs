' =====================================================================
' FolderPad launcher (Windows Script Host, pure ASCII - no console flash)
' Usage:
'   FolderPad.vbs                    start the GUI
'   FolderPad.vbs "D:\some\folder"   convert that folder into a pad
' This file only forwards to FolderPad.ps1; all logic lives in the script.
' =====================================================================
Option Explicit

Dim fso, sh, baseDir, ps1, psExe, args, i, a, cmd

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = baseDir & "\FolderPad.ps1"
psExe = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(ps1) Then
    MsgBox "Cannot find FolderPad.ps1: " & vbCrLf & ps1, 16, "FolderPad"
    WScript.Quit 1
End If

args = ""
For i = 0 To WScript.Arguments.Count - 1
    a = WScript.Arguments(i)
    args = args & " " & Chr(34) & a & Chr(34)
Next

cmd = Chr(34) & psExe & Chr(34) & " -NoProfile -Sta -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & ps1 & Chr(34) & args
sh.Run cmd, 0, False
