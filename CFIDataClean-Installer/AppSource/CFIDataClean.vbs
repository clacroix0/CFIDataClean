Option Explicit

Dim fso
Dim shell
Dim root
Dim appDir
Dim ps1
Dim psExe
Dim cmd

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

root = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = fso.BuildPath(root, "_CFIDataClean_AppFiles")
ps1 = fso.BuildPath(appDir, "ForestInventoryCleaner.ps1")

If Not fso.FolderExists(appDir) Then
    MsgBox "CFI DataClean could not find the app files folder:" & vbCrLf & vbCrLf & appDir, vbCritical, "CFI DataClean"
    WScript.Quit 1
End If

If Not fso.FileExists(ps1) Then
    MsgBox "CFI DataClean could not find the PowerShell app file:" & vbCrLf & vbCrLf & ps1, vbCritical, "CFI DataClean"
    WScript.Quit 1
End If

psExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(psExe) Then
    psExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

If Not fso.FileExists(psExe) Then
    MsgBox "Windows PowerShell was not found.", vbCritical, "CFI DataClean"
    WScript.Quit 1
End If

shell.CurrentDirectory = appDir

cmd = """" & psExe & """ -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File """ & ps1 & """"

shell.Run cmd, 0, False