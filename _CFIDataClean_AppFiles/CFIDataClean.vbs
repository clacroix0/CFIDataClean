Option Explicit

Dim shell, fso, scriptDir, appScript, ps32, command, exitCode

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appScript = fso.BuildPath(scriptDir, "ForestInventoryCleaner.ps1")
ps32 = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(ps32) Then
    ps32 = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

If Not fso.FileExists(appScript) Then
    MsgBox "DADA could not find its app files." & vbCrLf & vbCrLf & _
        "Expected file:" & vbCrLf & appScript & vbCrLf & vbCrLf & _
        "If you opened DADA from a ZIP/compressed folder, click Extract All first, then run DADA from the extracted folder.", _
        vbCritical, "Database Dad"
    WScript.Quit 1
End If

command = """" & ps32 & """ -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & appScript & """"
exitCode = shell.Run(command, 0, True)

If exitCode <> 0 Then
    MsgBox "Database Dad closed because of an error. Try _DADA_AppFiles\DADA-debug.cmd to see the detailed PowerShell message.", vbExclamation, "Database Dad"
End If

WScript.Quit exitCode
