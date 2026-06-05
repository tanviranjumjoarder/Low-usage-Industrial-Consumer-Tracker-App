' ============================================================
'  Silent launcher for Electricity Bill Merger
'  Double-click this file to start the app (no console window)
' ============================================================
Option Explicit

Dim sh, fso, scriptDir, pyScript, cmd, pyExe

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyScript  = scriptDir & "\merge_app.py"

If Not fso.FileExists(pyScript) Then
    MsgBox "merge_app.py was not found next to this launcher." & vbCrLf & _
           "Expected at: " & pyScript, vbCritical, "Electricity Bill Merger"
    WScript.Quit 1
End If

' Prefer pythonw.exe (no console). Fall back to python.exe if pythonw is not on PATH.
pyExe = "pythonw"
cmd   = """" & pyExe & """ """ & pyScript & """"

' Run hidden (0) and don't wait for exit (False)
On Error Resume Next
sh.Run cmd, 0, False
If Err.Number <> 0 Then
    Err.Clear
    pyExe = "python"
    cmd   = """" & pyExe & """ """ & pyScript & """"
    sh.Run cmd, 0, False
    If Err.Number <> 0 Then
        MsgBox "Could not find Python on this computer." & vbCrLf & _
               "Please install Python 3.8+ from https://www.python.org and try again.", _
               vbCritical, "Electricity Bill Merger"
    End If
End If
