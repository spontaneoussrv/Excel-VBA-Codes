Attribute VB_Name = "ScriptingRuntime"
Option Explicit

Sub AddScriptingRuntime()
    Dim ref As Object
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References.AddFromFile("C:\Windows\System32\scrrun.dll")
    On Error GoTo 0
    If ref Is Nothing Then
        MsgBox "Microsoft Scripting Runtime is already added or could not be added."
    Else
        MsgBox "Microsoft Scripting Runtime added successfully."
    End If
End Sub

