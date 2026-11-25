Attribute VB_Name = "TrimmerKeys"
Option Explicit

' Gives the trimmer its keys when the add-in Trimmer.xlam loads (designs/stage-B-trimmer.md 13):
' Ctrl+Shift+T runs the trim, Ctrl+Shift+Y the self-test. Excel runs Auto_Open when it loads an
' add-in and Auto_Close when it unloads it. The keys are only set while the add-in is loaded.
Public Const KEY_TRIM As String = "^+T"
Public Const KEY_SELFTEST As String = "^+Y"

Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey KEY_TRIM, "'" & ThisWorkbook.Name & "'!TrimStock"
    Application.OnKey KEY_SELFTEST, "'" & ThisWorkbook.Name & "'!RunSelfTest"
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey KEY_TRIM
    Application.OnKey KEY_SELFTEST
End Sub
