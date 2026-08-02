' vybe-test: vb/vb_parse_enum_ignore_case/test_vb_enum_try_parse_generic_ignore_case
' origin: languages/vb/tests/vb/test_vb_parse_enum_ignore_case.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Enum LogLevel
    Debug
    Warning
    Error
End Enum

Module Program
    Sub Main()
        Dim level As LogLevel
        Dim ok = [Enum].TryParse(Of LogLevel)("warning", True, level)
        __Check(CStr(ok & "|" & level.ToString()), "True|Warning")
    End Sub
End Module
