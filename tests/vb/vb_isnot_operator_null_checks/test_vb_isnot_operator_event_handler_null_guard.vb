' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_event_handler_null_guard
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

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

Module Program
    Public Event CustomEvent As EventHandler

    Sub Main()
        ' In VB.NET CustomEventEvent field can be checked for IsNot Nothing before raising
        __Check(CStr(CustomEventEvent IsNot Nothing), "False")
    End Sub
End Module
