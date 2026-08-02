' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_when_filter_clause
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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
    Private Sub PerformAction(code As Integer)
        Try
            Throw New ArgumentException("Error", "ParamName")
        Catch ex As ArgumentException When code = 1
            __Check(CStr("Handled Code 1"), "Handled Code 1")
        Catch ex As ArgumentException When code = 2
            __Check(CStr("Handled Code 2"), "Handled Code 2")
        Catch ex As Exception
            __Check(CStr("Handled Fallback"), "Handled Fallback")
        End Try
    End Sub

    Sub Main()
        PerformAction(1)
        PerformAction(2)
        PerformAction(3)
    End Sub
End Module
