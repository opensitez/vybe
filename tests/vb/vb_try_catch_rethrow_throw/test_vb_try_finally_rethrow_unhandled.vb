' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_finally_rethrow_unhandled
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
    Private Sub Outer()
        Try
            Inner()
        Catch ex As Exception
            __Check(CStr("Outer Caught: " & ex.Message), "Inner Finally Executed")
        End Try
    End Sub

    Private Sub Inner()
        Try
            Throw New Exception("Deep Error")
        Finally
            __Check(CStr("Inner Finally Executed"), "Outer Caught: Deep Error")
        End Try
    End Sub

    Sub Main()
        Outer()
    End Sub
End Module
