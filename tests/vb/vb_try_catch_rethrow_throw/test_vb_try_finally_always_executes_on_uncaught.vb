' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_finally_always_executes_on_uncaught
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
    Private Sub Execute()
        Try
            __Check(CStr("Inside Try"), "Inside Try")
            Throw New InvalidOperationException()
        Finally
            __Check(CStr("Inside Finally"), "Inside Finally")
        End Try
    End Sub

    Sub Main()
        Try
            Execute()
        Catch ex As Exception
            __Check(CStr("Caught in Main"), "Caught in Main")
        End Try
    End Sub
End Module
