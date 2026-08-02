' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_throw_ex_resets_stack_trace
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

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
    Private Sub Level1()
        Throw New InvalidOperationException("Root")
    End Sub

    Private Sub Level2()
        Try
            Level1()
        Catch ex As Exception
            Throw ex ' Re-throwing ex resets the stack trace to Level2
        End Try
    End Sub

    Sub Main()
        Try
            Level2()
        Catch ex As Exception
            __Check(CStr(ex.StackTrace.Contains("Level2")), "True")
        End Try
    End Sub
End Module
