' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_thrown_from_lambda_closure
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
    Sub Main()
        Dim lambda As Action = Sub() Throw New InvalidOperationException("Closure Error")
        Try
            lambda()
        Catch ex As Exception
            __Check(CStr("Caught Lambda: " & ex.Message), "Caught Lambda: Closure Error")
        End Try
    End Sub
End Module
