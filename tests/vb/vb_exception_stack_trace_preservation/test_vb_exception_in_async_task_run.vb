' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_in_async_task_run
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
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Try
            Dim t = Task.Run(Sub() Throw New InvalidOperationException("Async Error"))
            t.Wait()
        Catch ex As AggregateException
            __Check(CStr("Async Caught: " & ex.InnerException.Message), "Async Caught: Async Error")
        End Try
    End Sub
End Module
