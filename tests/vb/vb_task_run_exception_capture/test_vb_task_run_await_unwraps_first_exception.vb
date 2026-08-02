' vybe-test: vb/vb_task_run_exception_capture/test_vb_task_run_await_unwraps_first_exception
' origin: languages/vb/tests/vb/test_vb_task_run_exception_capture.rs

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
    Private Async Function RunFaultyTaskAsync() As Task
        Await Task.Run(Sub()
            Throw New ArgumentNullException("param", "Argument null in task")
        End Sub)
    End Function

    Sub Main()
        Try
            Dim t = RunFaultyTaskAsync()
            t.Wait()
        Catch ex As AggregateException
            Dim inner = ex.InnerException
            __Check(CStr(inner.GetType().Name & ": " & inner.Message), "ArgumentNullException: Argument null in task
Parameter name: param")
        End Try
    End Sub
End Module
