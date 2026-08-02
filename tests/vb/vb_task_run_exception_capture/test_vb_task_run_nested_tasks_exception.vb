' vybe-test: vb/vb_task_run_exception_capture/test_vb_task_run_nested_tasks_exception
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
    Sub Main()
        Dim parent = Task.Run(Function()
            Dim child = Task.Run(Sub() Throw New Exception("Nested Exception"))
            child.Wait()
            Return 0
        End Function)

        Try
            parent.Wait()
        Catch ex As AggregateException
            Dim flat = ex.Flatten()
            __Check(CStr(flat.InnerException.Message), "Nested Exception")
        End Try
    End Sub
End Module
