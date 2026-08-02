' vybe-test: vb/vb_async_task_run_lambda/test_vb_async_function_return_task_of_t
' origin: languages/vb/tests/vb/test_vb_async_task_run_lambda.rs

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

Imports System.Threading.Tasks

Module Program
    Async Function ComputeAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 42
    End Function

    Sub Main()
        Dim t = ComputeAsync()
        __Check(CStr(t.Result), "42")
    End Sub
End Module
