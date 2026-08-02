' vybe-test: vb/vb_async_task_delay_cancellation/test_vb_async_task_when_all_multiple_delays
' origin: languages/vb/tests/vb/test_vb_async_task_delay_cancellation.rs

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
    Private Async Function RunParallelDelaysAsync() As Task(Of String)
        Dim t1 = Task.Delay(10)
        Dim t2 = Task.Delay(20)
        Await Task.WhenAll(t1, t2)
        Return "All Delays Finished"
    End Function

    Sub Main()
        Dim t = RunParallelDelaysAsync()
        __Check(CStr(t.Result), "All Delays Finished")
    End Sub
End Module
