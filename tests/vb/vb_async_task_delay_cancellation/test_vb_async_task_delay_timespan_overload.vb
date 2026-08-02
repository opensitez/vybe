' vybe-test: vb/vb_async_task_delay_cancellation/test_vb_async_task_delay_timespan_overload
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

Imports System
Imports System.Threading.Tasks

Module Program
    Private Async Function RunTimeSpanDelayAsync() As Task(Of String)
        Await Task.Delay(TimeSpan.FromMilliseconds(10))
        Return "TimeSpan Delay Passed"
    End Function

    Sub Main()
        Dim t = RunTimeSpanDelayAsync()
        __Check(CStr(t.Result), "TimeSpan Delay Passed")
    End Sub
End Module
