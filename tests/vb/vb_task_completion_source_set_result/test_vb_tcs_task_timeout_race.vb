' vybe-test: vb/vb_task_completion_source_set_result/test_vb_tcs_task_timeout_race
' origin: languages/vb/tests/vb/test_vb_task_completion_source_set_result.rs

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
    Private Async Function GetWithTimeoutAsync(tcs As TaskCompletionSource(Of String), timeoutMs As Integer) As Task(Of String)
        Dim delayTask = Task.Delay(timeoutMs)
        Dim completed = Await Task.WhenAny(tcs.Task, delayTask)
        If completed Is tcs.Task Then
            Return Await tcs.Task
        Else
            Return "Timeout"
        End If
    End Function

    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim t = GetWithTimeoutAsync(tcs, 5)
        ' Don't resolve tcs, let timeout win
        __Check(CStr(t.Result), "Timeout")
    End Sub
End Module
