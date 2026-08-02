' vybe-test: vb/vb_task_completion_source_set_result/test_vb_tcs_async_await_bridging
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
    Private Async Function BridgeAsync(tcs As TaskCompletionSource(Of String)) As Task(Of String)
        Dim res = Await tcs.Task
        Return "Bridged: " & res
    End Function

    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim bgTask = BridgeAsync(tcs)
        tcs.SetResult("Payload")
        __Check(CStr(bgTask.Result), "Bridged: Payload")
    End Sub
End Module
