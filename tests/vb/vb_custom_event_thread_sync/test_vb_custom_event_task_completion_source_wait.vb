' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_task_completion_source_wait
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

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

Class TaskCompletionPublisher
    Public Event TaskCompleted As EventHandler

    Public Sub RunWork()
        RaiseEvent TaskCompleted(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Boolean)()
        Dim pub As New TaskCompletionPublisher()

        AddHandler pub.TaskCompleted, Sub(s, e) tcs.SetResult(True)
        pub.RunWork()

        __Check(CStr("Task Result: " & tcs.Task.Result), "Task Result: True")
    End Sub
End Module
