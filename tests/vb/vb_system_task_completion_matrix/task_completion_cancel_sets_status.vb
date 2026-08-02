' vybe-test: vb/vb_system_task_completion_matrix/task_completion_cancel_sets_status
' origin: languages/vb/tests/vb/test_vb_system_task_completion_matrix.rs

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
Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim task As Task = cts.Task

        cts.Cancel()

        __Check(CStr(task.IsCanceled), "True")
        __Check(CStr(cts.IsCancellationRequested), "True")
    End Sub
End Module
