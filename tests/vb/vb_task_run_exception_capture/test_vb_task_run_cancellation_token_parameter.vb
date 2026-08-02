' vybe-test: vb/vb_task_run_exception_capture/test_vb_task_run_cancellation_token_parameter
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

Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()
        Dim t = Task.Run(Sub()
            ' Action doesn't run because token was already canceled
        End Sub, cts.Token)

        __Check(CStr(t.IsCanceled), "True")
    End Sub
End Module
