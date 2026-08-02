' vybe-test: vb/vb_async_cancellation_token/test_vb_async_cancellation_token_source_cancel_after
' origin: languages/vb/tests/vb/test_vb_async_cancellation_token.rs

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

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.CancelAfter(50)
        __Check(CStr(cts.IsCancellationRequested), "False")
        Thread.Sleep(100)
        __Check(CStr(cts.IsCancellationRequested), "True")
    End Sub
End Module
