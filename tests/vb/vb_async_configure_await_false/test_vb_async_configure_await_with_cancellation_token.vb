' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_with_cancellation_token
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

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
    Private Async Function CancellableConfiguredAsync(cts As CancellationTokenSource) As Task(Of String)
        Await Task.Delay(10, cts.Token).ConfigureAwait(False)
        Return "Success"
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim t = CancellableConfiguredAsync(cts)
        __Check(CStr(t.Result), "Success")
    End Sub
End Module
