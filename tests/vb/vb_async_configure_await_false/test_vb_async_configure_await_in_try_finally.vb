' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_in_try_finally
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

Imports System.Threading.Tasks

Module Program
    Private Async Function ResourceAsync() As Task(Of String)
        Dim res = "Opened"
        Try
            Await Task.Delay(5).ConfigureAwait(False)
            Return res & " -> Processed"
        Finally
            __Check(CStr("Finally Block Executed"), "Finally Block Executed")
        End Try
    End Function

    Sub Main()
        Dim t = ResourceAsync()
        __Check(CStr(t.Result), "Opened -> Processed")
    End Sub
End Module
