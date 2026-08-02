' vybe-test: vb/vb_async_catch_finally/async_await_in_catch_finally
' origin: languages/vb/tests/vb/test_vb_async_catch_finally.rs

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

Module M
    Async Function LogErrorAsync() As Task
        Await Task.Delay(10)
        __Check(CStr("Error Logged Async"), "Error Logged Async")
    End Function

    Async Function CleanupAsync() As Task
        Await Task.Delay(10)
        __Check(CStr("Cleaned Up Async"), "Cleaned Up Async")
    End Function

    Async Function DoWorkAsync() As Task
        Try
            Throw New Exception("Fail")
        Catch ex As Exception
            Await LogErrorAsync()
        Finally
            Await CleanupAsync()
        End Try
    End Function

    Sub Main()
        DoWorkAsync().Wait()
    End Sub
End Module
