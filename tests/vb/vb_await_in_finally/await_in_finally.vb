' vybe-test: vb/vb_await_in_finally/await_in_finally
' origin: languages/vb/tests/vb/test_vb_await_in_finally.rs

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
    Async Function CleanupAsync() As Task
        __Check(CStr("Cleaned"), "Cleaned")
    End Function

    Async Function TestAsync() As Task
        Try
            ' do nothing
        Finally
            ' Await inside Finally (added in VB 14)
            Await CleanupAsync()
        End Try
    End Function

    Sub Main()
        TestAsync().Wait()
    End Sub
End Module
