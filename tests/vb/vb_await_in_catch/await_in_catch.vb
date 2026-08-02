' vybe-test: vb/vb_await_in_catch/await_in_catch
' origin: languages/vb/tests/vb/test_vb_await_in_catch.rs

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
        __Check(CStr("Logged"), "Logged")
    End Function

    Async Function TestAsync() As Task
        Try
            Throw New System.Exception()
        Catch ex As System.Exception
            ' Await inside Catch (added in VB 14)
            Await LogErrorAsync()
        End Try
    End Function

    Sub Main()
        TestAsync().Wait()
    End Sub
End Module
