' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_nested_methods
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
    Private Async Function InnerAsync() As Task(Of String)
        Await Task.Delay(5).ConfigureAwait(False)
        Return "Inner"
    End Function

    Private Async Function OuterAsync() As Task(Of String)
        Dim res = Await InnerAsync().ConfigureAwait(False)
        Return "Outer -> " & res
    End Function

    Sub Main()
        Dim t = OuterAsync()
        __Check(CStr(t.Result), "Outer -> Inner")
    End Sub
End Module
