' vybe-test: vb/vb_async_await_syntax/async_await_syntax
' origin: languages/vb/tests/vb/test_vb_async_await_syntax.rs

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
    Async Function GetDataAsync() As Task(Of Integer)
        Await Task.Delay(1)
        Return 42
    End Function

    Sub Main()
        ' Using Wait() for console app simplicity, testing Async/Await compilation
        Dim task = GetDataAsync()
        task.Wait()
        __Check(CStr(task.Result), "42")
    End Sub
End Module
