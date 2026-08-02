' vybe-test: vb/vb_async_task_when_all_any/test_vb_async_task_when_all
' origin: languages/vb/tests/vb/test_vb_async_task_when_all_any.rs

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
    Async Function FetchOne() As Task(Of String)
        Await Task.Delay(5)
        Return "One"
    End Function

    Async Function FetchTwo() As Task(Of String)
        Await Task.Delay(5)
        Return "Two"
    End Function

    Async Function RunAllAsync() As Task
        Dim results As String() = Await Task.WhenAll(FetchOne(), FetchTwo())
        __Check(CStr(String.Join(",", results)), "One,Two")
    End Function

    Sub Main()
        RunAllAsync().Wait()
    End Sub
End Module
