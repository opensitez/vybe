' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_all_any_async_predicates
' origin: languages/vb/tests/vb/test_vb_async_linq_combined_pipeline.rs

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

Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function IsPositiveAsync(n As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return n > 0
    End Function

    Sub Main()
        Dim numbers As Integer() = {1, 5, 10}
        Dim tasks = numbers.Select(Async Function(n) Await IsPositiveAsync(n)).ToArray()
        Task.WaitAll(tasks)

        Dim allPositive = tasks.All(Function(t) t.Result)
        __Check(CStr(allPositive), "True")
    End Sub
End Module
