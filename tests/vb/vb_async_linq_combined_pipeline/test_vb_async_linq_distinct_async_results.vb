' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_distinct_async_results
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
    Private Async Function NormalizeAsync(s As String) As Task(Of String)
        Await Task.Yield()
        Return s.Trim().ToUpper()
    End Function

    Sub Main()
        Dim raw As String() = {"apple", " Apple ", "APPLE", "banana"}
        Dim tasks = raw.Select(Function(s) NormalizeAsync(s)).ToArray()
        Task.WaitAll(tasks)

        Dim unique = tasks.Select(Function(t) t.Result).Distinct().OrderBy(Function(x) x)
        __Check(CStr(String.Join(",", unique)), "APPLE,BANANA")
    End Sub
End Module
