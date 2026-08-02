' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_where_filter_pipeline
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

Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function IsValidAsync(num As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return num Mod 2 = 0
    End Function

    Sub Main()
        Dim numbers As Integer() = {10, 15, 20, 25, 30}
        ' Filter asynchronously
        Dim tasks = numbers.Select(Async Function(n) New With {.Val = n, .Keep = Await IsValidAsync(n)}).ToArray()
        Task.WaitAll(tasks)

        Dim evens = tasks.Where(Function(x) x.Result.Keep).Select(Function(x) x.Result.Val)
        __Check(CStr(String.Join(",", evens)), "10,20,30")
    End Sub
End Module
