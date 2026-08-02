' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_order_by_async_computed_key
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
    Private Async Function ComputeWeightAsync(s As String) As Task(Of Integer)
        Await Task.Yield()
        Return s.Length
    End Function

    Sub Main()
        Dim words As String() = {"Elephant", "Cat", "Giraffe"}
        Dim tasks = words.Select(Async Function(w) New With {.Word = w, .Weight = Await ComputeWeightAsync(w)}).ToArray()
        Task.WaitAll(tasks)

        Dim sortedWords = tasks.Select(Function(t) t.Result).OrderBy(Function(x) x.Weight).Select(Function(x) x.Word)
        __Check(CStr(String.Join(",", sortedWords)), "Cat,Giraffe,Elephant")
    End Sub
End Module
