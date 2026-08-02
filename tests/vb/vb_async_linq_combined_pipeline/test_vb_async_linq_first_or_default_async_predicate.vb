' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_first_or_default_async_predicate
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
    Private Async Function CheckMatchAsync(s As String) As Task(Of Boolean)
        Await Task.Yield()
        Return s.StartsWith("B")
    End Function

    Sub Main()
        Dim items As String() = {"Apple", "Banana", "Cherry"}
        Dim tasks = items.Select(Async Function(item) New With {.Text = item, .IsMatch = Await CheckMatchAsync(item)}).ToArray()
        Task.WaitAll(tasks)

        Dim firstMatch = tasks.FirstOrDefault(Function(t) t.Result.IsMatch)
        __Check(CStr(If(firstMatch IsNot Nothing, firstMatch.Result.Text, "None")), "Banana")
    End Sub
End Module
