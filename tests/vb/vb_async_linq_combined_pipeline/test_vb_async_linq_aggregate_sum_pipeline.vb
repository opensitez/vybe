' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_aggregate_sum_pipeline
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
    Private Async Function ComputeValueAsync(x As Integer) As Task(Of Integer)
        Await Task.Yield()
        Return x * 10
    End Function

    Sub Main()
        Dim inputs As Integer() = {1, 2, 3, 4}
        Dim tasks = inputs.Select(Function(x) ComputeValueAsync(x)).ToArray()
        Task.WaitAll(tasks)

        Dim totalSum = tasks.Sum(Function(t) t.Result)
        __Check(CStr(totalSum), "100")
    End Sub
End Module
