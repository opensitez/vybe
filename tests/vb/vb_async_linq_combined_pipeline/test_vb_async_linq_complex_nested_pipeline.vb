' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_complex_nested_pipeline
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
    Private Async Function TransformRecordAsync(val As Integer) As Task(Of Integer)
        Await Task.Yield()
        Return val * 3
    End Function

    Sub Main()
        Dim input As Integer() = {1, 2, 3, 4, 5}
        Dim pipelineTask = Task.Run(Async Function()
            ' Filter even -> Multiply by 3 -> Sum
            Dim evens = input.Where(Function(n) n Mod 2 = 0)
            Dim tasks = evens.Select(Function(n) TransformRecordAsync(n)).ToArray()
            Await Task.WhenAll(tasks)
            Return tasks.Sum(Function(t) t.Result)
        End Function)

        pipelineTask.Wait()
        __Check(CStr(pipelineTask.Result), "18")
    End Sub
End Module
