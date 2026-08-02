' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_take_while_condition
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
    Private Async Function IsUnderLimitAsync(val As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return val < 100
    End Function

    Sub Main()
        Dim items As Integer() = {10, 50, 120, 30}
        Dim tasks = items.Select(Async Function(x) New With {.Val = x, .Valid = Await IsUnderLimitAsync(x)}).ToArray()
        Task.WaitAll(tasks)

        Dim validSequence = tasks.Select(Function(t) t.Result).TakeWhile(Function(x) x.Valid).Select(Function(x) x.Val)
        __Check(CStr(String.Join(",", validSequence)), "10,50")
    End Sub
End Module
