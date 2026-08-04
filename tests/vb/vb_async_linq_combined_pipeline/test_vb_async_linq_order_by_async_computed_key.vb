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
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
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
        __P(CStr(String.Join(",", sortedWords)))
        __Check("Cat,Giraffe,Elephant")
    End Sub
End Module
