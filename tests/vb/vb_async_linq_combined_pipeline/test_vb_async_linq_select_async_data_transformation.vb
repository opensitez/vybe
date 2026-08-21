' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_select_async_data_transformation
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

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks
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


Module Program
    Private Async Function FetchDataAsync(id As Integer) As Task(Of String)
        Await Task.Yield()
        Return "Item_" & id
    End Function

    Sub Main()
        Dim ids As Integer() = {1, 2, 3}
        Dim tasks = ids.Select(Function(i) FetchDataAsync(i)).ToArray()
        Task.WaitAll(tasks)

        Dim results = tasks.Select(Function(t) t.Result).ToList()
        __P(CStr(String.Join(",", results)))
        __Check("Item_1,Item_2,Item_3")
    End Sub
End Module
