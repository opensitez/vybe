' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_group_by_pipeline
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

Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Class Metric
        Public Category As String
        Public Value As Integer
    End Class

    Private Async Function GetMetricsAsync() As Task(Of List(Of Metric))
        Await Task.Yield()
        Return New List(Of Metric) From {
            New Metric With {.Category = "CPU", .Value = 50},
            New Metric With {.Category = "RAM", .Value = 70},
            New Metric With {.Category = "CPU", .Value = 60}
        }
    End Function

    Sub Main()
        Dim t = GetMetricsAsync()
        t.Wait()

        Dim grouped = t.Result.GroupBy(Function(m) m.Category)
        For Each g In grouped.OrderBy(Function(g) g.Key)
            __P(CStr(g.Key & ":" & g.Average(Function(m) m.Value)))
        Next
        __Check("CPU:55
RAM:70")
    End Sub
End Module
