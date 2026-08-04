' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_task_when_any_first_responder
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
    Private Async Function SlowTaskAsync() As Task(Of String)
        Await Task.Delay(50)
        Return "Slow"
    End Function

    Private Async Function FastTaskAsync() As Task(Of String)
        Await Task.Yield()
        Return "Fast"
    End Function

    Sub Main()
        Dim tSlow = SlowTaskAsync()
        Dim tFast = FastTaskAsync()

        Dim winner = Task.WhenAny(tSlow, tFast)
        winner.Wait()
        __P(CStr(winner.Result.Result))
        __Check("Fast")
    End Sub
End Module
