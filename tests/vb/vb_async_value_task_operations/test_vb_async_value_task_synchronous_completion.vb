' vybe-test: vb/vb_async_value_task_operations/test_vb_async_value_task_synchronous_completion
' origin: languages/vb/tests/vb/test_vb_async_value_task_operations.rs

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

Imports System.Threading.Tasks

Module Program
    Function GetCachedValueAsync(cached As Boolean) As ValueTask(Of Integer)
        If cached Then
            Return New ValueTask(Of Integer)(100)
        End If
        Return New ValueTask(Of Integer)(ComputeAsync())
    End Function

    Async Function ComputeAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 200
    End Function

    Async Function RunAsync() As Task
        Dim v1 As Integer = Await GetCachedValueAsync(True)
        Dim v2 As Integer = Await GetCachedValueAsync(False)
        __P(CStr(v1 & ":" & v2))
    End Function

    Sub Main()
        RunAsync().Wait()
        __Check("100:200")
    End Sub
End Module
