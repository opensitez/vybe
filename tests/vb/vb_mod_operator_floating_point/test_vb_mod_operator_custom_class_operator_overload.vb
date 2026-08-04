' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_custom_class_operator_overload
' origin: languages/vb/tests/vb/test_vb_mod_operator_floating_point.rs

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

Module Program
    Class ClockTime
        Public Hours As Integer
        Public Sub New(h As Integer)
            Hours = h
        End Sub
        Public Shared Operator Mod(a As ClockTime, b As Integer) As ClockTime
            Return New ClockTime(a.Hours Mod b)
        End Operator
    End Class

    Sub Main()
        Dim t As New ClockTime(27)
        Dim wrapped = t Mod 24
        __P(CStr(wrapped.Hours))
        __Check("3")
    End Sub
End Module
