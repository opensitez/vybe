' vybe-test: vb/vb_system_floating_point_matrix/floating_point_matrix_rounding_and_precision_contracts
' origin: languages/vb/tests/vb/test_vb_system_floating_point_matrix.rs

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

Imports System

Module M
    Sub Main()
        Dim values() As Double = {3.5, -3.5, 2.1, -2.1, 0.0, 1.499, 2.5}

        __P(CStr(Math.Round(3.5, 0)))
        __P(CStr(Math.Round(2.5, 0)))
        __P(CStr(Math.Ceiling(-3.5)))
        __P(CStr(Math.Floor(2.5)))
        __P(CStr(Math.Truncate(values(0))))
        __P(CStr(Math.Truncate(values(1))))
        __P(CStr(Math.Sign(values(2))))
        __P(CStr(Math.Sign(values(5))))
        __Check("4
2
-3
2
3
-3
1
1")
    End Sub
End Module
