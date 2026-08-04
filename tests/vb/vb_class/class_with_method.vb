' vybe-test: vb/vb_class/class_with_method
' origin: languages/vb/tests/vb/vb_class_test.rs

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
    Class Calculator
        Public Result As Double

        Sub New()
            Me.Result = 0
        End Sub

        Function Add(a As Double, b As Double) As Double
            Add = a + b
        End Function

        Sub AddToResult(value As Double)
            Me.Result = Me.Result + value
        End Sub
    End Class

    Sub Main()
        Dim calc As New Calculator()
        __P(CStr(calc.Add(3, 4)))
        calc.AddToResult(10)
        calc.AddToResult(20)
        __P(CStr(calc.Result))
        __Check("7
30")
    End Sub
End Module
