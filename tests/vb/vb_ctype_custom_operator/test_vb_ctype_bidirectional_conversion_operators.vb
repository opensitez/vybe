' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_bidirectional_conversion_operators
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Class Celsius
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub
End Class

Class Fahrenheit
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub

    Public Shared Widening Operator CType(c As Celsius) As Fahrenheit
        Return New Fahrenheit(c.Degrees * 9.0 / 5.0 + 32.0)
    End Shared Widening Operator

    Public Shared Widening Operator CType(f As Fahrenheit) As Celsius
        Return New Celsius((f.Degrees - 32.0) * 5.0 / 9.0)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim c As New Celsius(100)
        Dim f As Fahrenheit = CType(c, Fahrenheit)
        Dim restoredC As Celsius = CType(f, Celsius)
        __P(CStr(f.Degrees & "|" & restoredC.Degrees))
        __Check("212|100")
    End Sub
End Module
