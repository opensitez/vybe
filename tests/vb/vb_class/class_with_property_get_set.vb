' vybe-test: vb/vb_class/class_with_property_get_set
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
    Class Temperature
        Private _celsius As Double

        Sub New(c As Double)
            _celsius = c
        End Sub

        Property Celsius() As Double
            Get
                Return _celsius
            End Get
            Set(value As Double)
                _celsius = value
            End Set
        End Property

        Property Fahrenheit() As Double
            Get
                Return _celsius * 9 / 5 + 32
            End Get
            Set(value As Double)
                _celsius = (value - 32) * 5 / 9
            End Set
        End Property
    End Class

    Sub Main()
        Dim t As New Temperature(100)
        __P(CStr(t.Celsius))
        __P(CStr(t.Fahrenheit))
        t.Fahrenheit = 32
        __P(CStr(t.Celsius))
        __Check("100
212
0")
    End Sub
End Module
