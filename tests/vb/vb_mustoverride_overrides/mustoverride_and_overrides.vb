' vybe-test: vb/vb_mustoverride_overrides/mustoverride_and_overrides
' origin: languages/vb/tests/vb/test_vb_mustoverride_overrides.rs

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

MustInherit Class Shape
    Public MustOverride Function GetArea() As Double
    Public MustOverride Property Name As String
End Class

Class Circle
    Inherits Shape
    
    Private _name As String = "Circle"
    Private _radius As Double
    
    Public Sub New(radius As Double)
        _radius = radius
    End Sub
    
    Public Overrides Function GetArea() As Double
        Return Math.PI * _radius * _radius
    End Function
    
    Public Overrides Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim s As Shape = New Circle(10)
        __P(CStr(s.Name))
        __P(CStr(Math.Round(s.GetArea())))
        __Check("Circle
314")
    End Sub
End Module
