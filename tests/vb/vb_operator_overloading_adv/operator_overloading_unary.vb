' vybe-test: vb/vb_operator_overloading_adv/operator_overloading_unary
' origin: languages/vb/tests/vb/test_vb_operator_overloading_adv.rs

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

Class Vector
    Public X, Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    ' Unary operator -
    Public Shared Operator -(v As Vector) As Vector
        Return New Vector(-v.X, -v.Y)
    End Operator
    
    ' Unary operator Not
    Public Shared Operator Not(v As Vector) As Vector
        Return New Vector(Not v.X, Not v.Y)
    End Operator
End Class

Module M
    Sub Main()
        Dim v As New Vector(5, -10)
        Dim vNeg = -v
        __P(CStr(vNeg.X))
        __P(CStr(vNeg.Y))
        
        Dim vNot = Not v
        __P(CStr(vNot.X)) ' Not 5 = -6
        __Check("-5
10
-6")
    End Sub
End Module
