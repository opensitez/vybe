' vybe-test: vb/vb_operator_overloading_custom/operator_overloading_custom
' origin: languages/vb/tests/vb/test_vb_operator_overloading_custom.rs

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
    Public X As Integer
    Public Y As Integer
    
    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        Return New Vector() With {.X = v1.X + v2.X, .Y = v1.Y + v2.Y}
    End Operator
    
    Public Shared Operator -(v1 As Vector) As Vector
        Return New Vector() With {.X = -v1.X, .Y = -v1.Y}
    End Operator
End Class

Module M
    Sub Main()
        Dim v1 As New Vector() With {.X = 1, .Y = 2}
        Dim v2 As New Vector() With {.X = 3, .Y = 4}
        
        Dim v3 = v1 + v2
        Dim v4 = -v1
        
        __P(CStr(v3.X))
        __P(CStr(v4.X))
        __Check("4
-1")
    End Sub
End Module
