' vybe-test: vb/vb_istrue_isfalse_operators/istrue_isfalse_operators
' origin: languages/vb/tests/vb/test_vb_istrue_isfalse_operators.rs

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

Class Truthy
    Public Value As Integer
    
    Public Shared Operator IsTrue(ByVal obj As Truthy) As Boolean
        Return obj.Value > 0
    End Operator
    
    Public Shared Operator IsFalse(ByVal obj As Truthy) As Boolean
        Return obj.Value <= 0
    End Operator
End Class

Module M
    Sub Main()
        Dim t1 As New Truthy() With {.Value = 10}
        Dim t2 As New Truthy() With {.Value = -5}
        
        If t1 Then
            __P(CStr("t1 is true"))
        End If
        
        If Not t2 Then
            __P(CStr("t2 is false"))
        End If
        __Check("t1 is true
t2 is false")
    End Sub
End Module
