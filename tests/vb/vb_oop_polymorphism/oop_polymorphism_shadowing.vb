' vybe-test: vb/vb_oop_polymorphism/oop_polymorphism_shadowing
' origin: languages/vb/tests/vb/test_vb_oop_polymorphism.rs

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

Class Base
    Public Overridable Function GetName() As String
        Return "Base"
    End Function
End Class

Class Derived1
    Inherits Base
    Public Overrides Function GetName() As String
        Return "Derived1"
    End Function
End Class

Class Derived2
    Inherits Base
    ' Shadows the base method, doesn't override
    Public Shadows Function GetName() As String
        Return "Derived2"
    End Function
End Class

Module M
    Sub Main()
        Dim d1 As New Derived1()
        Dim d2 As New Derived2()
        
        Dim b1 As Base = d1
        Dim b2 As Base = d2
        
        __P(CStr(b1.GetName()))
        __P(CStr(b2.GetName()))
        __P(CStr(d2.GetName()))
        __Check("Derived1
Base
Derived2")
    End Sub
End Module
