' vybe-test: vb/vb_casts_advanced/casts_trycast
' origin: languages/vb/tests/vb/test_vb_casts_advanced.rs

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

Class Animal
End Class

Class Dog
    Inherits Animal
    Public Sub Bark()
        __P(CStr("Woof"))
    End Sub
End Class

Module M
    Sub Main()
        Dim a1 As Animal = New Dog()
        Dim a2 As Animal = New Animal()
        
        ' TryCast returns Nothing if the cast fails (only for reference types)
        Dim d1 As Dog = TryCast(a1, Dog)
        Dim d2 As Dog = TryCast(a2, Dog)
        
        If d1 IsNot Nothing Then d1.Bark()
        __P(CStr(d2 Is Nothing))
        __Check("Woof
True")
    End Sub
End Module
