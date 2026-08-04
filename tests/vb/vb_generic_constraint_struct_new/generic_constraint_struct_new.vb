' vybe-test: vb/vb_generic_constraint_struct_new/generic_constraint_struct_new
' origin: languages/vb/tests/vb/test_vb_generic_constraint_struct_new.rs

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

' As Structure requires T to be a value type
Class ValueCache(Of T As Structure)
    Public Property Item As T
End Class

' As New requires T to have a parameterless constructor
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class Person
    Public Property Name As String = "Bob"
End Class

Module M
    Sub Main()
        Dim vc As New ValueCache(Of Integer)()
        vc.Item = 42
        __P(CStr(vc.Item))
        
        Dim f As New Factory(Of Person)()
        Dim p = f.Create()
        __P(CStr(p.Name))
        __Check("42
Bob")
    End Sub
End Module
