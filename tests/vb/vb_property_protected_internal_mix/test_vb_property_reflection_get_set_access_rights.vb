' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_reflection_get_set_access_rights
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class Sample
    ' ⛔ `{ Get; Private Set; }` is C# AUTO-PROPERTY syntax — VB has no
    ' brace form, and an auto-property cannot carry a per-accessor access
    ' modifier. The expanded property with a backing field is VB's spelling.
    Private _text As String = "Init"
    Public Property Text As String
        Get
            Return _text
        End Get
        Private Set(value As String)
            _text = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Sample).GetProperty("Text")
        __P(CStr((prop.GetMethod IsNot Nothing) & "|" & (prop.SetMethod IsNot Nothing)))
        __Check("True|True")
    End Sub
End Module
