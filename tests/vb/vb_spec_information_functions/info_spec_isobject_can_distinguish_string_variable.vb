' vybe-test: vb/vb_spec_information_functions/info_spec_isobject_can_distinguish_string_variable
' origin: languages/vb/tests/vb/test_vb_spec_information_functions.rs

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

Module M
    Sub Main()
        Dim value As String = "vb"
        ' `IsObject` was dropped from Microsoft.VisualBasic in .NET — even an
        ' explicit `Imports Microsoft.VisualBasic.Information` answers BC30451.
        ' The .NET spelling of "is this a reference, not a value" is the runtime
        ' type test below.
        '
        ' ⛔THE ANSWER CHANGED, deliberately. VB6's `IsObject` said False for a
        ' String; .NET says a String IS a reference type, so this now reads
        ' True. The expectation was re-derived from real VB.NET rather than
        ' carried over — the test asserts .NET semantics now, not VB6's.
        __P(CStr(Not CObj(value).GetType().IsValueType))
        __Check("True")
    End Sub
End Module
