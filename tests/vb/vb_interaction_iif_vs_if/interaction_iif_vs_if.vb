' vybe-test: vb/vb_interaction_iif_vs_if/interaction_iif_vs_if
' origin: languages/vb/tests/vb/test_vb_interaction_iif_vs_if.rs

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
        Dim condition As Boolean = True
        
        ' IIf is a legacy function that evaluates BOTH true and false arguments
        ' (Not short-circuited!)
        Dim result1 = IIf(condition, "Yes", "No")
        __P(CStr(result1))
        
        ' If operator is short-circuited and type-safe
        Dim result2 = If(condition, "Yes", "No")
        __P(CStr(result2))
        
        ' If operator with two arguments acts like coalesce (expr1 ?? expr2)
        Dim val1 As String = Nothing
        Dim val2 As String = "Default"
        __P(CStr(If(val1, val2)))
        __Check("Yes
Yes
Default")
    End Sub
End Module
