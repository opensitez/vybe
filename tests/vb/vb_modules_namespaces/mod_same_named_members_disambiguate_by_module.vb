' vybe-test: vb/vb_modules_namespaces/mod_same_named_members_disambiguate_by_module
' origin: languages/vb/tests/vb/test_vb_modules_namespaces.rs

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


' Two modules may declare the SAME member name. Which one an unqualified
' reference means is decided in tiers, and both tiers below are asserted here.
'
' Measured against real VB.NET (dotnet SDK, `dotnet new console -lang VB`):
'   A.Same() / B.Same()          -> "A" / "B"   qualified, always legal
'   Same() inside Module A       -> "A"         the CONTAINING module wins
'   Same() from a third module   -> BC30562 'Same' is ambiguous between
'                                   declarations in Modules 'A, B'
'
' The BC30562 case is deliberately NOT exercised: this suite has no
' expect-compile-error mode, so a test for it could only assert the wrong
' thing. The two tiers that DO have an observable answer are asserted, and
' they are the ones that must keep working when the ambiguous case starts
' being rejected.
Module A
Public Function Same() As String
Return "A"
End Function
Public Function FromInsideA() As String
Return Same()
End Function
End Module

Module B
Public Function Same() As String
Return "B"
End Function
End Module

Module M
Sub Main()
__P(A.Same())
__P(B.Same())
__P(A.FromInsideA())
__Check("A" & vbLf & "B" & vbLf & "A")
End Sub
End Module
