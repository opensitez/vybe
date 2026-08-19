' vybe-test: vb/vb_modules_namespaces/ns_module_members_are_namespace_members
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


' A MODULE's members have namespace-level accessibility: `N1.DoubleIt` names
' the same function as `N1.Mod1.DoubleIt`, with the module name skipped. This
' holds with NO import at all — it is a property of the declaration, not of a
' `Imports` directive, which is what separates it from
' `ns_imports_makes_module_members_unqualified`.
'
' Measured against real VB.NET (dotnet SDK, `dotnet new console -lang VB`):
'   N1.DoubleIt(10)      -> 20   with no Imports
'   N1.Mod1.DoubleIt(5)  -> 10   with no Imports
'   DoubleIt(21)         -> BC30451 'DoubleIt' is not declared, with no Imports
Namespace N1
Public Module Mod1
Public Function DoubleIt(v As Integer) As Integer
Return v * 2
End Function
End Module
End Namespace

Module M
Sub Main()
__P(CStr(N1.DoubleIt(10)))
__P(CStr(N1.Mod1.DoubleIt(5)))
__Check("20" & vbLf & "10")
End Sub
End Module
