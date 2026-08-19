' vybe-test: vb/vb_modules_namespaces/ns_imports_makes_module_members_unqualified
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


' `Imports N1` makes the members of every MODULE in N1 callable with no
' qualifier at all — a module is a named container whose contents belong to the
' enclosing namespace, unlike a class, whose members need an instance.
'
' The sibling `mod_extension_methods_require_imports` reaches DoubleIt as an
' EXTENSION method (`x.DoubleIt()`), which resolves through the receiver and so
' never asks this question. The bare-call route is the one asserted here.
Namespace N1
Public Module Mod1
Public Function DoubleIt(v As Integer) As Integer
Return v * 2
End Function
End Module
End Namespace

Imports N1
Module M
Sub Main()
__P(CStr(DoubleIt(21)))
__Check("42")
End Sub
End Module
