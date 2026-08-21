' vybe-test: vb/vb_overloads_statement/statement_overloads
' origin: languages/vb/tests/vb/test_vb_overloads_statement.rs

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
    Public Sub Process(x As Integer)
        __P(CStr("Process Integer: " & x))
    End Sub
End Class

Class Derived
    Inherits Base

    ' Overloads is technically optional when the signatures are different,
    ' but it's used to explicitly define overloaded methods across inheritance bounds
    Public Overloads Sub Process(x As String)
    __P(CStr("Process String: " & x))
End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.Process(10)
        d.Process("Hello")
        __Check("Process Integer: 10
Process String: Hello")
    End Sub
End Module
