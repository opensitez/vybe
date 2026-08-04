' vybe-test: vb/vb_abstract_class_inheritance_chain/test_vb_mustinherit_concrete_base_methods
' origin: languages/vb/tests/vb/test_vb_abstract_class_inheritance_chain.rs

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

MustInherit Class BaseLogger
    Public Sub Log(msg As String)
        WriteEntry(FormatMessage(msg))
    End Sub

    Protected MustOverride Sub WriteEntry(formatted As String)

    Protected Virtual Function FormatMessage(msg As String) As String
        Return "[LOG] " & msg
    End Function
End Class

Class ConsoleLogger
    Inherits BaseLogger
    Protected Overrides Sub WriteEntry(formatted As String)
        __P(CStr(formatted))
    End Sub
End Class

Module Program
    Sub Main()
        Dim logger As BaseLogger = New ConsoleLogger()
        logger.Log("System initialized")
        __Check("[LOG] System initialized")
    End Sub
End Module
