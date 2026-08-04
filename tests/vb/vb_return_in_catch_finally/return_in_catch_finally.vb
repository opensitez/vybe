' vybe-test: vb/vb_return_in_catch_finally/return_in_catch_finally
' origin: languages/vb/tests/vb/test_vb_return_in_catch_finally.rs

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
    Function TestReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            Return 1
        Finally
            ' VB.NET allows Return in Finally?
            ' Wait, Return in Finally is a compiler error in VB.NET (and C#)!
            ' So we just test Return in Catch and modifying the return value implicitly by assigning to the function name
            __P(CStr("Finally executed"))
        End Try
    End Function

    Function TestImplicitReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            TestImplicitReturn = 2
            Exit Function
        Finally
            __P(CStr("Finally executed 2"))
        End Try
    End Function

    Sub Main()
        __P(CStr(TestReturn()))
        __P(CStr(TestImplicitReturn()))
        __Check("Finally executed
1
Finally executed 2
2")
    End Sub
End Module
