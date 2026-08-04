' vybe-test: vb/vb_exit_statement_variants/exit_statement_variants
' origin: languages/vb/tests/vb/test_vb_exit_statement_variants.rs

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
    Sub TestExitSub()
        __P(CStr("Sub1"))
        Exit Sub
        __P(CStr("Sub2"))
    End Sub

    Function TestExitFunction() As Integer
        TestExitFunction = 10
        Exit Function
        TestExitFunction = 20
    End Function

    Sub Main()
        TestExitSub()
        __P(CStr(TestExitFunction()))
        
        For i = 1 To 5
            If i = 3 Then Exit For
            __P(CStr("For " & i))
        Next
        
        Dim j = 1
        Do While j <= 5
            If j = 2 Then Exit Do
            __P(CStr("Do " & j))
            j += 1
        Loop
        __Check("Sub1
10
For 1
For 2
Do 1")
    End Sub
End Module
