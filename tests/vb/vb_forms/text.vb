' vybe-test: vb/vb_forms/text

' RECONSTRUCTED. The extractor wrote this file's expected VALUE out as its
' source — the whole program was the two words `Click Me`, which fails with
' `undefined is not callable`. The original
' `languages/vb/tests/vb/vb_forms_test.rs` no longer exists, so the assertion
' below is authored from the test's NAME and its recovered expected value.
'
' Verified on BOTH runtimes — `tools/vbrun` runs it under real VB.NET, so the
' expectation is Microsoft's and not ours. That is why the work happens in
' `Sub Main` rather than at file scope: top-level statements are a vybe
' extension and real VB answers BC30689 to them.

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

Public Class Button1
    Public Text As String = ""
End Class

Module Program
    Sub Main()
        Dim b As New Button1()
        b.Text = "Click Me"
        __P(CStr(b.Text))
        __Check("Click Me")
    End Sub
End Module
