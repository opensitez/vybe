' vybe-test: vb/vb_forms/d04_handler_is_callable

' RECONSTRUCTED. The extractor wrote this file's expected VALUE out as its
' source — the whole program was the single word `btn1`, which parses as a
' bare identifier and fails with `undefined is not callable`. The original
' `languages/vb/tests/vb/vb_forms_test.rs` no longer exists, so the assertion
' below is authored from the test's NAME and its recovered expected value.
'
' `AddressOf` on an instance method is the thing under test: the delegate must
' carry its receiver, so calling it has to reach the SAME object's field.
'
' Verified on BOTH runtimes — `tools/vbrun` runs it under real VB.NET, so the
' expectation is Microsoft's and not ours. Hence `Sub Main` (top-level
' statements are a vybe extension; real VB answers BC30689) and the explicit
' `Action(Of String)` (VB cannot infer a delegate type from `AddressOf`).

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

Public Class Form1
    Public Clicked As String = ""

    Public Sub Btn1_Click(sender As String)
        Clicked = sender
    End Sub
End Class

Module Program
    Sub Main()
        Dim f As New Form1()
        Dim handler As Action(Of String) = AddressOf f.Btn1_Click
        handler("btn1")
        __P(CStr(f.Clicked))
        __Check("btn1")
    End Sub
End Module
