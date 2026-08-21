' vybe-test: vb/vb_forms/h02_msgbox_with_title
' vybe-test-mode: compile

' RECONSTRUCTED. The extractor wrote this file's expected VALUE out as its
' source — the whole program was the three words `Are you sure?`, which is not
' VB and failed to parse. The original
' `languages/vb/tests/vb/vb_forms_test.rs` no longer exists, so the program
' below is authored from the test's NAME and its recovered message.
'
' ⛔COMPILE MODE, deliberately. `MsgBox` lowers to `web:window.alert`, and
' `alert` is MODAL: running this opens a real dialog on the desktop and blocks
' until someone dismisses it. Verified by doing exactly that — the dialog
' appeared, which is the feature working, and is also precisely why this cannot
' be a `Run` test. Compile mode asserts what is assertable without a human: that
' the three-argument `MsgBox(prompt, style, title)` form parses and lowers.
'
' REAL VB AGREES, and was asked: `tools/vbrun` compiles this file with no `BC`
' diagnostic — so the three-argument shape is Microsoft's, not our invention —
' and then dies at RUNTIME with
' `System.PlatformNotSupportedException: Method requires System.Windows.Forms`.
' Neither runtime can execute this headlessly on macOS, which is why the
' assertion stops at compilation. Note the asymmetry: vybe SHOWS this dialog on
' macOS where real .NET cannot, because our `MsgBox` reaches `web:window.alert`
' on our own widget stack instead of WinForms.
'
' The TITLE is accepted and then dropped, which is correct and not a gap: HTML
' §8.6 gives `alert()` a message and no title, because the user agent owns the
' dialog's chrome so a page cannot impersonate a system dialog.

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

Module Program
    Sub Main()
        MsgBox("Are you sure?", MsgBoxStyle.OkOnly, "Confirm")
        __P(CStr("shown"))
        __Check("shown")
    End Sub
End Module
