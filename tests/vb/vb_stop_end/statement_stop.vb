' vybe-test: vb/vb_stop_end/statement_stop
' origin: languages/vb/tests/vb/test_vb_stop_end.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Module M
    Sub Main()
        __Check(CStr("Before Stop"), "Before Stop")
        ' Stop suspends execution (in a debugger), but compiler supports parsing it.
        ' Without a debugger attached, behavior varies, but we check parsing/compilation.
        ' Since we don't want to actually suspend our test runner, we will just parse it
        ' but not hit it, or let it compile. Wait, Stop sometimes terminates or throws.
        ' Let's just put it in a non-executed block to verify parser.
        If False Then
            Stop
        End If
        __Check(CStr("After Stop"), "After Stop")
    End Sub
End Module
