' vybe-test: vb/vb_spec_environment_interaction/interaction_spec_beep_returns_to_program_flow
' origin: languages/vb/tests/vb/test_vb_spec_environment_interaction.rs

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
        __Check(CStr(Module Program
    Sub Main()
        Beep()
        Console.WriteLine("after-beep")
    End Sub
End Module), "after-beep")
    End Sub
End Module
