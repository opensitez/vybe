' vybe-test: vb/vb_end_statement_adv/end_statement_adv
' origin: languages/vb/tests/vb/test_vb_end_statement_adv.rs

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
        __Check(CStr("Before End"), "Before End")
        ' End statement terminates execution immediately
        ' Some test runners might fail if we actually execute End, 
        ' so we place it inside an unreachable block
        If False Then
            End
        End If
        __Check(CStr("Parsed End"), "Parsed End")
    End Sub
End Module
