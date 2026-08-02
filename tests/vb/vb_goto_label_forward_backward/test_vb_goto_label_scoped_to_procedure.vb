' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_label_scoped_to_procedure
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

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

Module Program
    Private Sub ProcA()
        GoTo CommonLabel
CommonLabel:
        __Check(CStr("ProcA Label"), "ProcA Label")
    End Sub

    Private Sub ProcB()
        GoTo CommonLabel
CommonLabel:
        __Check(CStr("ProcB Label"), "ProcB Label")
    End Sub

    Sub Main()
        ProcA()
        ProcB()
    End Sub
End Module
