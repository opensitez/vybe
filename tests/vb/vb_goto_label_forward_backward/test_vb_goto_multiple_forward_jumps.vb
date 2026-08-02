' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_multiple_forward_jumps
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
    Sub Main()
        GoTo L1
L2:
        __Check(CStr("Step 2"), "Step 1")
        GoTo L3
L1:
        __Check(CStr("Step 1"), "Step 2")
        GoTo L2
L3:
        __Check(CStr("Step 3"), "Step 3")
    End Sub
End Module
