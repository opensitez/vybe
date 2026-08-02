' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_backward_counter_accumulation
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
        Dim total = 0
        Dim i = 1
Accumulate:
        total += i
        i += 1
        If i <= 5 Then GoTo Accumulate
        __Check(CStr(total), "15")
    End Sub
End Module
