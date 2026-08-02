' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_safe_long_addition
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

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
        Dim a As Long = 5000000000L
        Dim b As Long = 5000000000L
        Dim c As Long = a + b
        __Check(CStr(c), "10000000000")
    End Sub
End Module
