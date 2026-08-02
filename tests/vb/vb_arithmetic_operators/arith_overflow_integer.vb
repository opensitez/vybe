' vybe-test: vb/vb_arithmetic_operators/arith_overflow_integer
' origin: languages/vb/tests/vb/test_vb_arithmetic_operators.rs

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
Dim x = 2147483647
Try
Dim z = x + 1
Catch
__Check(CStr("Caught"), "Caught")
End Try
End Sub
End Module
