' vybe-test: vb/vb_sub_func_returns/sub_recursive
' origin: languages/vb/tests/vb/test_vb_sub_func_returns.rs

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
Dim sum = 0
Sub Count(n As Integer)
If n <= 0 Then Exit Sub
sum += 1
Count(n - 1)
End Sub
Sub Main()
Count(3)
__Check(CStr(sum), "3")
End Sub
End Module
