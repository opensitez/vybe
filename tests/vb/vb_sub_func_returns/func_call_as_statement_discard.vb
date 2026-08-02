' vybe-test: vb/vb_sub_func_returns/func_call_as_statement_discard
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
Function GetV() As Integer
Return 1
End Function
Sub Main()
GetV()
__Check(CStr("OK"), "OK")
End Sub
End Module
