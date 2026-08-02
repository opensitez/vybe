' vybe-test: vb/vb_variables_and_consts/var_static_in_function
' origin: languages/vb/tests/vb/test_vb_variables_and_consts.rs

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
Function F() As Integer
Static x As Integer = 10
x += 10
Return x
End Function
Sub Main()
__Check(CStr(F()), "20")
End Sub
End Module
