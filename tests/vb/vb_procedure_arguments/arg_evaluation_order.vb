' vybe-test: vb/vb_procedure_arguments/arg_evaluation_order
' origin: languages/vb/tests/vb/test_vb_procedure_arguments.rs

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
Dim v = 1
Function F() As Integer
v += 1
Return v
End Function
Sub Print(a As Integer, b As Integer)
__Check(CStr(a & b), "23")
End Sub
Sub Main()
Print(F(), F())
End Sub
End Module
