' vybe-test: vb/vb_interop/b18_nested_function_calls_as_arguments
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Function Add(a As Integer, b As Integer) As Integer
    Return a + b
End Function
Function Mul(a As Integer, b As Integer) As Integer
    Return a * b
End Function
__Check(CStr(Add(Mul(2, 3), Mul(4, 5))), "26")
