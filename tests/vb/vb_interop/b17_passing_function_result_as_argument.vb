' vybe-test: vb/vb_interop/b17_passing_function_result_as_argument
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

Function Double(n As Integer) As Integer
    Return n * 2
End Function
Function AddOne(n As Integer) As Integer
    Return n + 1
End Function
__Check(CStr(AddOne(Double(5))), "11")
