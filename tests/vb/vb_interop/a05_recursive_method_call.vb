' vybe-test: vb/vb_interop/a05_recursive_method_call
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

Public Class MathHelper
    Public Function Factorial(n As Integer) As Integer
        If n <= 1 Then
            Return 1
        End If
        Return n * Factorial(n - 1)
    End Function
End Class
Dim m As New MathHelper()
__Check(CStr(m.Factorial(5)), "120")
