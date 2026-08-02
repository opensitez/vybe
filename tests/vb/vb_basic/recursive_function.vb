' vybe-test: vb/vb_basic/recursive_function
' origin: languages/vb/tests/vb/vb_basic_test.rs

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
    Function Factorial(n As Integer) As Integer
        If n <= 1 Then
            Return 1
        End If
        Return n * Factorial(n - 1)
    End Function

    Sub Main()
        __Check(CStr(Factorial(5)), "120")
    End Sub
End Module
