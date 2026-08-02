' vybe-test: vb/vb_comprehensive/function_with_return_statement
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

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
    Function MaxVal(a As Integer, b As Integer) As Integer
        If a > b Then
            Return a
        End If
        Return b
    End Function

    Sub Main()
        __Check(CStr(MaxVal(10, 20)), "20")
    End Sub
End Module
