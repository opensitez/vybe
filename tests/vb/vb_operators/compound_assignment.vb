' vybe-test: vb/vb_operators/compound_assignment
' origin: languages/vb/tests/vb/test_vb_operators.rs

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
        Dim x As Integer = 10
        x += 5
        __Check(CStr(x), "15")
        x -= 3
        __Check(CStr(x), "12")
        x *= 2
        __Check(CStr(x), "24")
    End Sub
End Module
