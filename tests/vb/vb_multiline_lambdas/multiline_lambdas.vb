' vybe-test: vb/vb_multiline_lambdas/multiline_lambdas
' origin: languages/vb/tests/vb/test_vb_multiline_lambdas.rs

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
        ' Multiline Sub lambda
        Dim printSum = Sub(x As Integer, y As Integer)
                           Dim result = x + y
                           __Check(CStr(result), "12")
                       End Sub
                       
        ' Multiline Function lambda
        Dim getGreeting = Function(name As String) As String
                              Dim prefix = "Hello, "
                              Return prefix & name
                          End Function
                          
        printSum(5, 7)
        __Check(CStr(getGreeting("Alice")), "Hello, Alice")
    End Sub
End Module
