' vybe-test: vb/vb_modules/function_multiple_return_paths
' origin: languages/vb/tests/vb/test_vb_modules.rs

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
    Function Classify(x As Integer) As String
        If x > 0 Then Return "positive"
        If x < 0 Then Return "negative"
        Return "zero"
    End Function
    Sub Main()
        __Check(CStr(Classify(5)), "positive")
        __Check(CStr(Classify(-3)), "negative")
        __Check(CStr(Classify(0)), "zero")
    End Sub
End Module
