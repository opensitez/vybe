' vybe-test: vb/vb_like_operator/like_operator_wildcards
' origin: languages/vb/tests/vb/test_vb_like_operator.rs

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
        Dim s1 As String = "Bat"
        Dim s2 As String = "Cat"
        Dim s3 As String = "Hat"
        
        ' ? matches any single character
        __Check(CStr(s1 Like "?at"), "True")
        
        ' * matches zero or more characters
        __Check(CStr(s2 Like "C*"), "True")
        
        ' # matches any single digit
        __Check(CStr("123" Like "1#3"), "True")
        __Check(CStr("1a3" Like "1#3"), "False")
        
        ' Character lists
        __Check(CStr(s1 Like "[BCH]at"), "True")
        __Check(CStr("Mat" Like "[BCH]at"), "False")
        
        ' Character list negation
        __Check(CStr("Mat" Like "[!BCH]at"), "True")
    End Sub
End Module
