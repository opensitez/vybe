' vybe-test: vb/vb_option_compare_text/option_compare_binary
' origin: languages/vb/tests/vb/test_vb_option_compare_text.rs

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

Option Compare Binary

Module M
    Sub Main()
        Dim s1 As String = "HELLO"
        Dim s2 As String = "hello"
        
        ' With Option Compare Binary, case-sensitive string comparison is used
        __Check(CStr(s1 = s2), "False")
    End Sub
End Module
