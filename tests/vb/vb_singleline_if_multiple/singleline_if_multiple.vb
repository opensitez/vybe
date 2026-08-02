' vybe-test: vb/vb_singleline_if_multiple/singleline_if_multiple
' origin: languages/vb/tests/vb/test_vb_singleline_if_multiple.rs

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
        Dim x = 10
        Dim y = 0
        Dim z = 0
        
        ' Single line If with multiple statements separated by colon
        If x = 10 Then y = 1 : z = 2
        
        __Check(CStr(y & "-" & z), "1-2")
    End Sub
End Module
