' vybe-test: vb/vb_interaction_iif/interaction_if_operator_null_coalescing
' origin: languages/vb/tests/vb/test_vb_interaction_iif.rs

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
        Dim s1 As String = Nothing
        Dim s2 As String = "Fallback"
        
        ' If operator with two arguments acts like null-coalescing
        Dim result As String = If(s1, s2)
        __Check(CStr(result), "Fallback")
    End Sub
End Module
