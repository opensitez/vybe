' vybe-test: vb/vb_null_conditional/null_conditional_operator
' origin: languages/vb/tests/vb/test_vb_null_conditional.rs

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

Class Data
    Public Property Value As String
End Class

Module M
    Sub Main()
        Dim d As Data = Nothing
        
        ' Using ? before dot checks if d is nothing
        Dim len As Integer? = d?.Value?.Length
        
        __Check(CStr(len.HasValue), "False")
        
        d = New Data() With { .Value = "Test" }
        len = d?.Value?.Length
        __Check(CStr(len.HasValue), "True")
        __Check(CStr(len.Value), "4")
    End Sub
End Module
