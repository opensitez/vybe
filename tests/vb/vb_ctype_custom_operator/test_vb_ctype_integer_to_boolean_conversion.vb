' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_integer_to_boolean_conversion
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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
    Sub Main()
        Dim b1 As Boolean = CType(-1, Boolean)
        Dim b2 As Boolean = CType(100, Boolean)
        Dim b3 As Boolean = CType(0, Boolean)
        __Check(CStr(b1 & "|" & b2 & "|" & b3), "True|True|False")
    End Sub
End Module
