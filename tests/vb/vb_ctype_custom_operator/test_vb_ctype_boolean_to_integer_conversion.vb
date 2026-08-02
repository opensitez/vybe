' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_boolean_to_integer_conversion
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
        ' In VB.NET CType(True, Integer) = -1, CType(False, Integer) = 0!
        Dim tVal As Integer = CType(True, Integer)
        Dim fVal As Integer = CType(False, Integer)
        __Check(CStr(tVal & "|" & fVal), "-1|0")
    End Sub
End Module
