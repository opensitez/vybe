' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_primitive_widening_conversions
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
        Dim b As Byte = 200
        Dim i As Integer = CType(b, Integer)
        Dim d As Double = CType(i, Double)
        Dim dec As Decimal = CType(d, Decimal)
        __Check(CStr(dec), "200")
    End Sub
End Module
