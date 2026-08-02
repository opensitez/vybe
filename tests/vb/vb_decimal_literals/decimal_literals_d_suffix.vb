' vybe-test: vb/vb_decimal_literals/decimal_literals_d_suffix
' origin: languages/vb/tests/vb/test_vb_decimal_literals.rs

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
        ' The D suffix specifies a Decimal literal
        Dim dec As Decimal = 12.345D
        Dim bigDec = 9999999999999999999.99D
        
        __Check(CStr(dec.GetType().Name), "Decimal")
        __Check(CStr(bigDec.GetType().Name), "Decimal")
    End Sub
End Module
