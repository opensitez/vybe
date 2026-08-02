' vybe-test: vb/vb_hex_octal_literals/hex_octal_literals_advanced
' origin: languages/vb/tests/vb/test_vb_hex_octal_literals.rs

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
        ' Hex literal
        Dim h As Integer = &HFF
        
        ' Octal literal
        Dim o As Integer = &O77
        
        ' Binary literal (VB 15 / VB.NET 2017+)
        Dim b As Integer = &B1010
        
        __Check(CStr(h), "255")
        __Check(CStr(o), "63")
        __Check(CStr(b), "10")
    End Sub
End Module
