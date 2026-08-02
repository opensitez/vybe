' vybe-test: vb/vb_operator_precedence/operator_precedence_math
' origin: languages/vb/tests/vb/test_vb_operator_precedence.rs

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
        ' ^ has higher precedence than *, /
        __Check(CStr(2 + 3 * 4 ^ 2), "50") ' 2 + 3 * 16 = 2 + 48 = 50
        
        ' \ (integer division) has lower precedence than *, / but higher than +, -
        __Check(CStr(10 + 20 \ 3), "16") ' 10 + 6 = 16
        
        ' Mod has lower precedence than \
        __Check(CStr(10 Mod 3 + 1), "2") ' 1 + 1 = 2
    End Sub
End Module
