' vybe-test: vb/vb_operators/arithmetic_basic
' origin: languages/vb/tests/vb/test_vb_operators.rs

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
        __Check(CStr(10 + 5), "15")
        __Check(CStr(10 - 5), "5")
        __Check(CStr(10 * 5), "50")
        __Check(CStr(10 / 4), "2.5")
        __Check(CStr(10 \ 3), "3")
        __Check(CStr(10 Mod 3), "1")
    End Sub
End Module
