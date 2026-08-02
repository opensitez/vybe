' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_bankers_rounding_float_operands
' origin: languages/vb/tests/vb/test_vb_mod_operator_floating_point.rs

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
        ' 2.5 rounds to 2 (even), 3.5 rounds to 4 (even)
        ' 2 Mod 4 = 2
        Dim res = 2.5 Mod 3.5
        __Check(CStr(res), "2")
    End Sub
End Module
