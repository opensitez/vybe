' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_uinteger_operands
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
        Dim u1 As UInteger = 4000000000UI
        Dim u2 As UInteger = 3UI
        Dim res = u1 Mod u2
        __Check(CStr(res), "1")
    End Sub
End Module
