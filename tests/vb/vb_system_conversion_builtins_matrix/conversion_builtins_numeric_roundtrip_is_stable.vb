' vybe-test: vb/vb_system_conversion_builtins_matrix/conversion_builtins_numeric_roundtrip_is_stable
' origin: languages/vb/tests/vb/test_vb_system_conversion_builtins_matrix.rs

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

Imports System

Module M
    Sub Main()
        __Check(CStr(Convert.ToInt32(123)), "123")
        __Check(CStr(Convert.ToInt64(12.5)), "12")
        __Check(CStr(Convert.ToDecimal("10.25")), "10.25")
        __Check(CStr(Convert.ToString(7.5)), "7.5")
    End Sub
End Module
