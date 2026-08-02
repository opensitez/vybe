' vybe-test: vb/vb_type_casting/type_casting_primitives
' origin: languages/vb/tests/vb/test_vb_type_casting.rs

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
        __Check(CStr(CInt(2.5)), "2") ' 2 (Banker's rounding)
        __Check(CStr(CDbl("3.14")), "3.14")
        __Check(CStr(CStr(42)), "42")
        __Check(CStr(CBool(1)), "True")
    End Sub
End Module
