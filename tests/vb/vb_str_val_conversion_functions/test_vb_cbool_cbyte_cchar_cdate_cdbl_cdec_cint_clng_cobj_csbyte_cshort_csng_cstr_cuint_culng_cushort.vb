' vybe-test: vb/vb_str_val_conversion_functions/test_vb_cbool_cbyte_cchar_cdate_cdbl_cdec_cint_clng_cobj_csbyte_cshort_csng_cstr_cuint_culng_cushort
' origin: languages/vb/tests/vb/test_vb_str_val_conversion_functions.rs

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
        Dim n = CInt("123")
        Dim d = CDbl("45.67")
        Dim s = CStr(999)
        Dim b = CBool(1)
        __Check(CStr(n & "|" & d & "|" & s & "|" & b), "123|45.67|999|True")
    End Sub
End Module
