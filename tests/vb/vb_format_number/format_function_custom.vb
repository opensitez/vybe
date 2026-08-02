' vybe-test: vb/vb_format_number/format_function_custom
' origin: languages/vb/tests/vb/test_vb_format_number.rs

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
        __Check(CStr(Format(1234.5, "0,0.00")), "1,234.50")
        __Check(CStr(Format(#12/25/2026#, "yyyy-MM-dd")), "2026-12-25")
    End Sub
End Module
