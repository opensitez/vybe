' vybe-test: vb/vb_string_interpolation_fmt_align/string_interpolation_fmt_align
' origin: languages/vb/tests/vb/test_vb_string_interpolation_fmt_align.rs

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
        Dim val As Double = 42.5
        
        ' String interpolation with alignment and format specifier
        Dim s = $"[{val,10:F2}]"
        __Check(CStr(s), "[     42.50]")
    End Sub
End Module
