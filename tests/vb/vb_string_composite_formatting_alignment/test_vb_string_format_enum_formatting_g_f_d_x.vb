' vybe-test: vb/vb_string_composite_formatting_alignment/test_vb_string_format_enum_formatting_g_f_d_x
' origin: languages/vb/tests/vb/test_vb_string_composite_formatting_alignment.rs

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

Enum Level
    Low = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim l = Level.High
        __Check(CStr(String.Format("{0:G}|{0:D}|{0:X}", l)), "High|2|00000002")
    End Sub
End Module
