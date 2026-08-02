' vybe-test: vb/vb_string_composite_formatting_alignment/test_vb_string_format_conditional_three_section_specifier
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

Imports System.Globalization

Module Program
    Sub Main()
        Dim fmt = "{0:+#,##0;-#,##0;ZERO}"
        __Check(CStr(String.Format(CultureInfo.InvariantCulture, fmt, 50) & "|" & String.Format(CultureInfo.InvariantCulture, fmt, -50) & "|" & String.Format(CultureInfo.InvariantCulture, fmt, 0)), "+50|-50|ZERO")
    End Sub
End Module
