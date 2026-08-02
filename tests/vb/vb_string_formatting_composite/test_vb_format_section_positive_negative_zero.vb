' vybe-test: vb/vb_string_formatting_composite/test_vb_format_section_positive_negative_zero
' origin: languages/vb/tests/vb/test_vb_string_formatting_composite.rs

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
        Dim pos As Double = 5.0
        Dim neg As Double = -5.0
        Dim zero As Double = 0.0
        Dim fmt As String = "{0:Pos:#;Neg:#;Zero}"
        __Check(CStr(String.Format(fmt, pos)), "Pos:5")
        __Check(CStr(String.Format(fmt, neg)), "Neg:5")
        __Check(CStr(String.Format(fmt, zero)), "Zero")
    End Sub
End Module
