' vybe-test: vb/vb_datetime_format_matrix/datetime_parse_exact_with_invariant_culture
' origin: languages/vb/tests/vb/test_vb_datetime_format_matrix.rs

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

Module M
    Sub Main()
        Dim d As Date = Date.ParseExact("2024-03-21", "yyyy-MM-dd", CultureInfo.InvariantCulture)
        __Check(CStr(d.Year), "2024")
        __Check(CStr(d.Month), "3")
        __Check(CStr(d.Day), "21")
    End Sub
End Module
