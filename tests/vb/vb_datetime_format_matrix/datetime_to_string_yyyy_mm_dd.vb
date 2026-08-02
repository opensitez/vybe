' vybe-test: vb/vb_datetime_format_matrix/datetime_to_string_yyyy_mm_dd
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

Module M
    Sub Main()
        Dim d As Date = New Date(2024, 6, 15)
        __Check(CStr(d.ToString("yyyy-MM-dd")), "2024-06-15")
    End Sub
End Module
