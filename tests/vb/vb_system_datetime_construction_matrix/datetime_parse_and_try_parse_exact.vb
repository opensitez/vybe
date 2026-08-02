' vybe-test: vb/vb_system_datetime_construction_matrix/datetime_parse_and_try_parse_exact
' origin: languages/vb/tests/vb/test_vb_system_datetime_construction_matrix.rs

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
        Dim parsed As DateTime = DateTime.Parse("2026-07-21")
        Dim ok As Boolean
        Dim exact As DateTime

        ok = DateTime.TryParseExact("21/07/2026", "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, exact)

        __Check(CStr(parsed.Year), "2026")
        __Check(CStr(ok), "True")
        __Check(CStr(exact.Month), "7")
    End Sub
End Module
