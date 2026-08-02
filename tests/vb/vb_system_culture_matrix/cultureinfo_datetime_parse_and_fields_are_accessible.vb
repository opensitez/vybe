' vybe-test: vb/vb_system_culture_matrix/cultureinfo_datetime_parse_and_fields_are_accessible
' origin: languages/vb/tests/vb/test_vb_system_culture_matrix.rs

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

Imports System
Imports System.Globalization

    Module M
    Sub Main()
        Dim dt As DateTime = DateTime.Parse("07/21/2026", CultureInfo.GetCultureInfo("en-US"))
        __Check(CStr(dt.Year), "2026")
        __Check(CStr(dt.Month), "7")
        __Check(CStr(dt.Day), "21")
    End Sub
End Module
