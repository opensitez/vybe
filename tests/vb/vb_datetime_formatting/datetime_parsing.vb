' vybe-test: vb/vb_datetime_formatting/datetime_parsing
' origin: languages/vb/tests/vb/test_vb_datetime_formatting.rs

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
        Dim d1 As Date = Date.Parse("2024-01-01", CultureInfo.InvariantCulture)
        __Check(CStr(d1.Year), "2024")
        
        Dim d2 As Date
        If Date.TryParseExact("20240101", "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, d2) Then
            __Check(CStr(d2.Month), "1")
        End If
    End Sub
End Module
