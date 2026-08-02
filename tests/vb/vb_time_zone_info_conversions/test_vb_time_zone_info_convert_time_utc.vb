' vybe-test: vb/vb_time_zone_info_conversions/test_vb_time_zone_info_convert_time_utc
' origin: languages/vb/tests/vb/test_vb_time_zone_info_conversions.rs

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

Module Program
    Sub Main()
        Dim dtUtc As New DateTime(2025, 6, 1, 12, 0, 0, DateTimeKind.Utc)
        Dim localDt = TimeZoneInfo.ConvertTime(dtUtc, TimeZoneInfo.Utc)
        __Check(CStr(localDt.Hour), "12")
    End Sub
End Module
