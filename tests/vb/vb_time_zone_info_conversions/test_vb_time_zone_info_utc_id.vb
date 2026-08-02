' vybe-test: vb/vb_time_zone_info_conversions/test_vb_time_zone_info_utc_id
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
        Dim utcZone = TimeZoneInfo.Utc
        __Check(CStr(utcZone.Id), "UTC")
        __Check(CStr(utcZone.BaseUtcOffset.TotalHours), "0")
    End Sub
End Module
