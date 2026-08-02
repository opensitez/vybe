' vybe-test: vb/vb_date_time_offset_utc_conversion/test_vb_date_time_offset_min_max_value
' origin: languages/vb/tests/vb/test_vb_date_time_offset_utc_conversion.rs

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
        __Check(CStr(DateTimeOffset.MinValue.Year & "|" & DateTimeOffset.MaxValue.Year), "1|9999")
    End Sub
End Module
