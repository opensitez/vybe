' vybe-test: vb/vb_date_time_offset_utc_conversion/test_vb_date_time_offset_hash_code_consistency
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
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.Zero)
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.Zero)
        __Check(CStr(dto1.GetHashCode() = dto2.GetHashCode()), "True")
    End Sub
End Module
