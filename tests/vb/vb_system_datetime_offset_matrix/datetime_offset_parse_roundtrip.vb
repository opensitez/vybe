' vybe-test: vb/vb_system_datetime_offset_matrix/datetime_offset_parse_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_datetime_offset_matrix.rs

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

Module M
    Sub Main()
        Dim dto As DateTimeOffset = DateTimeOffset.Parse("2024-07-21T03:04:05+02:00")
        Dim text As String = dto.ToString("o")
        Dim again As DateTimeOffset = DateTimeOffset.Parse(text)
        __Check(CStr(dto = again), "True")
    End Sub
End Module
