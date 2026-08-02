' vybe-test: vb/vb_date_time_compare_is_leap_year/test_vb_date_time_binary_representation_to_from
' origin: languages/vb/tests/vb/test_vb_date_time_compare_is_leap_year.rs

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
        Dim d1 As New DateTime(2025, 10, 31, 15, 45, 0, DateTimeKind.Utc)
        Dim bin = d1.ToBinary()
        Dim d2 = DateTime.FromBinary(bin)
        __Check(CStr((d1 = d2) & "|" & d2.Kind.ToString()), "True|Utc")
    End Sub
End Module
