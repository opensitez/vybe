' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_datetime_formatting
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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
        Dim dates As DateTime() = {New DateTime(2025, 1, 1), New DateTime(2025, 12, 31)}
        Dim formatted As String() = Array.ConvertAll(dates, Function(d) d.ToString("yyyy-MM-dd"))
        __Check(CStr(String.Join(";", formatted)), "2025-01-01;2025-12-31")
    End Sub
End Module
