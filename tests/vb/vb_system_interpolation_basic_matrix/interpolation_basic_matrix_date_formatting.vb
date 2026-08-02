' vybe-test: vb/vb_system_interpolation_basic_matrix/interpolation_basic_matrix_date_formatting
' origin: languages/vb/tests/vb/test_vb_system_interpolation_basic_matrix.rs

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
        Dim dt As Date = #2026-07-21 08:15:00#
        Dim s As String = $"{dt:yyyy-MM-dd}"
        Dim t As String = $"{dt:HH:mm}"

        __Check(CStr(s), "2026-07-21")
        __Check(CStr(t), "08:15")
    End Sub
End Module
