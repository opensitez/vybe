' vybe-test: vb/vb_system_conversion_builtins_matrix/conversion_builtins_date_and_time_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_conversion_builtins_matrix.rs

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
        Dim d As DateTime = Convert.ToDateTime("2026-07-21T12:00:00")
        __Check(CStr(d.Year), "2026")
        __Check(CStr(d.Month), "7")
        __Check(CStr(d.Day), "21")
        __Check(CStr(Convert.ToString(d.Date)), "07/21/2026")
    End Sub
End Module
