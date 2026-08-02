' vybe-test: vb/vb_system_floating_point_matrix/floating_point_matrix_exponential_logs_and_trig
' origin: languages/vb/tests/vb/test_vb_system_floating_point_matrix.rs

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
        __Check(CStr(Math.Round(Math.Log(Math.Exp(2)), 6)), "2")
        __Check(CStr(Math.Round(Math.Sqrt(81), 6)), "9")
        __Check(CStr(Math.Round(Math.Sin(0), 6)), "0")
        __Check(CStr(Math.Round(Math.Cos(Math.PI), 6)), "-1")
        __Check(CStr(Math.Round(Math.Tan(Math.PI / 4), 6)), "1")
    End Sub
End Module
