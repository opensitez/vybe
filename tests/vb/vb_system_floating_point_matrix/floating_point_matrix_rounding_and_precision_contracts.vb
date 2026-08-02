' vybe-test: vb/vb_system_floating_point_matrix/floating_point_matrix_rounding_and_precision_contracts
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
        Dim values() As Double = {3.5, -3.5, 2.1, -2.1, 0.0, 1.499, 2.5}

        __Check(CStr(Math.Round(3.5, 0)), "4")
        __Check(CStr(Math.Round(2.5, 0)), "2")
        __Check(CStr(Math.Ceiling(-3.5)), "-3")
        __Check(CStr(Math.Floor(2.5)), "2")
        __Check(CStr(Math.Truncate(values(0))), "3")
        __Check(CStr(Math.Truncate(values(1))), "-3")
        __Check(CStr(Math.Sign(values(2))), "1")
        __Check(CStr(Math.Sign(values(5))), "1")
    End Sub
End Module
