' vybe-test: vb/vb_math_operations/math_rounding_methods
' origin: languages/vb/tests/vb/test_vb_math_operations.rs

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

Imports System.Math

Module M
    Sub Main()
        __Check(CStr(Round(2.5)), "2") ' Banker's rounding
        __Check(CStr(Round(3.5)), "4")
        __Check(CStr(Ceiling(2.1)), "3")
        __Check(CStr(Floor(2.9)), "2")
    End Sub
End Module
