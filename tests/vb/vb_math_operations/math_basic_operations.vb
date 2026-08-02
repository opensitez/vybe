' vybe-test: vb/vb_math_operations/math_basic_operations
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
        __Check(CStr(Abs(-10)), "10")
        __Check(CStr(Max(5, 10)), "10")
        __Check(CStr(Min(5, 10)), "5")
        __Check(CStr(Pow(2, 3)), "8")
        __Check(CStr(Sqrt(16)), "4")
    End Sub
End Module
