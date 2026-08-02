' vybe-test: vb/vb_implicit_conversions/implicit_double_to_integer_bankers_rounding_up
' origin: languages/vb/tests/vb/test_vb_implicit_conversions.rs

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

Option Strict Off
Module M
Sub Main()
Dim d As Double = 3.5
Dim i As Integer = d
__Check(CStr(i), "4")
End Sub
End Module
