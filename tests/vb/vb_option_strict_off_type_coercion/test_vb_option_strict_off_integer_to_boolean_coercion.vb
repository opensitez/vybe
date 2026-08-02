' vybe-test: vb/vb_option_strict_off_type_coercion/test_vb_option_strict_off_integer_to_boolean_coercion
' origin: languages/vb/tests/vb/test_vb_option_strict_off_type_coercion.rs

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

Module Program
    Sub Main()
        Dim zero As Object = 0
        Dim nonZero As Object = 5
        Dim b1 As Boolean = zero
        Dim b2 As Boolean = nonZero
        __Check(CStr(b1 & "|" & b2), "False|True")
    End Sub
End Module
