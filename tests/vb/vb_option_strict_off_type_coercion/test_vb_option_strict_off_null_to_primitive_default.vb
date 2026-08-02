' vybe-test: vb/vb_option_strict_off_type_coercion/test_vb_option_strict_off_null_to_primitive_default
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
        Dim obj As Object = Nothing
        Dim n As Integer = obj
        Dim b As Boolean = obj
        Dim s As String = obj
        __Check(CStr(n & "|" & b & "|" & (s Is Nothing)), "0|False|True")
    End Sub
End Module
