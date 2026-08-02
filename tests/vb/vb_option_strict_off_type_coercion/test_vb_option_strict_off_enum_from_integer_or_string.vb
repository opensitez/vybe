' vybe-test: vb/vb_option_strict_off_type_coercion/test_vb_option_strict_off_enum_from_integer_or_string
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

Enum Status
    Inactive = 0
    Active = 1
End Enum

Module Program
    Sub Main()
        Dim obj1 As Object = 1
        Dim s1 As Status = obj1
        __Check(CStr(s1.ToString()), "Active")
    End Sub
End Module
