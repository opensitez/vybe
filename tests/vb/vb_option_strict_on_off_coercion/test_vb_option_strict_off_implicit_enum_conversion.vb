' vybe-test: vb/vb_option_strict_on_off_coercion/test_vb_option_strict_off_implicit_enum_conversion
' origin: languages/vb/tests/vb/test_vb_option_strict_on_off_coercion.rs

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

Enum Mode
    Off = 0
    OnVal = 1
End Enum

Module Program
    Sub Main()
        Dim num As Integer = 1
        Dim m As Mode = num ' Implicit conversion from Integer to Enum
        __Check(CStr(m.ToString()), "OnVal")
    End Sub
End Module
