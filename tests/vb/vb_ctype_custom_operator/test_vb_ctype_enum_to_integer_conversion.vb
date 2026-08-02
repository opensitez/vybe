' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_enum_to_integer_conversion
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Enum Level
    Low = 10
    High = 90
End Enum

Module Program
    Sub Main()
        Dim l = Level.High
        Dim val As Integer = CType(l, Integer)
        Dim restored As Level = CType(val, Level)
        __Check(CStr(val & "|" & restored.ToString()), "90|High")
    End Sub
End Module
