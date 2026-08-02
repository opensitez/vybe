' vybe-test: vb/vb_numeric_type_conversions_adv/test_vb_conv_enum_to_underlying_and_back
' origin: languages/vb/tests/vb/test_vb_numeric_type_conversions_adv.rs

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

Enum ColorCode As Byte
    Red = 1
    Green = 2
    Blue = 3
End Enum

Module Program
    Sub Main()
        Dim c As ColorCode = ColorCode.Green
        Dim b As Byte = CByte(c)
        Dim c2 As ColorCode = CType(3, ColorCode)
        __Check(CStr(b), "2")
        __Check(CStr(c2.ToString()), "Blue")
    End Sub
End Module
