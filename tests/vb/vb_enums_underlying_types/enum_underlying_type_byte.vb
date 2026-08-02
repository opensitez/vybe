' vybe-test: vb/vb_enums_underlying_types/enum_underlying_type_byte
' origin: languages/vb/tests/vb/test_vb_enums_underlying_types.rs

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

Enum SmallNumber As Byte
    Zero = 0
    One = 1
    Two = 2
End Enum

Module M
    Sub Main()
        Dim s As SmallNumber = SmallNumber.Two
        ' Prints the numeric value when cast to Byte
        __Check(CStr(CByte(s)), "2")
    End Sub
End Module
