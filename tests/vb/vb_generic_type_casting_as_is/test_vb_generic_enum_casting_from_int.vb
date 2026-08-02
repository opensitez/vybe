' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_enum_casting_from_int
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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
    Low = 1
    High = 2
End Enum

Module Program
    Private Function EnumCast(Of TEnum As Structure)(val As Integer) As TEnum
        Return CType(CObj(val), TEnum)
    End Function

    Sub Main()
        Dim l As Level = EnumCast(Of Level)(2)
        __Check(CStr(l.ToString()), "High")
    End Sub
End Module
