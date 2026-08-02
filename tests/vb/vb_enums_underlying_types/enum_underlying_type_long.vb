' vybe-test: vb/vb_enums_underlying_types/enum_underlying_type_long
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

Enum BigEnum As Long
    Max = 9223372036854775807
    Min = -9223372036854775808
End Enum

Module M
    Sub Main()
        __Check(CStr(CLng(BigEnum.Max)), "9223372036854775807")
    End Sub
End Module
