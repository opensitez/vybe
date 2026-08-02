' vybe-test: vb/vb_enum_underlying_types/enum_underlying_types
' origin: languages/vb/tests/vb/test_vb_enum_underlying_types.rs

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

' Enum with explicit underlying type Byte
Enum Status As Byte
    Active = 1
    Inactive = 2
    Pending = 3
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Active
        
        ' Check underlying type
        __Check(CStr(s.GetTypeCode().ToString()), "Byte")
        __Check(CStr(s), "Active")
    End Sub
End Module
