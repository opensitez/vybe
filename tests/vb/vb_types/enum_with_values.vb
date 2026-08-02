' vybe-test: vb/vb_types/enum_with_values
' origin: languages/vb/tests/vb/test_vb_types.rs

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

Enum Status
    Active = 1
    Inactive = 0
    Pending = 2
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Pending
        __Check(CStr(s), "2")
    End Sub
End Module
