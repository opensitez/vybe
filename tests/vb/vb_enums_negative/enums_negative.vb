' vybe-test: vb/vb_enums_negative/enums_negative
' origin: languages/vb/tests/vb/test_vb_enums_negative.rs

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

Enum Status As Short
    Error = -1
    Pending = 0
    Active = 1
    Completed = 2
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Error
        __Check(CStr(s), "Error")
        
        Dim val As Short = CShort(s)
        __Check(CStr(val), "-1")
    End Sub
End Module
