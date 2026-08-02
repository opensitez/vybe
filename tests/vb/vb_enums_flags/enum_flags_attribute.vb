' vybe-test: vb/vb_enums_flags/enum_flags_attribute
' origin: languages/vb/tests/vb/test_vb_enums_flags.rs

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

<Flags>
Enum FileAccess
    None = 0
    Read = 1
    Write = 2
    ReadWrite = Read Or Write
End Enum

Module M
    Sub Main()
        Dim access As FileAccess = FileAccess.ReadWrite
        __Check(CStr(access.ToString()), "ReadWrite")
        
        Dim singleAccess As FileAccess = FileAccess.Read
        __Check(CStr(singleAccess.ToString()), "Read")
    End Sub
End Module
