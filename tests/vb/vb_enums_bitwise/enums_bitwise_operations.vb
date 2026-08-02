' vybe-test: vb/vb_enums_bitwise/enums_bitwise_operations
' origin: languages/vb/tests/vb/test_vb_enums_bitwise.rs

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
Enum Permissions As Byte
    None = 0
    Read = 1
    Write = 2
    Execute = 4
End Enum

Module M
    Sub Main()
        Dim p As Permissions = Permissions.Read Or Permissions.Write
        
        __Check(CStr(p.HasFlag(Permissions.Read)), "True")
        __Check(CStr(p.HasFlag(Permissions.Execute)), "False")
        
        ' Bitwise And to check flag
        Dim isWrite = (p And Permissions.Write) = Permissions.Write
        __Check(CStr(isWrite), "True")
        
        ' Removing a flag
        p = p And Not Permissions.Read
        __Check(CStr(p.HasFlag(Permissions.Read)), "False")
    End Sub
End Module
