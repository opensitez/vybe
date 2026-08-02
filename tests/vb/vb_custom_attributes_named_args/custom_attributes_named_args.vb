' vybe-test: vb/vb_custom_attributes_named_args/custom_attributes_named_args
' origin: languages/vb/tests/vb/test_vb_custom_attributes_named_args.rs

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

<AttributeUsage(AttributeTargets.Class)>
Class RoleAttribute
    Inherits Attribute
    
    Public Property RoleName As String
    Public Property AccessLevel As Integer
End Class

<Role(RoleName:="Admin", AccessLevel:=10)>
Class SecureData
End Class

Module M
    Sub Main()
        Dim attrs = GetType(SecureData).GetCustomAttributes(GetType(RoleAttribute), False)
        If attrs.Length > 0 Then
            Dim r = CType(attrs(0), RoleAttribute)
            __Check(CStr(r.RoleName), "Admin")
            __Check(CStr(r.AccessLevel), "10")
        End If
    End Sub
End Module
