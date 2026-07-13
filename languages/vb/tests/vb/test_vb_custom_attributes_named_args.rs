use super::helpers::run_vb;

#[test]
fn custom_attributes_named_args() {
    let out = run_vb(
        r#"
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
            Console.WriteLine(r.RoleName)
            Console.WriteLine(r.AccessLevel)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Admin", "10"]);
}
