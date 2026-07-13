use super::helpers::run_vb;

#[test]
fn info_typename_vartype() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "test"
        Dim i As Integer = 10
        
        ' TypeName returns a friendly string of the type
        Console.WriteLine(TypeName(s))
        Console.WriteLine(TypeName(i))
        
        ' VarType returns an enum value from Microsoft.VisualBasic.VariantType
        Console.WriteLine(CInt(VarType(s))) ' VariantType.String = 8
        Console.WriteLine(CInt(VarType(i))) ' VariantType.Integer = 3
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["String", "Integer", "8", "3"]);
}
