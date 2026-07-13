use super::helpers::run_vb;

#[test]
fn late_binding_ctype_dynamic() {
    let out = run_vb(
        r#"
Option Strict Off

Class Person
    Public Property Name As String
End Class

Module M
    Sub Main()
        Dim o As Object = New Person() With {.Name = "Bob"}
        
        ' Late binding assignment
        o.Name = "Alice"
        
        ' Late binding retrieval
        Dim n As String = CType(o.Name, String)
        Console.WriteLine(n)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice"]);
}
