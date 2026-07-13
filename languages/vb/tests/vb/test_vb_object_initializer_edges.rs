use super::helpers::run_vb;

#[test]
fn object_initializer_with() {
    let out = run_vb(
        r#"
Class Person
    Public Property Name As String
    Public Property Age As Integer
End Class

Module M
    Sub Main()
        ' With block object initialization
        Dim p1 As New Person With {
            .Name = "Alice",
            .Age = 30
        }
        
        ' Anonymous type with inferred names from properties
        Dim p2 = New With { p1.Name, p1.Age }
        
        Console.WriteLine(p1.Name)
        Console.WriteLine(p2.Age)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}
