use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Object Initializers Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn object_initializers_nested() {
    let out = run_vb(
        r#"
Class Address
    Public Property City As String
    Public Property Zip As String
End Class

Class Person
    Public Property Name As String
    Public Property HomeAddress As Address
End Class

Module M
    Sub Main()
        ' Nested object initializers
        Dim p As New Person() With {
            .Name = "Alice",
            .HomeAddress = New Address() With {
                .City = "New York",
                .Zip = "10001"
            }
        }
        
        Console.WriteLine(p.Name)
        Console.WriteLine(p.HomeAddress.City)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "New York"]);
}
