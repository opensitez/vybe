use super::helpers::run_vb;

#[test]
fn with_statement_nested() {
    let out = run_vb(
        r#"
Class Address
    Public Property City As String
    Public Property Zip As String
End Class

Class Person
    Public Property Name As String
    Public Property Home As New Address()
End Class

Module M
    Sub Main()
        Dim p As New Person()
        
        ' Nested With statements
        With p
            .Name = "Alice"
            With .Home
                .City = "Wonderland"
                .Zip = "12345"
            End With
        End With
        
        Console.WriteLine(p.Name)
        Console.WriteLine(p.Home.City)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "Wonderland"]);
}
