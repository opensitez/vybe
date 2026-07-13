use super::helpers::run_vb;

#[test]
fn select_case_typeof_is() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Class Cat
    Inherits Animal
End Class

Module M
    Sub TestType(obj As Object)
        ' VB doesn't natively support Select Case TypeOf obj Is ...
        ' We use Select Case True
        Select Case True
            Case TypeOf obj Is Dog
                Console.WriteLine("Dog")
            Case TypeOf obj Is Cat
                Console.WriteLine("Cat")
            Case Else
                Console.WriteLine("Unknown")
        End Select
    End Sub

    Sub Main()
        TestType(New Dog())
        TestType(New Cat())
        TestType(New Animal())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Dog", "Cat", "Unknown"]);
}
