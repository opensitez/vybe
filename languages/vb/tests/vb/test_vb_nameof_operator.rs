use super::helpers::run_vb;

#[test]
fn nameof_operator() {
    let out = run_vb(
        r#"
Class Person
    Public Property Name As String
End Class

Module M
    Sub Main()
        ' NameOf returns the name of a variable, type, or member
        Console.WriteLine(NameOf(Person))
        Console.WriteLine(NameOf(Person.Name))
        
        Dim i = 10
        Console.WriteLine(NameOf(i))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Person", "Name", "i"]);
}
