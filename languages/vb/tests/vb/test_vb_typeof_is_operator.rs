use super::helpers::run_vb;

#[test]
fn typeof_is_operator() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim d As New Dog()
        Dim a As Animal = d
        
        ' TypeOf ... Is / IsNot
        Console.WriteLine(TypeOf d Is Dog)
        Console.WriteLine(TypeOf d Is Animal)
        Console.WriteLine(TypeOf d IsNot String)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}
