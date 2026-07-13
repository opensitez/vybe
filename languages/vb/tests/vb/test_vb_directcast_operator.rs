use super::helpers::run_vb;

#[test]
fn directcast_operator() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
    Public Sub Bark()
        Console.WriteLine("Woof")
    End Sub
End Class

Module M
    Sub Main()
        Dim a As Animal = New Dog()
        
        ' DirectCast requires the run-time type of an object variable to be the same as the specified type.
        ' It is faster than CType but throws an exception if the cast fails.
        Dim d As Dog = DirectCast(a, Dog)
        d.Bark()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof"]);
}
