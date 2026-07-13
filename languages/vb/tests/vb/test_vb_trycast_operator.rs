use super::helpers::run_vb;

#[test]
fn trycast_operator() {
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
        Dim a2 As Animal = New Animal()
        
        ' TryCast attempts to cast to a reference type, returning Nothing if it fails
        Dim d1 As Dog = TryCast(a, Dog)
        If d1 IsNot Nothing Then
            d1.Bark()
        End If
        
        Dim d2 As Dog = TryCast(a2, Dog)
        Console.WriteLine(d2 Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof", "True"]);
}
