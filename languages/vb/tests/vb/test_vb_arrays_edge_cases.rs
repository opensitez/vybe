use super::helpers::run_vb;

#[test]
fn arrays_jagged_initialization() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim jagged()() As Integer = {
            New Integer() {1, 2},
            New Integer() {3, 4, 5}
        }
        
        Console.WriteLine(jagged(0).Length)
        Console.WriteLine(jagged(1).Length)
        Console.WriteLine(jagged(1)(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "3", "5"]);
}

#[test]
fn arrays_covariance_ref_types() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim dogs(2) As Dog
        Dim animals() As Animal = dogs
        
        Console.WriteLine(animals.Length)
        
        Try
            ' This fails at runtime (ArrayTypeMismatchException)
            animals(0) = New Animal()
        Catch ex As System.ArrayTypeMismatchException
            Console.WriteLine("Mismatch")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "Mismatch"]);
}
