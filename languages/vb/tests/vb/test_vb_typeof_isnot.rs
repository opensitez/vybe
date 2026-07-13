use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TypeOf ... IsNot Operator
// ═══════════════════════════════════════════════════════════

#[test]
fn typeof_isnot() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim obj As Object = New Animal()
        
        ' VB.NET allows TypeOf ... IsNot
        If TypeOf obj IsNot Dog Then
            Console.WriteLine("Not a dog")
        End If
        
        Dim dogObj As Object = New Dog()
        If TypeOf dogObj IsNot Dog Then
            Console.WriteLine("This shouldn't print")
        Else
            Console.WriteLine("Is a dog")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Not a dog", "Is a dog"]);
}
