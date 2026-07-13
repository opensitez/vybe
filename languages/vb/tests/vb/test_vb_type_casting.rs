use super::helpers::run_vb;

#[test]
fn type_casting_ctype_directcast_trycast() {
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
        
        ' DirectCast
        Dim d2 As Dog = DirectCast(a, Dog)
        Console.WriteLine(d2 IsNot Nothing)
        
        ' TryCast
        Dim d3 As Dog = TryCast(a, Dog)
        Console.WriteLine(d3 IsNot Nothing)
        
        ' CType
        Dim d4 As Dog = CType(a, Dog)
        Console.WriteLine(d4 IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn type_casting_primitives() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(CInt(2.5)) ' 2 (Banker's rounding)
        Console.WriteLine(CDbl("3.14"))
        Console.WriteLine(CStr(42))
        Console.WriteLine(CBool(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "3.14", "42", "True"]);
}
