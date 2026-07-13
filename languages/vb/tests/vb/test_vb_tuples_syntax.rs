use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Tuples Syntax (VB 15+)
// ═══════════════════════════════════════════════════════════

#[test]
fn tuple_literal_unnamed() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim t = (1, "Apple")
        Console.WriteLine(t.Item1)
        Console.WriteLine(t.Item2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "Apple"]);
}

#[test]
fn tuple_literal_named() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim person = (Name:="Alice", Age:=30)
        Console.WriteLine(person.Name)
        Console.WriteLine(person.Age)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn tuple_function_return() {
    let out = run_vb(
        r#"
Module M
    Function GetCoordinates() As (X As Integer, Y As Integer)
        Return (10, 20)
    End Function

    Sub Main()
        Dim coords = GetCoordinates()
        Console.WriteLine(coords.X)
        Console.WriteLine(coords.Y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}
