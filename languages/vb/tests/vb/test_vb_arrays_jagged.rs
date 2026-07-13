use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Arrays (Jagged Arrays)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_jagged_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim jagged(2)() As Integer
        
        jagged(0) = New Integer() {1, 2}
        jagged(1) = New Integer() {3, 4, 5}
        jagged(2) = New Integer() {6}
        
        Console.WriteLine(jagged(1)(2))
        Console.WriteLine(jagged.Length)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "3"]);
}

#[test]
fn array_jagged_initialization() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim jagged As Integer()() = {
            New Integer() {10, 20},
            New Integer() {30, 40, 50}
        }
        
        Console.WriteLine(jagged(0)(1))
        Console.WriteLine(jagged(1)(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20", "30"]);
}

#[test]
fn array_jagged_multidimensional_mix() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Array of 2D arrays
        Dim mix(1)(,) As Integer
        
        mix(0) = New Integer(1, 1) {{1, 2}, {3, 4}}
        mix(1) = New Integer(0, 0) {{9}}
        
        Console.WriteLine(mix(0)(1, 0))
        Console.WriteLine(mix(1)(0, 0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "9"]);
}
