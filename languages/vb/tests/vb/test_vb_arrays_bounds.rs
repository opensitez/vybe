use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Arrays (Bounds, N+1 rule)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_declaration_upper_bound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' In VB, declaring an array with (5) means the upper bound is 5,
        ' so there are 6 elements (0 through 5).
        Dim arr(5) As Integer
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_declaration_explicit_lower_bound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' VB supports explicit bounds 0 To N in declarations
        Dim arr(0 To 4) As Integer
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_initialization_with_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' If bounds are provided and initialized, they must match the initializers length
        ' Dim arr(2) As Integer = {1, 2, 3} ' 3 elements (0,1,2)
        Dim arr(2) As Integer = {10, 20, 30}
        Console.WriteLine(arr(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn array_multidimensional_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim grid(2, 3) As Integer
        ' Length is total elements: (2+1) * (3+1) = 3 * 4 = 12
        Console.WriteLine(grid.Length)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12"]);
}
