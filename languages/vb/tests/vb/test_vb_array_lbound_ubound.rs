use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Methods (LBound and UBound)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_lbound_ubound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr(5) As Integer
        
        Console.WriteLine(LBound(arr))
        Console.WriteLine(UBound(arr))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "5"]);
}

#[test]
fn array_multidimensional_lbound_ubound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Explicit bounds 1 To 3 for dimension 1, and 0 To 5 for dimension 2
        ' VB supports non-zero lower bounds in declarations
        Dim grid(1 To 3, 0 To 5) As Integer
        
        ' Dimension is 1-based index in LBound/UBound
        Console.WriteLine(LBound(grid, 1))
        Console.WriteLine(UBound(grid, 1))
        
        Console.WriteLine(LBound(grid, 2))
        Console.WriteLine(UBound(grid, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "3", "0", "5"]);
}
