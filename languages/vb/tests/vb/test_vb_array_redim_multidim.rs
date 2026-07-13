use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Arrays (ReDim Multidimensional)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_redim_multidimensional() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim grid(1, 1) As Integer
        grid(0, 0) = 1
        grid(1, 1) = 2
        
        ' You can only ReDim Preserve the LAST dimension of a multidimensional array
        ReDim Preserve grid(1, 2)
        grid(1, 2) = 3
        
        Console.WriteLine(grid(0, 0))
        Console.WriteLine(grid(1, 1))
        Console.WriteLine(grid(1, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
