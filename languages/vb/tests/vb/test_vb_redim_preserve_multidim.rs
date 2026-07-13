use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ReDim Preserve Multi-Dimensional
// ═══════════════════════════════════════════════════════════

#[test]
fn redim_preserve_multidim() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Only the last dimension can be resized when using Preserve
        Dim arr(1, 1) As Integer
        arr(0, 0) = 1
        arr(0, 1) = 2
        arr(1, 0) = 3
        arr(1, 1) = 4
        
        ReDim Preserve arr(1, 2)
        arr(0, 2) = 5
        arr(1, 2) = 6
        
        Console.WriteLine(arr(0, 0))
        Console.WriteLine(arr(1, 1))
        Console.WriteLine(arr(1, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "4", "6"]);
}
