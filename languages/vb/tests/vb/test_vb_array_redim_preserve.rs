use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Arrays (ReDim Preserve)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_redim_preserve() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3}
        
        ' ReDim Preserve resizes array and keeps existing elements
        ReDim Preserve arr(4)
        arr(3) = 4
        arr(4) = 5
        
        For i As Integer = 0 To UBound(arr)
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4", "5"]);
}

#[test]
fn array_redim_no_preserve() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3}
        
        ' ReDim without Preserve clears elements (initializes to default)
        ReDim arr(2)
        arr(0) = 9
        
        For i As Integer = 0 To UBound(arr)
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9", "0", "0"]);
}
