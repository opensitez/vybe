use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Lambda Expressions Returning Arrays
// ═══════════════════════════════════════════════════════════

#[test]
fn lambda_arrays() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Lambda returning an array literal
        Dim getArray = Function() {1, 2, 3}
        
        Dim arr = getArray()
        For Each n In arr
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
