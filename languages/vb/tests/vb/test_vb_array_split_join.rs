use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Methods (Split and Join)
// ═══════════════════════════════════════════════════════════

#[test]
fn array_split_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "apple,banana,cherry"
        Dim parts As String() = Split(text, ",")
        
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn array_join_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr As String() = {"red", "green", "blue"}
        Dim result As String = Join(arr, "-")
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["red-green-blue"]);
}
