use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multi-line String Literals
// ═══════════════════════════════════════════════════════════

#[test]
fn string_multiline_literal() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' VB.NET allows strings to span multiple lines without concatenation
        Dim query As String = "SELECT *
FROM Users
WHERE Age > 18"
        
        Console.WriteLine(query.Contains("FROM"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn string_multiline_interpolated() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim table As String = "Orders"
        Dim query As String = $"SELECT *
FROM {table}
WHERE Total > 100"
        
        Console.WriteLine(query.Contains("Orders"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
