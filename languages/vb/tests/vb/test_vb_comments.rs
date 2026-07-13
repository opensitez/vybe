use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Comments (REM and Single Quote)
// ═══════════════════════════════════════════════════════════

#[test]
fn comments_rem_and_quote() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        REM This is a comment using REM keyword
        Console.WriteLine("Start")
        ' This is a standard single quote comment
        Dim x As Integer = 5 REM Inline REM comment
        Dim y As Integer = 10 ' Inline quote comment
        Console.WriteLine(x + y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start", "15"]);
}
