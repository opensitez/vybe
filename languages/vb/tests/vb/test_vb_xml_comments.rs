use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Comments
// ═══════════════════════════════════════════════════════════

#[test]
fn xml_comments_parsing() {
    let out = run_vb(
        r#"
Module M
    ''' <summary>
    ''' This is a summary.
    ''' </summary>
    ''' <param name="msg">The message to print.</param>
    Sub Main()
        Console.WriteLine("Parsed OK")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed OK"]);
}
