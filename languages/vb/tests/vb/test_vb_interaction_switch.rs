use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interaction Builtins (Switch)
// ═══════════════════════════════════════════════════════════

#[test]
fn interaction_switch_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Integer = 15
        
        ' Evaluates pairs of expressions
        Dim result As String = CStr(Switch(
            value < 10, "Small",
            value >= 10 And value < 20, "Medium",
            value >= 20, "Large"
        ))
        
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Medium"]);
}
