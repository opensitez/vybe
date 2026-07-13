use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interaction Builtins (IIf, If Operator)
// ═══════════════════════════════════════════════════════════

#[test]
fn interaction_iif_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 10
        ' IIf evaluates both true and false parts, but returns one
        Dim result As String = CStr(IIf(x > 5, "Greater", "Lesser"))
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Greater"]);
}

#[test]
fn interaction_if_operator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 3
        ' The If operator short-circuits
        Dim result As String = If(x > 5, "Greater", "Lesser")
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Lesser"]);
}

#[test]
fn interaction_if_operator_null_coalescing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s1 As String = Nothing
        Dim s2 As String = "Fallback"
        
        ' If operator with two arguments acts like null-coalescing
        Dim result As String = If(s1, s2)
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Fallback"]);
}
