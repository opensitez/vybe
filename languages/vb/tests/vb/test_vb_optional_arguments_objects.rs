use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Optional Arguments with Objects
// ═══════════════════════════════════════════════════════════

#[test]
fn optional_arguments_objects() {
    let out = run_vb(
        r#"
Class Configuration
End Class

Module M
    ' Optional parameters of object type must be Nothing
    Sub Initialize(Optional config As Configuration = Nothing)
        If config Is Nothing Then
            Console.WriteLine("Default Config")
        Else
            Console.WriteLine("Custom Config")
        End If
    End Sub

    Sub Main()
        Initialize()
        Initialize(New Configuration())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Default Config", "Custom Config"]);
}
