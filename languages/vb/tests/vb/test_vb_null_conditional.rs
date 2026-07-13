use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Null-Conditional Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn null_conditional_operator() {
    let out = run_vb(
        r#"
Class Data
    Public Property Value As String
End Class

Module M
    Sub Main()
        Dim d As Data = Nothing
        
        ' Using ? before dot checks if d is nothing
        Dim len As Integer? = d?.Value?.Length
        
        Console.WriteLine(len.HasValue)
        
        d = New Data() With { .Value = "Test" }
        len = d?.Value?.Length
        Console.WriteLine(len.HasValue)
        Console.WriteLine(len.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "True", "4"]);
}
