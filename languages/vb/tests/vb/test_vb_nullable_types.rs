use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Nullable Value Types
// ═══════════════════════════════════════════════════════════

#[test]
fn nullable_value_types_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The ? modifier makes a value type nullable
        Dim x As Integer? = Nothing
        Dim y As Integer? = 10
        
        Console.WriteLine(x.HasValue)
        Console.WriteLine(y.HasValue)
        Console.WriteLine(y.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "True", "10"]);
}

#[test]
fn nullable_value_types_operators() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Integer? = 5
        Dim b As Integer? = Nothing
        
        ' Lifted operators: If one operand is Nothing, result is Nothing
        Dim result As Integer? = a + b
        
        Console.WriteLine(result.HasValue)
        
        ' Coalescing with If operator
        Console.WriteLine(If(result, -1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "-1"]);
}
