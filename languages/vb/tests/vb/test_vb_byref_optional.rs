use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ByRef and Optional Interactions
// ═══════════════════════════════════════════════════════════

#[test]
fn optional_byref_parameter() {
    let out = run_vb(
        r#"
Module M
    ' ByRef parameter can be Optional with a default value
    Sub UpdateValue(Optional ByRef val As Integer = 5)
        val += 10
        Console.WriteLine("Inside: " & val.ToString())
    End Sub

    Sub Main()
        Dim x As Integer = 2
        ' Passing explicitly (mutates x)
        UpdateValue(x)
        Console.WriteLine("After: " & x.ToString())
        
        ' Omitting creates a temporary variable initialized to 5
        UpdateValue()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inside: 12", "After: 12", "Inside: 15"]);
}
