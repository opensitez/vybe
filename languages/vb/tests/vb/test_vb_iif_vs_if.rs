use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: IIf vs If Operator (Short-circuiting)
// ═══════════════════════════════════════════════════════════

#[test]
fn iif_vs_if_short_circuit() {
    let out = run_vb(
        r#"
Module M
    Function LogAndReturn(val As Integer) As Integer
        Console.WriteLine("Evaluated " & val)
        Return val
    End Function

    Sub Main()
        Dim condition As Boolean = True
        
        Console.WriteLine("Testing IIf (evaluates both):")
        ' IIf is a function, so all arguments are evaluated before passing
        Dim result1 = IIf(condition, LogAndReturn(1), LogAndReturn(2))
        
        Console.WriteLine("Testing If (short-circuits):")
        ' If operator short-circuits
        Dim result2 = If(condition, LogAndReturn(3), LogAndReturn(4))
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec![
            "Testing IIf (evaluates both):",
            "Evaluated 1",
            "Evaluated 2",
            "Testing If (short-circuits):",
            "Evaluated 3"
        ]
    );
}
