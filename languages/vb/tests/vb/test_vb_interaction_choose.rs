use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interaction Builtins (Choose)
// ═══════════════════════════════════════════════════════════

#[test]
fn interaction_choose_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Choose is 1-based index
        Dim choice As String = CStr(Choose(2, "Apple", "Banana", "Cherry"))
        Console.WriteLine(choice)
        
        ' Out of bounds returns Nothing (null)
        Dim invalidChoice As Object = Choose(4, "Apple", "Banana", "Cherry")
        Console.WriteLine(IsNothing(invalidChoice))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Banana", "True"]);
}
