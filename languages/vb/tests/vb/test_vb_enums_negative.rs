use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Enum with Negative Underlying Values
// ═══════════════════════════════════════════════════════════

#[test]
fn enums_negative() {
    let out = run_vb(
        r#"
Enum Status As Short
    Error = -1
    Pending = 0
    Active = 1
    Completed = 2
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Error
        Console.WriteLine(s)
        
        Dim val As Short = CShort(s)
        Console.WriteLine(val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Error", "-1"]);
}
