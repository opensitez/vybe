use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Enums (Bitwise Operations)
// ═══════════════════════════════════════════════════════════

#[test]
fn enum_bitwise_or() {
    let out = run_vb(
        r#"
Enum Status
    None = 0
    Active = 1
    Visible = 2
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Active Or Status.Visible
        Console.WriteLine(CInt(s))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enum_bitwise_and() {
    let out = run_vb(
        r#"
Enum Status
    None = 0
    Active = 1
    Visible = 2
    Both = 3
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Both And Status.Visible
        Console.WriteLine(CInt(s))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}
