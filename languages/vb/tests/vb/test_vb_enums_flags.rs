use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Enums (Flags Attribute)
// ═══════════════════════════════════════════════════════════

#[test]
fn enum_flags_attribute() {
    let out = run_vb(
        r#"
<Flags>
Enum FileAccess
    None = 0
    Read = 1
    Write = 2
    ReadWrite = Read Or Write
End Enum

Module M
    Sub Main()
        Dim access As FileAccess = FileAccess.ReadWrite
        Console.WriteLine(access.ToString())
        
        Dim singleAccess As FileAccess = FileAccess.Read
        Console.WriteLine(singleAccess.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["ReadWrite", "Read"]);
}
