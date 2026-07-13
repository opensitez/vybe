use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Attributes (Method Level)
// ═══════════════════════════════════════════════════════════

#[test]
fn attribute_method_obsolete() {
    let out = run_vb(
        r#"
Class LegacyCode
    <Obsolete("Use NewMethod instead")>
    Public Sub OldMethod()
        Console.WriteLine("Old")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New LegacyCode()
        ' Should run even if obsolete
        l.OldMethod()
        Console.WriteLine("Done")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Old", "Done"]);
}
