use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Environment Properties & Info
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_environment_processor_count() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Environment.ProcessorCount > 0)
        Console.WriteLine(Environment.NewLine.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}
