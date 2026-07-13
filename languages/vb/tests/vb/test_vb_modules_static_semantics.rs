use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Modules (Static Semantics)
// ═══════════════════════════════════════════════════════════

#[test]
fn module_implicit_static() {
    let out = run_vb(
        r#"
Module Globals
    ' Module members are implicitly shared/static
    Public Counter As Integer = 0
    
    Public Sub Increment()
        Counter += 1
    End Sub
End Module

Module M
    Sub Main()
        ' Can be accessed without qualification (globally available in namespace)
        Increment()
        Increment()
        
        ' Or with qualification
        Globals.Increment()
        
        Console.WriteLine(Counter)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}
