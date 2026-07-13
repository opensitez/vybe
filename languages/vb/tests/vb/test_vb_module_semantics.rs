use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Module Semantics (Shared by default)
// ═══════════════════════════════════════════════════════════

#[test]
fn module_semantics() {
    let out = run_vb(
        r#"
Module GlobalConfig
    ' Modules are essentially NotInheritable classes with only Shared members
    Public Property AppName As String = "VybeApp"
    
    Public Sub PrintConfig()
        Console.WriteLine(AppName)
    End Sub
End Module

Module M
    Sub Main()
        ' Members of a module can be accessed without qualification
        PrintConfig()
        
        ' Or with qualification
        GlobalConfig.AppName = "VybeApp V2"
        GlobalConfig.PrintConfig()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["VybeApp", "VybeApp V2"]);
}
