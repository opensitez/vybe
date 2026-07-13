use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Property Default Values
// ═══════════════════════════════════════════════════════════

#[test]
fn property_default_values() {
    let out = run_vb(
        r#"
Class Configuration
    ' Default values for auto-implemented properties
    Public Property MaxItems As Integer = 100
    Public Property Description As String = "Default Settings"
    Public Property IsEnabled As Boolean = True
End Class

Module M
    Sub Main()
        Dim config As New Configuration()
        Console.WriteLine(config.MaxItems)
        Console.WriteLine(config.Description)
        Console.WriteLine(config.IsEnabled)
        
        config.MaxItems = 50
        Console.WriteLine(config.MaxItems)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "Default Settings", "True", "50"]);
}
