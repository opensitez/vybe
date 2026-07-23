use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection PropertyInfo GetValue & SetValue
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_get_set_property_value() {
    let src = r#"
Imports System.Reflection

Class Configuration
    Public Property ServerHost As String = "127.0.0.1"
End Class

Module Program
    Sub Main()
        Dim cfg As New Configuration()
        Dim t As Type = cfg.GetType()
        Dim prop As PropertyInfo = t.GetProperty("ServerHost")

        Console.WriteLine(prop.GetValue(cfg))
        prop.SetValue(cfg, "192.168.1.1")
        Console.WriteLine(cfg.ServerHost)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["127.0.0.1", "192.168.1.1"]);
}
