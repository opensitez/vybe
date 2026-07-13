use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Late Binding (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn late_binding_property_access() {
    let out = run_vb(
        r#"
Class Item
    Public Property Name As String
End Class

Module M
    Sub Main()
        ' With Option Strict Off (default), Object variables can use late binding
        Dim obj As Object = New Item() With { .Name = "TestItem" }
        
        ' Late-bound property access
        Console.WriteLine(obj.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["TestItem"]);
}
