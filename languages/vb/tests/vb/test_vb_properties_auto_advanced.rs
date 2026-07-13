use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Auto-Implemented Properties (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn auto_implemented_property_initializers() {
    let out = run_vb(
        r#"
Class Item
    ' Auto-property with initializer
    Public Property Name As String = "Unknown"
    ' Auto-property with object initializer syntax for collection
    Public Property Tags As New System.Collections.Generic.List(Of String) From {"New"}
End Class

Module M
    Sub Main()
        Dim i As New Item()
        Console.WriteLine(i.Name)
        Console.WriteLine(i.Tags(0))
        
        i.Name = "Known"
        Console.WriteLine(i.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Unknown", "New", "Known"]);
}
