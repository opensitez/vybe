use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ReadOnly Auto-Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn readonly_auto_properties() {
    let out = run_vb(
        r#"
Class Item
    ' ReadOnly auto-property can be initialized inline
    Public ReadOnly Property Id As Integer = 42
    Public ReadOnly Property Name As String
    
    Public Sub New(name As String)
        ' Or initialized in the constructor
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim i As New Item("Test")
        Console.WriteLine(i.Id)
        Console.WriteLine(i.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Test"]);
}
