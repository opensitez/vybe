use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Static Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn static_properties() {
    let out = run_vb(
        r#"
Class Cache
    ' Static properties maintain state across all instances
    Public Shared Property LastAccessed As String = "None"
    Public Shared ReadOnly Property CreatedAt As Date = #2024-01-01#
    
    Public Sub Access(item As String)
        LastAccessed = item
    End Sub
End Class

Module M
    Sub Main()
        Dim c1 As New Cache()
        c1.Access("Item1")
        
        Dim c2 As New Cache()
        Console.WriteLine(Cache.LastAccessed)
        
        c2.Access("Item2")
        Console.WriteLine(Cache.LastAccessed)
        
        Console.WriteLine(Cache.CreatedAt.Year)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Item1", "Item2", "2024"]);
}
