use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Anonymous Types (Key Properties)
// ═══════════════════════════════════════════════════════════

#[test]
fn anonymous_type_key_properties() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The 'Key' modifier makes the property read-only and includes it in Equals/GetHashCode
        Dim p1 = New With { Key .Id = 1, .Name = "Alice" }
        Dim p2 = New With { Key .Id = 1, .Name = "Bob" }
        Dim p3 = New With { Key .Id = 2, .Name = "Alice" }
        
        ' Equality depends only on Key properties
        Console.WriteLine(p1.Equals(p2))
        Console.WriteLine(p1.Equals(p3))
        
        Console.WriteLine(p1.Id)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False", "1"]);
}
