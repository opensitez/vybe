use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Anonymous Types
// ═══════════════════════════════════════════════════════════

#[test]
fn anonymous_type_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Creates an instance of an anonymous type
        Dim person = New With { .Name = "Alice", .Age = 30 }
        
        Console.WriteLine(person.Name)
        Console.WriteLine(person.Age)
        
        ' In VB, properties of anonymous types without 'Key' modifier are mutable
        person.Age = 31
        Console.WriteLine(person.Age)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30", "31"]);
}
