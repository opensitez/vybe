use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Partial Classes with Attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn partial_classes_attributes() {
    let out = run_vb(
        r#"
<Serializable>
Partial Class Person
    Public Property FirstName As String
End Class

<Obsolete("Use NewPerson instead")>
Partial Class Person
    Public Property LastName As String
End Class

Module M
    Sub Main()
        Dim p As New Person() With { .FirstName = "John", .LastName = "Doe" }
        
        Dim attrs = p.GetType().GetCustomAttributes(False)
        Console.WriteLine(attrs.Length)
        
        For Each attr In attrs
            Console.WriteLine(attr.GetType().Name)
        Next
    End Sub
End Module
"#,
    );
    // Note: The order of attributes returned by GetCustomAttributes is not guaranteed,
    // so we can't assert the exact order, but we can verify the count.
    assert_eq!(out.len(), 3);
    assert_eq!(out[0], "2");
    assert!(out.contains(&"SerializableAttribute".to_string()));
    assert!(out.contains(&"ObsoleteAttribute".to_string()));
}
