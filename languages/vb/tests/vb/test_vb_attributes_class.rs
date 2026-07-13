use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Attributes (Class Level)
// ═══════════════════════════════════════════════════════════

#[test]
fn attribute_class_basic() {
    let out = run_vb(
        r#"
<Serializable>
Class DataHolder
    Public Value As String = "Test"
End Class

Module M
    Sub Main()
        Dim t As Type = GetType(DataHolder)
        ' Check if attribute is applied
        Dim isSerializable As Boolean = t.IsSerializable
        Console.WriteLine(isSerializable)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn attribute_custom_class() {
    let out = run_vb(
        r#"
<AttributeUsage(AttributeTargets.Class)>
Class AuthorAttribute
    Inherits Attribute
    
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

<Author("Jane Doe")>
Class MyComponent
End Class

Module M
    Sub Main()
        Dim attr As AuthorAttribute = DirectCast(Attribute.GetCustomAttribute(GetType(MyComponent), GetType(AuthorAttribute)), AuthorAttribute)
        Console.WriteLine(attr.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Jane Doe"]);
}
