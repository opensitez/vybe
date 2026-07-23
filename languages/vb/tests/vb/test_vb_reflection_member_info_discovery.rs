use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection Type & MemberInfo Discovery
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_get_methods() {
    let src = r#"
Imports System.Reflection

Class TargetClass
    Public Sub Foo()
    End Sub
    Public Function Bar() As Integer
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(TargetClass)
        Dim methods As MethodInfo() = t.GetMethods(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)
        Console.WriteLine(methods.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_reflection_get_properties() {
    let src = r#"
Imports System.Reflection

Class Person
    Public Property Name As String
    Public Property Age As Integer
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(Person)
        Dim props = t.GetProperties()
        Console.WriteLine(props.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}
