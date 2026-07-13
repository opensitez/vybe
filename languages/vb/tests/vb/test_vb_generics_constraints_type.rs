use super::helpers::run_vb;

#[test]
fn generics_constraints_class_struct() {
    let out = run_vb(
        r#"
' Constraint: T must be a reference type
Class RefContainer(Of T As Class)
    Public Item As T
End Class

' Constraint: T must be a value type
Class ValContainer(Of T As Structure)
    Public Item As T
End Class

Module M
    Sub Main()
        Dim r As New RefContainer(Of String)()
        r.Item = "Hello"
        Console.WriteLine(r.Item)
        
        Dim v As New ValContainer(Of Integer)()
        v.Item = 42
        Console.WriteLine(v.Item)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello", "42"]);
}

#[test]
fn generics_constraints_new() {
    let out = run_vb(
        r#"
' Constraint: T must have a parameterless constructor
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class MyClassWithConstructor
    Public ReadOnly Value As String = "Constructed"
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of MyClassWithConstructor)()
        Dim obj = f.Create()
        Console.WriteLine(obj.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Constructed"]);
}
