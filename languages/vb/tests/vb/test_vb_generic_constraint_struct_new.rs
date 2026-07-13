use super::helpers::run_vb;

#[test]
fn generic_constraint_struct_new() {
    let out = run_vb(
        r#"
' As Structure requires T to be a value type
Class ValueCache(Of T As Structure)
    Public Property Item As T
End Class

' As New requires T to have a parameterless constructor
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class Person
    Public Property Name As String = "Bob"
End Class

Module M
    Sub Main()
        Dim vc As New ValueCache(Of Integer)()
        vc.Item = 42
        Console.WriteLine(vc.Item)
        
        Dim f As New Factory(Of Person)()
        Dim p = f.Create()
        Console.WriteLine(p.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Bob"]);
}
